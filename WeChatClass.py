import os
import datetime

from appium.options.android import UiAutomator2Options
from appium import webdriver
from appium.webdriver.common.appiumby import AppiumBy as By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import xlwings as xw

from WeChatLocator import *


class WeChat:
    def __init__(self):
        # 初始化Android驱动，并设置隐式等待时间为10秒
        self.driver = self.android_driver()
        print('初始化Android驱动成功！')
        self.driver.implicitly_wait(5)
        self.myself = self.get_myself_name()

    def android_driver(self):
        desired_caps = dict()
        desired_caps['automationName'] = 'Uiautomator2'
        desired_caps['platformName'] = 'Android'
        # Windows执行cmd命令，查看版本号
        r = os.popen('adb shell getprop ro.build.version.release')
        desired_caps['platformVersion'] = r.read().replace('\n', '')
        # adb查看deviceName
        r = os.popen('adb devices -l')
        device_name = r.read().split('model:')[1].split('device:')[0]
        desired_caps['deviceName'] = device_name
        # desired_caps['appPackage'] = 'com.tencent.mm'
        # desired_caps['appActivity'] = 'ui.LauncherUI'
        # ！！！特别重要，否则会重新格式化微信
        desired_caps['noReset'] = True
        desired_caps['newCommandTimeout'] = 1800
        # Appium-Python-Client 3以后版本需要设置options
        options = UiAutomator2Options()
        options.load_capabilities(desired_caps)
        self.driver = webdriver.Remote(f'http://localhost:4723', options=options)
        return self.driver

    def find_element_by_text(self, text):
        """通过精确匹配文本内容来定位元素"""
        return self.driver.find_element(By.ANDROID_UIAUTOMATOR, rf'new UiSelector().text("{text}")')

    def find_element_by_text_starts_with(self, text):
        """通过匹配文本开头的内容来定位元素"""
        return self.driver.find_element(By.ANDROID_UIAUTOMATOR, rf'new UiSelector().textStartsWith("{text}")')

    def find_element_by_text_contains(self, text):
        """通过模糊匹配文本内容来定位元素"""
        return self.driver.find_element(By.ANDROID_UIAUTOMATOR, rf'new UiSelector().textContains("{text}")')

    def find_element_by_text_matches(self, text):
        """使用正则表达式来匹配文本内容"""
        return self.driver.find_element(By.ANDROID_UIAUTOMATOR, rf'new UiSelector().textMatches("{text}")')

    def get_myself_name(self):
        """获取本人昵称"""
        self.find_element_by_text(WeChatTextLocator.MYSELF).click()  # 进入‘我’
        myself = self.driver.find_element(By.XPATH, WeChatXPathLocator.MYSELF_NAME).text  # 获取本人昵称
        self.find_element_by_text(WeChatTextLocator.ADDRESS_BOOK).click()  # 返回通讯录
        return myself

    def get_person_column(self):
        """获取通讯录人员栏"""
        # 因联系人的resource id会变动，所有要根据通讯录列表的类名找到联系人的resource id
        # 找到通讯录列表的每一项，前面会有群聊、标签等无关的元素，所以需要过滤掉
        person_column = self.driver.find_element(By.XPATH, WeChatXPathLocator.PERSON_COLUMN)
        return person_column

    @staticmethod
    def get_person_resource_id(person_column):
        """获取通讯录人员的resource id"""
        person_resource_id = person_column.get_attribute('resource-id')
        return person_resource_id

    @staticmethod
    def get_person_name(person_column):
        """获取通讯录人员的名称"""
        person_name = person_column.find_element(By.CLASS_NAME, WeChatClassLocator.TEXT_VIEW).text
        return person_name

    def is_friend(self):
        """判断是否为好友"""
        self.find_element_by_text('发消息').click()  # 点击发消息按钮
        self.driver.find_element(By.ACCESSIBILITY_ID, WeChatAccessibilityIDLocator.PLUS_BUTTON).click()  # 点击加号按钮
        self.find_element_by_text('转账').click()  # 点击转账按钮
        for text in ['1', '转账']:  # 依次点击['1', '转账']按钮
            self.find_element_by_text(text).click()
        flag = True
        try:
            # 尝试点击‘关闭’按钮，如果出现则来到转账页面，说明是好友
            element = WebDriverWait(self.driver, 3).until(EC.presence_of_element_located(
                (By.ACCESSIBILITY_ID, WeChatAccessibilityIDLocator.CLOSE_BUTTON)))
            element.click()
        except TimeoutException:
            # 点击‘知道了’按钮
            self.find_element_by_text_contains('知道了').click()
            flag = False
        # 返回至‘个人信息页’
        while True:
            # 如果‘发消息’不在当前页面，则返回
            if '发消息' not in self.driver.page_source:
                self.driver.back()
            else:
                break
        return flag

    def get_person_info(self, index: int, is_friend: bool = True):
        """获取通讯录人员信息
        :argument
            - is_friend: 开关标志位，默认打开，查询是否为好友
        """
        try:
            # 备注，昵称，微信号，地区，个性签名
            person_info_area = self.driver.find_element(By.XPATH, WeChatXPathLocator.PERSON_INFO_AREA)
            info_dict = {'序号': index, '备注': '', '昵称': '', '微信号': '', '地区': '', '个性签名': '',
                         '对方是否是好友': '',
                         '日期': datetime.datetime.now().strftime("%Y_%m_%d %H:%M:%S"),
                         '头像': ''}
            info_list = [info.text for info in person_info_area.find_elements(By.CLASS_NAME, WeChatClassLocator.TEXT_VIEW)]
            # 根据 info_list 的长度和内容填充剩余信息
            if len(info_list) == 4:
                info_dict['备注'], info_dict['昵称'], info_dict['微信号'], info_dict['地区'] = info_list
            elif len(info_list) == 3:
                if '微信号' in info_list[1]:
                    info_dict['昵称'], info_dict['微信号'], info_dict['地区'] = info_list
                else:
                    info_dict['备注'], info_dict['昵称'], info_dict['微信号'] = info_list
            elif len(info_list) == 2:
                info_dict['昵称'], info_dict['微信号'] = info_list
            # 格式化信息，去除多余的前缀
            info_dict['昵称'] = info_dict['昵称'].lstrip('昵称:  ') if '昵称' in info_dict['昵称'] else info_dict['昵称']
            info_dict['微信号'] = info_dict['微信号'].lstrip('微信号:  ') if '微信号' in info_dict['微信号'] else info_dict[
                '微信号']
            info_dict['地区'] = info_dict['地区'].lstrip('地区:  ') if '地区' in info_dict['地区'] else info_dict['地区']

            # 头像保存的路径
            head_dir = 'head_' + str(datetime.datetime.now().strftime("%Y_%m_%d"))
            if not os.path.exists(head_dir):
                os.mkdir(head_dir)
            head_name = info_dict['备注'] if info_dict['备注'] else info_dict['昵称']
            head_path = os.path.join(head_dir, head_name + '.png')
            # 对头像截图，保存在head_dir文件夹下
            head = self.driver.find_element(
                By.ACCESSIBILITY_ID, WeChatAccessibilityIDLocator.HEAD_BUTTON).screenshot(head_path)
            if head:
                # 如果头像截图成功，则添加头像路径到info_dict
                info_dict['头像'] = head_path
            self.find_element_by_text('更多信息').click()
            if '个性签名' in self.driver.page_source:
                info_dict['个性签名'] = self.driver.find_element(By.XPATH, WeChatXPathLocator.SIGNATURE).text
            # 点击‘返回’按钮
            self.driver.back()
            if is_friend:
                # 查看是否为好友
                is_friend = '是' if self.is_friend() else '否'
                info_dict['对方是否是好友'] = is_friend
            else:
                info_dict['对方是否是好友'] = '未知'

            print(info_dict)
            return info_dict
        except StaleElementReferenceException:  # 不知道为什么为StaleElementReferenceException，重新执行该函数
            self.get_person_info(index, is_friend)

    @staticmethod
    def save_to_excel(data_list):
        """保存数据到Excel"""
        excel_path = '微信好友.xlsx'
        if os.path.isfile(excel_path):
            # visible False，不前台打开Excel
            app = xw.App(visible=False, add_book=False)
            # 打开Excel
            wb = app.books.open(excel_path)
            # 新增工作表
            sht = wb.sheets.add()
        else:
            # 创建一个新的工作簿，不显示在前台
            app = xw.App(visible=False)
            # 打开第一个工作簿的第一个工作表
            wb = app.books[0]
            sheet = wb.sheets[0]
            # 保存工作簿
            wb.save(excel_path)
            # 关闭工作簿
            wb.close()
            # 关闭应用程序
            app.quit()
            # 重新打开工作表，不然添加图片会报错
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(excel_path)
            sht = wb.sheets[0]

        # 写入表头
        titles = [['序号', '备注', '昵称', '微信号', '地区', '个性签名', '对方是否是好友', '日期', '头像']]
        sht.range('a1').value = titles
        data_info = []
        head_path_list = []
        # 拆分数据列表中，单独用一个列表存放头像
        for info_dict in data_list:
            data = [info_dict.get(key) for key in titles[0] if key != '头像']
            data_info.append(data)
            head_path_list.append(info_dict.get('头像'))
        # 写入数据
        sht.range('a2').value = data_info
        # 自适应单元格大小
        rng = sht.range('a1').expand('table')
        rng.autofit()
        # 向头像列插入头像
        for num, hp in enumerate(head_path_list):
            picture_range = sht.range(f'i{num + 2}')
            picture = sht.pictures.add(hp, left=picture_range.left, top=picture_range.top, width=100, height=100)
        # 改变头像列宽行高
        n_rows = rng.rows.count
        cell = sht.range(f'i2:i{n_rows}')
        cell.column_width = 20
        cell.row_height = 100
        # 保存文件
        wb.save(excel_path)
        wb.close()

    def quit(self):
        self.driver.quit()
