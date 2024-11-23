# 因微信Appium的元素定位发生改变，导致原有的定位方式失效，所以重新定位
import time

from selenium.common import NoSuchElementException
from selenium.webdriver.support.wait import WebDriverWait

from WeChatClass import WeChat


# 计时装饰器
def time_decorator(func):
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        elapsed_time = end_time - start_time
        # 计算分钟和小时（如果需要）
        minutes, seconds = divmod(elapsed_time, 60)
        hours, minutes = divmod(minutes, 60)

        # 根据时间长度选择合适的时间格式进行打印
        if hours:
            print(f"{func.__name__}执行了{hours}小时{minutes:.0f}分钟{seconds:.2f}秒")
        elif minutes:
            print(f"{func.__name__}执行了{minutes:.0f}分钟{seconds:.2f}秒")
        else:
            print(f"{func.__name__}执行了{elapsed_time:.2f}秒")
        return result

    return wrapper


@time_decorator
def main():
    wechat = WeChat()
    data_list = []  # 用于存储联系人信息
    try:
        person_resource_id = wechat.get_person_resource_id(wechat.get_person_column())  # 获取联系人的resource-id
        all_person = wechat.driver.find_elements('id', person_resource_id)  # 获取当前页面的所有联系人
        pre_person = all_person[0]  # 获取第一个联系人
        next_person = all_person[1]  # 获取第二个联系人
        all_person_name_list = []  # 用于存储所有联系人姓名
        flag = True  # 标志位，用于判断是否滑动到底部
        index = 1
        while True:
            # 1.得到当前第一个联系人姓名
            current_person_name = wechat.get_person_name(pre_person)
            if current_person_name not in all_person_name_list:
                all_person_name_list.append(current_person_name)  # 将当前联系人姓名添加到列表中
                # 判断人员姓名不等于'微信团队', '文件传输助手', 本人
                if current_person_name not in ['微信团队', '文件传输助手', wechat.myself]:
                    try:
                        pre_person.click()  # 点击进入person详情界面
                        info_dict = wechat.get_person_info(index, is_friend=True)  # 获取人员信息
                        data_list.append(info_dict)
                        wechat.driver.back()  # 返回至通讯录
                        # TODO: 人员重名
                    except Exception as e:
                        print(f'获取人员【{current_person_name}】信息失败: {e}')
                    finally:
                        index += 1
            # 结束条件
            if not flag:
                break
            # 2.滑动联系人到屏幕0.7位置
            wechat.driver.swipe(500, pre_person.rect['y'] + pre_person.rect['height'], 500,
                                wechat.driver.get_window_size()['height'] * 0.7, 500)

            # 3.刷新页面，获取当前页面的所有联系人
            all_person = WebDriverWait(wechat.driver, 10).until(lambda x: x.find_elements('id', person_resource_id))
            # 4.获取第二个联系人姓名
            next_person_name = wechat.get_person_name(next_person)
            # 5.滑动后，将第二个联系人设置为第一个联系人
            for n, person in enumerate(all_person):
                try:  # 滑动后，找人名，可能滑过，导致这个页面第一个人没有“text”元素
                    # 得到当前人员的姓名
                    wechat.driver.implicitly_wait(1)    # 因设置了隐式等待为5，需等很久，所以现在将隐式等待改为1s
                    current_person_name = wechat.get_person_name(person)
                    wechat.driver.implicitly_wait(5)    # 将隐式等待改为5s，正常的情况
                except NoSuchElementException:
                    continue
                if current_person_name == next_person_name:
                    pre_person = person
                    # 最后一人时，all_person列表不能加1，会返回IndexError，正好作为结束标志
                    try:
                        next_person = all_person[n + 1]
                    except IndexError:
                        next_person = pre_person
                        flag = False
                    break
    except KeyboardInterrupt:
        print('停止运行')
    except Exception as e:
        print(e)
    finally:
        wechat.quit()
        wechat.save_to_excel(data_list)


if __name__ == '__main__':
    main()
