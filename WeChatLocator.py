class WeChatPackageLocator:
    appPackage = "com.tencent.mm"  # 微信包名
    appActivity = ".ui.LauncherUI"  # 启动页面类名


class WeChatClassLocator:
    """根据类名定位"""
    TEXT_VIEW = 'android.widget.TextView'  # TextView控件


class WeChatXPathLocator:
    """根据XPath定位"""
    PERSON_COLUMN = ('//androidx.recyclerview.widget.RecyclerView//android.widget.LinearLayout[2]'
                     '//android.widget.LinearLayout//android.widget.RelativeLayout')
    # 个人信息区域，根据微信号定位父元素，也就是个人信息区域
    PERSON_INFO_AREA = '//android.widget.TextView[starts-with(@text, "微信号")]/..'
    SIGNATURE = '//android.widget.TextView[starts-with(@text, "个性签名")]/../android.widget.TextView[2]'
    MYSELF_NAME = '//android.widget.LinearLayout/android.view.View'


class WeChatTextLocator:
    """根据文本内容定位"""
    ADDRESS_BOOK = '通讯录'  # 通讯录
    MYSELF = '我'  # 我


class WeChatAccessibilityIDLocator:
    """根据AccessibilityID定位"""
    PLUS_BUTTON = '更多功能按钮，已折叠'
    CLOSE_BUTTON = '关闭'
    RETURN_BUTTON = '返回'
    HEAD_BUTTON = '头像'
