class globalVar():
    APP_TITLE = u'音符文档助手'
    STATUS = (u'注意：\n1、本软件由音符基于python开发;\n2、本软件不会获取您的个人信息;\n'
                   '3、本软件目前有PDF拆分、word转PDF及PDF转word功能;\n4、本软件可能存在bug'
                   '当出现卡顿时请不要惊讶，告诉作者即可;\n5、软件正在进一步维护中，感谢使用;\n'
                   '6、禁止一切违法利用行为，软件最终解释权归作者所有.')
    APP_ICON = 'res/text.ico'

    # 文件列表
    fileList = []

    # 当前操作状态
    textStatus = None
    status = "word转pdf"

    def __init__(self):
        self.APP_TITLE = u'音符文档助手'
        self.STATUS = (u'注意：\n1、本软件由音符基于python开发;\n2、本软件不会获取您的个人信息;\n'
                  '3、本软件目前有PDF拆分、word转PDF及PDF转word功能;\n4、本软件可能存在bug'
                  '当出现卡顿时请不要惊讶，告诉作者即可;\n5、软件正在进一步维护中，感谢使用;\n'
                  '6、禁止一切违法利用行为，软件最终解释权归作者所有.')
        self.APP_ICON = 'res/text.ico'
