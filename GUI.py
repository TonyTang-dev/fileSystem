# -*- coding: utf-8 -*-

import wx
import win32api
import sys
import os
import wx.lib.agw.aui as aui
from wx.adv import Animation, AnimationCtrl
import glob
import fitz
import time

# 模块
from word2pdf import doc2pdf
from globalVar import globalVar


filePath_w2p = ""


class FileDrop(wx.FileDropTarget):
    def __init__(self, panel, statusText):
        wx.FileDropTarget.__init__(self)
        self.text = statusText
        self.panel = panel

    def OnDropFiles(self, x, y, filePath):
        globalVar.fileList = filePath

        path = ""
        for i in filePath:
            path = path+"\n"+i

        self.text.Label = "="*10 + "选中文件路径"+"="*10 + path

        return True


class mainFrame(wx.Frame):
    '''程序主窗口类，继承自wx.Frame'''

    id_open = 1
    id_help = 2

    id_word2pdf = 3
    id_pdf2word = 4
    id_mergePdf = 5
    id_cutPdf = 6
    id_img2pdf = 7
    id_pdf2img = 8
    id_author = 9

    print(id_help,id_author)

    fileName = ""

    def __init__(self, parent):
        '''构造函数'''

        wx.Frame.__init__(self, parent, -1, globalVar.APP_TITLE)
        self.SetBackgroundColour(wx.Colour(224, 224, 224))
        self.SetSize((620, 400))
        self.SetMaxSize((620, 400))
        self.Center()

        if hasattr(sys, "frozen") and getattr(sys, "frozen") == "windows_exe":
            exeName = win32api.GetModuleFileName(win32api.GetModuleHandle(None))
            icon = wx.Icon(exeName, wx.BITMAP_TYPE_ICO)
        else:
            icon = wx.Icon(globalVar.APP_ICON, wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)

        self.tb1 = self._CreateToolBar('F')
        self.tb2 = self._CreateToolBar()
        # self.tbv = self._CreateToolBar('V')

        p_left = wx.Panel(self, -1)
        p_center0 = wx.Panel(self, -1)

        image_file = 'res/3.jpg'
        to_bmp_image = wx.Image(image_file, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        self.bitmap = wx.StaticBitmap(p_center0, -1, to_bmp_image, (0, 0), (400, 200))

        p_center1 = wx.Panel(self, -1)
        p_bottom = wx.Panel(self, -1)
        p_center0.SetBackgroundColour("White")
        p_bottom.SetBackgroundColour("White")
        image_file = 'res/addfile.png'
        to_bmp_image = wx.Image(image_file, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        self.bitmap = wx.StaticBitmap(p_bottom, -1, to_bmp_image, (0, 0), (400, 140))

        statusText0 = wx.StaticText(p_left, -1, globalVar.STATUS, pos=(0, 10), size=(200, 200), style=wx.ALIGN_LEFT)
        # statusText1 = wx.StaticText(p_left, -1, u"当前状态：word转PDF", pos=(15, 30), size=(200, -1), style=wx.ALIGN_CENTER)
        # statusText2 = wx.StaticText(p_left, -1, u"当前状态：PDF转word", pos=(15, 30), size=(200, -1), style=wx.ALIGN_CENTER)
        # statusText3 = wx.StaticText(p_left, -1, u"当前状态：其他", pos=(15, 30), size=(200, -1), style=wx.ALIGN_CENTER)

        # statusText0 = wx.StaticText(p_bottom, -1, "将文件拖曳到此开始实现文件转PDF", pos=(5, 10), size=(200, 200),
        #                             style=wx.ALIGN_LEFT)

        filepathText0 = wx.StaticText(p_center0, -1, "", pos=(0, 21), size=(500, -1), style=wx.ALIGN_LEFT)
        # 文件拖曳
        fileDrop = FileDrop(p_center0, filepathText0)
        p_bottom.SetDropTarget(fileDrop)

        btn = wx.Button(p_left, -1, u'开始输出', pos=(30, 230), size=(100, -1))
        btn.Bind(wx.EVT_BUTTON, lambda e, mark=filepathText0, panel=p_center0: self.OnSwitch(e, mark, panel))
        btn.SetBackgroundColour('white')

        text0 = wx.StaticText(p_center0, -1, u'当前操作：' + globalVar.status, pos=(0, 0), size=(400, 20), style=wx.ALIGN_CENTER)
        globalVar.textStatus = text0
        text0.SetFont(wx.Font(10, wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        # text0.SetBackgroundColour("red")
        text1 = wx.StaticText(p_center1, -1, u'我是第2页', pos=(40, 100), size=(200, -1), style=wx.ALIGN_LEFT)
        line = wx.StaticText(p_center0, -1, u'', pos=(0, 20), size=(400, 1), style=wx.ALIGN_CENTER)
        line.SetBackgroundColour("black")

        self._mgr = aui.AuiManager()
        self._mgr.SetManagedWindow(self)

        self._mgr.AddPane(self.tb1,
                          aui.AuiPaneInfo().Name("ToolBar1").Caption(u"工具条").ToolbarPane().Top().Row(0).Position(
                              0).Floatable(False)
                          )
        self._mgr.AddPane(self.tb2,
                          aui.AuiPaneInfo().Name("ToolBar2").Caption(u"工具条").ToolbarPane().Top().Row(0).Position(
                              1).Floatable(True)
                          )
        # self._mgr.AddPane(self.tbv, 
        # aui.AuiPaneInfo().Name("ToolBarV").Caption(u"工具条").ToolbarPane().Right().Floatable(True)
        # )

        self._mgr.AddPane(p_left,
                          aui.AuiPaneInfo().Name("LeftPanel").Left().Layer(1).MinSize((200, -1)).Caption(
                              u"操作区").MinimizeButton(True).MaximizeButton(True).CloseButton(False)
                          )

        self._mgr.AddPane(p_center0,
                          aui.AuiPaneInfo().Name("CenterPanel0").CenterPane().Show()
                          )

        self._mgr.AddPane(p_center1,
                          aui.AuiPaneInfo().Name("CenterPanel1").CenterPane().Hide()
                          )

        self._mgr.AddPane(p_bottom,
                          aui.AuiPaneInfo().Name("BottomPanel").Bottom().MinSize((-1, 100)).Caption(
                              u"消息区").CaptionVisible(False).Resizable(True)
                          )

        self._mgr.Update()


    def _CreateToolBar(self, d='H'):
        '''创建工具栏'''

        bmp_open = wx.Bitmap('res/1.png', wx.BITMAP_TYPE_ANY)
        bmp_save = wx.Bitmap('res/1.png', wx.BITMAP_TYPE_ANY)
        bmp_help = wx.Bitmap('res/1.png', wx.BITMAP_TYPE_ANY)
        bmp_about = wx.Bitmap('res/1.png', wx.BITMAP_TYPE_ANY)

        if d.upper() in ['V', 'VERTICAL']:
            tb = aui.AuiToolBar(self, -1, wx.DefaultPosition, wx.DefaultSize,
                                agwStyle=aui.AUI_TB_TEXT | aui.AUI_TB_VERTICAL)
        else:
            tb = aui.AuiToolBar(self, -1, wx.DefaultPosition, wx.DefaultSize, agwStyle=aui.AUI_TB_TEXT)
        tb.SetToolBitmapSize(wx.Size(16, 16))

        if d.upper() != 'F':
            tb.AddSimpleTool(self.id_mergePdf, u'PDF合并', bmp_open, u'合并多个PDF为一个PDF')
            tb.AddSimpleTool(self.id_cutPdf, u'PDF拆分', bmp_open, u'将一个PDF拆分成多个')
            tb.AddSimpleTool(self.id_img2pdf, u'图片转PDF', bmp_save, u'将图片放到PDF文件中')
            tb.AddSeparator()
            tb.AddSimpleTool(self.id_pdf2img, u'PDF转图片', bmp_help, u'PDF每页转成一张图片')
            tb.AddSimpleTool(self.id_author, u'作者', bmp_about, u'关于作者')
            tb.Bind(wx.EVT_TOOL, self.dealFunction)
        else:
            tb.AddSimpleTool(self.id_open, u'File', bmp_open, u'打开文件')
            tb.AddSimpleTool(self.id_help, u'教程', bmp_help, u'使用教程')
            tb.AddSeparator()
            tb.AddSimpleTool(self.id_word2pdf, u'word转PDF', bmp_save, u'word文件转PDF文件')
            tb.AddSimpleTool(self.id_pdf2word, u'PDF转word', bmp_about, u'将PDF文件转为word文件')

            tb.Bind(wx.EVT_TOOL, self.dealFunction)
        tb.Realize()
        return tb

    def OnSwitch(self, evt, filepathText0, panel):
        path = filepathText0.Label.split("\n")[1]  # 获得文件路径

        if path == "":
            d = wx.MessageDialog(None, u"请先选择文件再点击转换哦", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal() == wx.ID_OK:
                pass
            d.Destroy()
            return
        statusText0 = wx.StaticText(panel, -1, "转换中···", pos=(80, 72), size=(72, -1), style=wx.ALIGN_CENTER)
        statusText0.SetBackgroundColour("White")
        animation = AnimationCtrl(panel, -1, Animation('res/5.gif'), pos=(80, 90))  # 创建一个动画
        animation.Play()  # 播放动图
        mainFrame(None).word2PDF(path)  # 开始转换
        animation.Stop()
        statusText0.Show(0)
        animation.Destroy()


    def word2pdf(self):
        if len(globalVar.fileList) == 0:
            d = wx.MessageDialog(None, u"请先选中文件再进行操作哦！", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal() == wx.ID_OK:
                pass
            d.Destroy()
            return

        flag = 1

        result = 1
        if flag == 1:
            result = doc2pdf(globalVar[0])
        if result == 1:
            d = wx.MessageDialog(None, u"转换失败，可能已经存在该PDF文件", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal() == wx.ID_OK:
                pass
            d.Destroy()
        else:
            pass
            # d=wx.MessageDialog(None, u"转换成功", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            # if d.ShowModal()==wx.ID_OK:
            # pass
            # d.Destroy()
        return

    def pdf2word(self):
        if len(globalVar.fileList) == 0:
            d = wx.MessageDialog(None, u"请先选中文件再进行操作哦！", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal() == wx.ID_OK:
                pass
            d.Destroy()
            return
        for i in globalVar.fileList:
            if i[-3:] != "pdf":
                d = wx.MessageDialog(None, u"文件"+i+"不是pdf文件，不能转换哦！", u"提示", wx.YES_NO | wx.ICON_QUESTION)
                if d.ShowModal() == wx.ID_OK:
                    pass
                d.Destroy()
                continue

            doc = fitz.open(i)
            docName = i[0:-4]+".docx"
            resultDoc = open(docName, "wb")
            for page in doc:
                text = page.getText().encode("utf8")
                resultDoc.write(text)
            resultDoc.close()
        return


    def mergePdf(self):
        if len(globalVar.fileList) == 0:
            d = wx.MessageDialog(None, u"请先选中文件再进行操作哦！", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal() == wx.ID_OK:
                pass
            d.Destroy()
            return
        docName = "YF操作PDF文件"
        times = 0
        while os.path.exists(docName + "(合并).pdf") and times < 8:
            docName += "1"
            times += 1

        if times == 8:
            d = wx.MessageDialog(None, u"当前目录存在多个相似文件，请先移除此类文件再操作", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal() == wx.ID_OK:
                pass
            d.Destroy()
            return

        resultDoc = fitz.open()
        for i in globalVar.fileList:
            if i[-3:] != "pdf":
                d = wx.MessageDialog(None, u"文件" + i + "不是pdf文件，不能合并哦！", u"提示", wx.YES_NO | wx.ICON_QUESTION)
                if d.ShowModal() == wx.ID_OK:
                    pass
                d.Destroy()
                continue
            doc = fitz.open(i)
            resultDoc.insertPDF(doc)
        resultDoc.save(docName + "(合并).pdf")

    def cutPdf(self):
        if len(globalVar.fileList) == 0:
            d = wx.MessageDialog(None, u"请先选中文件再进行操作哦！", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal() == wx.ID_OK:
                pass
            d.Destroy()
            return
        dlg = wx.TextEntryDialog(None, u"请按 'a-b' 的格式输入切割的起始页和结束页,如：1-23\n默认只对选中的第一个文件执行操作", u"输入提示", u"1-23")
        if dlg.ShowModal() == wx.ID_OK:
            message = dlg.GetValue()  # 获取文本框中输入的值

        dlg.Destroy()

    def img2pdf(self):
        pass
        # Doc = doc[0:-4]
        # doc = fitz.open(doc)
        # DOC = fitz.open()
        # DOC.insertPDF(doc, from_page=from_pa - 1, to_page=to_pa - 1)
        # DOC.save(Doc + "(拆分).pdf")
        # # print("操作完成，文件以保存在:"+Doc+"(拆分).pdf")
        # s1 = "\n操作完成，文件以保存在:\n" + Doc + "(拆分).pdf"

    def pdf2img(self):
        print("pdf2img")
        pass

    def openAuthor(self):
        print("author")
        pass

    def openFile(self):
        # 打开开文件对话框
        dlg = wx.FileDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            print(dlg.GetPath())  # 文件夹路径
            filePath_w2p = dlg.GetPath()
            self.word2PDF(filePath_w2p)

        dlg.Destroy()
        return

    def openHelp(self):
        print("openHelp")
        pass

    def dealFunction(self, event):
        index = event.GetId()

        # id_open = 1 id_help = 2  id_word2pdf = 3 id_pdf2word = 4
        # id_mergePdf = 5 id_cutPdf = 6 id_img2pdf = 7 id_pdf2img = 8 id_author = 9
        if index == 1:
            globalVar.status = "打开本地文件"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.openFile()
        elif index == 2:
            globalVar.status = "使用教程"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.openHelp()
        elif index == 3:
            globalVar.status = "word转pdf"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.word2pdf()
        elif index == 4:
            globalVar.status = "pdf转word"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.pdf2word()
        elif index == 5:
            globalVar.status = "合并pdf"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.mergePdf()
        elif index == 6:
            globalVar.status = "打开本地文件"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.cutPdf()
        elif index == 7:
            globalVar.status = "图片转pdf"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.img2pdf()
        elif index == 8:
            globalVar.status = "pdf转图片"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.pdf2img()
        elif index == 9:
            globalVar.status = "关于作者"
            globalVar.textStatus.SetLabel("当前操作："+globalVar.status)
            self.openAuthor()


class mainApp(wx.App):
    def OnInit(self):
        self.SetAppName(globalVar.APP_TITLE)
        self.Frame = mainFrame(None)
        self.Frame.Show()
        return True


if __name__ == "__main__":
    app = mainApp()
    app.MainLoop()
