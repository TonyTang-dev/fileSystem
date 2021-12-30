#-*- coding: utf-8 -*-

import wx
import win32api
import sys, os
import wx.lib.agw.aui as aui
from wx.adv import Animation, AnimationCtrl
import time

#模块
from word2pdf import doc2pdf

APP_TITLE = u'音符文档助手'
STATUS=(u'注意：\n1、本软件由音符基于python开发;\n2、本软件不会获取您的个人信息;\n'
            '3、本软件目前有PDF拆分、word转PDF及PDF转word功能;\n4、本软件可能存在bug'
            '当出现卡顿时请不要惊讶，告诉作者即可;\n5、软件正在进一步维护中，感谢使用;\n'
            '6、禁止一切违法利用行为，软件最终解释权归作者所有.')
APP_ICON = 'res/text.ico'
filePath_w2p=""

class FileDrop( wx.FileDropTarget ):
    def __init__(self, panel, statusText):
        wx.FileDropTarget.__init__(self)
        self.text=statusText
        self.panel=panel

    def OnDropFiles(self, x, y, filePath):
        path = filePath[0]
        # print(path)
        filePath_w2p = path #文件路径
        self.text.Label="选中文件路径：\n"+path
        # statusText0 = wx.StaticText(self.panel, -1, "转换中···", pos=(80, 72), size=(72, -1), style=wx.ALIGN_CENTER)
        # statusText0.SetBackgroundColour("White")
        # animation = AnimationCtrl(self.panel, -1, Animation('res/5.gif'), pos=(80, 90))    # 创建一个动画
        # animation.Play()    # 播放动图
        # mainFrame(None).word2PDF(path)  #开始转换
        # animation.Stop()
        # statusText0.Show(0)
        # animation.Destroy()
        
        return True

class mainFrame(wx.Frame):
    '''程序主窗口类，继承自wx.Frame'''
    
    id_open = wx.NewId()
    id_save = wx.NewId()
    id_quit = wx.NewId()
    id_help = wx.NewId()
    id_about = wx.NewId()
    
    id_CutPdf = wx.NewId()
    id_word2PDF = wx.NewId()
    id_PDF2word = wx.NewId()
    id_other = wx.NewId()
    fileName=""
    
    def __init__(self, parent):
        '''构造函数'''
        
        wx.Frame.__init__(self, parent, -1, APP_TITLE)
        self.SetBackgroundColour(wx.Colour(224, 224, 224))
        self.SetSize((600, 400))
        self.Resizable=-1
        self.Center()
        
        # to_bmp_image = wx.Image("res\\2.jpg", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        # self.bitmap = wx.StaticBitmap(self, -1, to_bmp_image, (0, 0))
        
        if hasattr(sys, "frozen") and getattr(sys, "frozen") == "windows_exe":
            exeName = win32api.GetModuleFileName(win32api.GetModuleHandle(None))
            icon = wx.Icon(exeName, wx.BITMAP_TYPE_ICO)
        else :
            icon = wx.Icon(APP_ICON, wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)
        
        self.tb1 = self._CreateToolBar('F')
        self.tb2 = self._CreateToolBar()
        # self.tbv = self._CreateToolBar('V')
        
        p_left = wx.Panel(self, -1)
        p_center0 = wx.Panel(self, -1)
        p_center1 = wx.Panel(self, -1)
        p_bottom = wx.Panel(self, -1)
        p_center0.SetBackgroundColour("White")
        p_bottom.SetBackgroundColour("White")
        
        
        statusText0 = wx.StaticText(p_left, -1, STATUS, pos=(0, 10), size=(200, 200), style=wx.ALIGN_LEFT)
        # statusText1 = wx.StaticText(p_left, -1, u"当前状态：word转PDF", pos=(15, 30), size=(200, -1), style=wx.ALIGN_CENTER)
        # statusText2 = wx.StaticText(p_left, -1, u"当前状态：PDF转word", pos=(15, 30), size=(200, -1), style=wx.ALIGN_CENTER)
        # statusText3 = wx.StaticText(p_left, -1, u"当前状态：其他", pos=(15, 30), size=(200, -1), style=wx.ALIGN_CENTER)
        
        
        statusText0 = wx.StaticText(p_bottom, -1, "将文件拖曳到此开始实现文件转PDF", pos=(5, 10), size=(200, 200), style=wx.ALIGN_LEFT)
        
        filepathText0 = wx.StaticText(p_center0, -1, "", pos=(5, 10), size=(500, -1), style=wx.ALIGN_LEFT)
        #文件拖曳
        fileDrop = FileDrop(p_center0,filepathText0)
        p_bottom.SetDropTarget( fileDrop)
        
        btn = wx.Button(p_left, -1, u'开始转换', pos=(30,230), size=(100, -1))
        btn.Bind(wx.EVT_BUTTON, lambda e, mark=filepathText0, panel=p_center0: self.OnSwitch(e, mark,panel))
        
        text0 = wx.StaticText(p_center0, -1, u'状态：文件转PDF', pos=(0, 170), size=(400, 40), style=wx.ALIGN_CENTER)
        # text0.SetBackgroundColour()
        text1 = wx.StaticText(p_center1, -1, u'我是第2页', pos=(40, 100), size=(200, -1), style=wx.ALIGN_LEFT)
     
        self._mgr = aui.AuiManager()
        self._mgr.SetManagedWindow(self)
        
        self._mgr.AddPane(self.tb1, 
            aui.AuiPaneInfo().Name("ToolBar1").Caption(u"工具条").ToolbarPane().Top().Row(0).Position(0).Floatable(False)
        )
        self._mgr.AddPane(self.tb2, 
            aui.AuiPaneInfo().Name("ToolBar2").Caption(u"工具条").ToolbarPane().Top().Row(0).Position(1).Floatable(True)
        )
        # self._mgr.AddPane(self.tbv, 
            # aui.AuiPaneInfo().Name("ToolBarV").Caption(u"工具条").ToolbarPane().Right().Floatable(True)
        # )
        
        self._mgr.AddPane(p_left,
            aui.AuiPaneInfo().Name("LeftPanel").Left().Layer(1).MinSize((200,-1)).Caption(u"操作区").MinimizeButton(True).MaximizeButton(True).CloseButton(True)
        )
        
        self._mgr.AddPane(p_center0,
            aui.AuiPaneInfo().Name("CenterPanel0").CenterPane().Show()
        )
        
        self._mgr.AddPane(p_center1,
            aui.AuiPaneInfo().Name("CenterPanel1").CenterPane().Hide()
        )
        
        self._mgr.AddPane(p_bottom,
            aui.AuiPaneInfo().Name("BottomPanel").Bottom().MinSize((-1,100)).Caption(u"消息区").CaptionVisible(False).Resizable(True)
        )
        
        self._mgr.Update()
    
    def setText(self, content):
        self.fileName=content
    
    def _CreateToolBar(self, d='H'):
        '''创建工具栏'''
        
        bmp_open = wx.Bitmap('res/file.png', wx.BITMAP_TYPE_ANY)
        bmp_save = wx.Bitmap('res/convert.png', wx.BITMAP_TYPE_ANY)
        bmp_help = wx.Bitmap('res/1.png', wx.BITMAP_TYPE_ANY)
        bmp_about = wx.Bitmap('res/menu.png', wx.BITMAP_TYPE_ANY)
        
        
        if d.upper() in ['V', 'VERTICAL']:
            tb = aui.AuiToolBar(self, -1, wx.DefaultPosition, wx.DefaultSize, agwStyle=aui.AUI_TB_TEXT|aui.AUI_TB_VERTICAL)
        else:
            tb = aui.AuiToolBar(self, -1, wx.DefaultPosition, wx.DefaultSize, agwStyle=aui.AUI_TB_TEXT)
        tb.SetToolBitmapSize(wx.Size(16, 16))
        
        if(d.upper() == 'F'):
            tb.AddSimpleTool(self.id_CutPdf, u'PDF拆分', bmp_open, u'PDF拆分')
            tb.AddSimpleTool(self.id_word2PDF, u'word转PDF', bmp_save, u'word转PDF')
            tb.AddSeparator()
            tb.AddSimpleTool(self.id_PDF2word, u'PDF转word', bmp_help, u'PDF转word')
            tb.AddSimpleTool(self.id_other, u'其他', bmp_about, u'其他')
            wx.EVT_TOOL( self, self.id_CutPdf, self.cutPDF )
            # wx.EVT_TOOL( self, self.id_word2PDF, self.OnSwitch )
            # wx.EVT_TOOL( self, self.id_PDF2word, self.OnSwitch )
            # wx.EVT_TOOL( self, self.id_other, self.OnSwitch )
        else:
            tb.AddSimpleTool(self.id_open, u'打开', bmp_open, u'打开文件')
            tb.AddSimpleTool(self.id_save, u'保存', bmp_save, u'保存文件')
            tb.AddSeparator()
            tb.AddSimpleTool(self.id_help, u'教程', bmp_help, u'帮助')
            tb.AddSimpleTool(self.id_about, u'关于', bmp_about, u'关于')
            wx.EVT_TOOL( self, self.id_open, self.OnOpenDirectory )
            # wx.EVT_TOOL( self, self.id_save, self.OnSwitch )
            # wx.EVT_TOOL( self, self.id_help, self.OnSwitch )
            # wx.EVT_TOOL( self, self.id_about, self.OnSwitch )
        tb.Realize()
        return tb
        
    def OnSwitch(self, evt, filepathText0, panel):
        path=filepathText0.Label.split("\n")[1]    #获得文件路径
        
        if  path== "":
            d=wx.MessageDialog(None, u"请先选择文件再点击转换哦", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal()==wx.ID_OK:
                pass
            d.Destroy()
            return
        statusText0 = wx.StaticText(panel, -1, "转换中···", pos=(80, 72), size=(72, -1), style=wx.ALIGN_CENTER)
        statusText0.SetBackgroundColour("White")
        animation = AnimationCtrl(panel, -1, Animation('res/5.gif'), pos=(80, 90))    # 创建一个动画
        animation.Play()    # 播放动图
        mainFrame(None).word2PDF(path)  #开始转换
        animation.Stop()
        statusText0.Show(0)
        animation.Destroy()
            
        
    def cutPDF(self, evt):
        print("当前处于PDF切割")
        # self._mgr.Update()
        
    def OnOpenDirectory(self, event):
        '''
        打开开文件对话框
        '''
        dlg = wx.FileDialog(self,u"选择文件夹",style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            #print(dlg.GetPath()) #文件夹路径
            filePath_w2p=dlg.GetPath()
            self.word2PDF(filePath_w2p)
            
        dlg.Destroy()
    
    def cutPDF(self,filePath):
        pass
        
    def word2PDF(self, filePath):
        flag=1
        # flag=0  #flag
        # doc2pdf(filePath)
        # dlg = wx.MessageDialog(None, u"确认将文件转为pdf?", u"提示", wx.YES_NO | wx.ICON_QUESTION)
        # if dlg.ShowModal() == wx.ID_YES:
            # flag=1
        # dlg.Destroy()
        
        result=1
        if flag==1:
            result = doc2pdf(filePath)
        if result==1:
            d=wx.MessageDialog(None, u"转换失败，可能已经存在该PDF文件", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            if d.ShowModal()==wx.ID_OK:
                pass
            d.Destroy()
        else:
            pass
            # d=wx.MessageDialog(None, u"转换成功", u"提示", wx.YES_NO | wx.ICON_QUESTION)
            # if d.ShowModal()==wx.ID_OK:
                # pass
            # d.Destroy()
        
class mainApp(wx.App):
    def OnInit(self):
        self.SetAppName(APP_TITLE)
        self.Frame = mainFrame(None)
        self.Frame.Show()
        return True

if __name__ == "__main__":
    app = mainApp()
    app.MainLoop()
