#-*- encoding:utf-8 -*-
import wx
import os
import sys
import xlrd
#import wx.lib.mixins.listctrl as listmix
import pickle
from ObjectListView import ObjectListView, ColumnDefn

class OlvObject(object):
    def __init__(self, filename, sheet):
        self.filename = filename
        self.sheet = sheet
    
    def GetId(self):
        return id(self)


class FileDropTarget(wx.FileDropTarget):  
    def __init__(self, frame):
        wx.FileDropTarget.__init__(self)  
        self.frame = frame
  
    def OnDropFiles(self, x, y, filenames):
        if len(filenames) != 1:
            self.frame.tc1.AppendText(u'错误：一次请只拖放一个文件（每个文件都需要选择工作表）'+os.linesep)
            return
        else:
            self.filename = filenames[0]
            self.frame.tc1.AppendText(u"选择文件".encode('gbk') + self.filename.encode('gbk'))
            self.frame.tc1.AppendText(os.linesep)  
            self.frame.readFile(self.filename)
            self.frame.lc1_show_sheet_list()


class SrzmFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, None, size=(800,600), title=u'收入证明')
        # create the controls
        self.statusBar = self.CreateStatusBar()
        self.init_data()
        self.createMainWindow()
        self.listfilename = 'file_list.txt'
        self.col_li = [u'姓名',u'应发工资',u'实发工资',u'费用补贴']
        self.InitLc2()

    def OnLc1ActiveItem(self, event):
        self.sb2.SetLabelText(u'将下一个工资文件拖放到这里')
        index = self.lc1.GetFocusedItem()
        if index == -1:
            self.setPromtingMsg(u'没有选择工作表\n')
            return
        else:
            self.workingSheet = self.workingBook.sheet_by_index(index)
            self.analyseSheet(self.workingSheet, self.dt.filename)
        self.lc1.DeleteAllItems()

    def analyseSheet(self, sheet, filename):
        self.setPromtingMsg(u'读取工作表\"'+sheet.name+u'\"...')
        row_slice1 = sheet.row_slice(1)  # list of xlrd.sheet.Cell ; 默认读取第二行 index:1
        title_list = []
        for index in row_slice1:
            title_list.append(index.value.strip())  # 空列也包含在内
        if set(self.col_li).issubset(set(title_list)):
            #self.setPromtingMsg(u'ok.')
            sheet_item = [sheet]
            for item in self.col_li:
                sheet_item.append(title_list.index(item))
            self.sheetlist.append(sheet_item)
            obj = OlvObject(filename, sheet.name)
            self.olv_items.append(obj)
            self.lc2.SetObjects(self.olv_items)
            self.setPromtingMsg(u'已更新文件列表"')
        else:
            self.setPromtingMsg(u'失败: 该工作表第2行不含\"应发工资\"或不含\"实发工资\"')

    def OnClear(self, event):
        self.init_data()
        self.clear_view()

    def init_data(self):
        self.sheetlist = [] #list of [sheet,index(姓名),index(应发),index(实发)]
        self.data = [] # list of {sheet:{u'姓名':name,u'应发工资':amount1,u'实发工资':amount2}}
        self.results = []
        self.filelist = []
        self.olv_items = []

    def clear_view(self):
        self.lc1.DeleteAllItems()
        #self.lc2.DeleteAllItems()
        self.tc3.SetValue("")
        self.search.Clear()

    def OnDoSearch(self, event):
        name = self.search.GetValue().strip()
        if not name:
            self.setPromtingMsg(u'没有输入姓名')
            return
        self.data = []
        for index_sheet,sheet in enumerate(self.sheetlist):
            for i in range(2,sheet[0].nrows):
                if sheet[0].row_slice(i)[sheet[1]].value == name:
                    d = {}
                    d[u'编号'] = self.olv_items[index_sheet].GetId()
                    for index_item,item in enumerate(self.col_li):
                        d[item] = sheet[0].row_slice(i)[sheet[index_item+1]].value
                    self.data.append(d)
        self.write_output()

    def OnLc2ActiveItem(self, event):
        index = self.lc2.GetFocusedItem()
        if index == -1:
            self.setPromtingMsg(u'没有选择工作表')
            return
        self.lc2.DeleteItem(index)
        del(self.sheetlist[index])
        del(self.olv_items[index])
        self.setPromtingMsg(u'已去除文件')

    def OnReadListFile(self, event):
        self.init_data()
        self.lc2.SetObjects(self.olv_items)
        import codecs
        try:
            f = codecs.open(self.listfilename,'r','gbk')
        except ValueError:
            self.setPromtingMsg(u'错误的文件编码')
            return
        lines = [line.strip() for line in f]
        f.close()
        for li in lines:
            if li[0] != u'#':
                l = li.split(';')
                self.filelist.append([l[0],l[1]])
        for sheet_name,file_name in self.filelist:
            self.readFile(file_name)
            self.workingSheet = self.workingBook.sheet_by_name(sheet_name)
            self.analyseSheet(self.workingSheet, file_name)
        self.setPromtingMsg(u'已读取文件列表')

    def OnOpenListFile(self, event):
        import win32api
        # win32api.ShellExecute(0, 'open', 'notepad.exe', self.listfilename,'',1)
        win32api.ShellExecute(0, 'open', self.listfilename,'','',1)


    def OnSaveLc2(self, event):
        f = open('lc2','w')
        pickle.dump(self.olv_items, f)
        f.close()

    def OnReadLc2(self, event):
        self.init_data()
        self.lc2.SetObjects(self.olv_items)
        try:
            f = open('lc2','r')
        except IOError:
            pass
        else:
            olv_list = pickle.load(f)
            for obj in olv_list:
                wb = xlrd.open_workbook(obj.filename, encoding_override="cp936")
                sheet = wb.sheet_by_name(obj.sheet)
                self.analyseSheet(sheet,obj.filename)
            f.close()

    def write_output(self):
        s = {u'工作表数':0}
        for i in self.col_li:
            if i != u'姓名':
                s[i] = 0
        if not self.data:
            self.setPromtingMsg(u'没有工资表，或没有找到数据')
            return
        for d in self.data:
            for i in [u'编号'] +self.col_li:
                self.tc3.AppendText(i)
                self.tc3.AppendText(":")
                self.tc3.AppendText(unicode(d[i]))
                self.tc3.AppendText(' ')
                if i == u'姓名':
                    s[i] = d[i]
                elif i == u'编号':
                    pass
                else:
                    s[i] += d[i]
            s[u'工作表数'] += 1
            self.tc3.AppendText(os.linesep)
        s[u'基本工资应发'] = s[u'应发工资']-s[u'费用补贴']
        self.tc3.AppendText(u'==>工作表数: '+str(s[u'工作表数']) + os.linesep)
        for i in (self.col_li +[u'基本工资应发']):
            if i not in [u'编号', u'姓名']:
                self.tc3.AppendText(i+u'平均:')
                self.tc3.AppendText(unicode(round(s[i]/s[u'工作表数'],2)))
            else:
                self.tc3.AppendText(i+u':')
                self.tc3.AppendText(unicode(s[i]))
            self.tc3.AppendText(os.linesep)
        self.tc3.AppendText(os.linesep)

    def readFile(self, filename):
        try:
            self.workingBook = xlrd.open_workbook(filename, encoding_override="cp936")
        except xlrd.XLRDError:
            self.setPromtingMsg(u"不是合法的excel文件")

    def lc1_show_sheet_list(self):
        sheet_name_list=self.workingBook.sheet_names()
        self.statusBar.SetStatusText(u'请选择工资表')
        self.sb2.SetLabelText(u'选择工资文件所在的sheet')
        for i in sheet_name_list:
            self.lc1.InsertStringItem(sheet_name_list.index(i), i)

    def setPromtingMsg(self, msg):
        self.tc1.AppendText(msg+os.linesep)
        self.statusBar.SetStatusText(msg,0)

    def createMainWindow(self):
        mainPanel = wx.Panel(self, wx.ID_ANY)
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        mainPanel.SetSizer(mainSizer)
        # Splitter
        splitter = wx.SplitterWindow(mainPanel,
                                     style=wx.SP_LIVE_UPDATE | wx.SP_3DSASH)
        splitter.SetMinimumPaneSize(230)
        mainSizer.Add(splitter, 2, wx.EXPAND)\
        # second number has to be '2' in case the splitter can expand
        # 设为0，对象将不改变尺寸；大于0，则sizer中的child
        # 根据因数分割sizer的总尺寸
        lPanel = wx.Panel(splitter, wx.NewId())
        rPanel = wx.Panel(splitter, wx.NewId())

        self.tc1 = wx.TextCtrl(rPanel, -1,"",style=wx.TE_MULTILINE|wx.TE_READONLY)
        self.lc1 = wx.ListCtrl(lPanel,id=wx.NewId(),style=wx.LC_LIST)
        self.ofb = wx.Button(lPanel, wx.NewId(), u'编辑文件清单', style=wx.BU_EXACTFIT)
        self.rfb = wx.Button(lPanel, wx.NewId(), u'读取文件清单', style=wx.BU_EXACTFIT)
        self.calcb = wx.Button(lPanel, wx.NewId(), u'计算')
        self.clearb = wx.Button(lPanel, wx.NewId(), u'全部清空')
        self.readb = wx.Button(rPanel, wx.NewId(), u'读取列表') #,(-1,-1),wx.DefaultSize)
        self.saveb = wx.Button(rPanel, wx.NewId(), u'保存列表') #,(-1,-1),wx.DefaultSize)
        self.search = wx.SearchCtrl(lPanel, size=(-1,-1), style=wx.TE_PROCESS_ENTER)

        self.sb1 = wx.StaticBox(rPanel, -1, u"log") 
        self.sbsz1 = wx.StaticBoxSizer(self.sb1, wx.VERTICAL)
        self.dt = FileDropTarget(self)
        self.lc1.SetDropTarget(self.dt)  
        self.sbsz1.Add(self.tc1, 1, wx.EXPAND|wx.ALL,5)

        self.sb2 = wx.StaticBox(lPanel, -1, u"将工资文件拖放到这里") 
        self.sbsz2 = wx.StaticBoxSizer(self.sb2, wx.VERTICAL)
        self.sbsz2.Add(self.lc1, 1, wx.EXPAND)

        self.sb3 = wx.StaticBox(lPanel, -1, u"操作说明")
        self.sbsz3 = wx.StaticBoxSizer(self.sb3, wx.VERTICAL)
        self.st = wx.StaticText(lPanel, -1, \
                u"1. 请将工资文件逐个拖放到下面；\n\
2. “双击”工资表所在的sheet;读取到的应发工资和实发工资数会列在此对话框的下方。\n\
3. 输入要计算的人名,\n4. 点击“计算”按钮计算出平均工资数和合计工资数。\n\n")
        self.sbsz3.Add(self.st, 1, wx.EXPAND|wx.ALL, 5)

        self.sb4 = wx.StaticBox(lPanel, -1, u"姓名")
        self.searchsz = wx.StaticBoxSizer(self.sb4,wx.VERTICAL)
        self.searchsz.Add(self.search,1,wx.EXPAND)

        self.btsz1 = wx.BoxSizer(wx.HORIZONTAL)
        self.btsz1.Add(self.ofb,1,wx.EXPAND)
        self.btsz1.Add(self.rfb,1,wx.EXPAND)
        self.btsz2 = wx.BoxSizer(wx.HORIZONTAL)
        self.btsz2.Add(self.clearb,1,wx.EXPAND)
        self.btsz2.Add(self.calcb,1,wx.EXPAND)


        self.sb5 = wx.StaticBox(rPanel, -1, u"输出")
        self.tc3 = wx.TextCtrl(rPanel,-1,"",style=wx.TE_MULTILINE|wx.TE_READONLY)
        self.sbsz5 = wx.StaticBoxSizer(self.sb5, wx.VERTICAL)
        self.sbsz5.Add(self.tc3,1, wx.EXPAND|wx.ALL,5)

        self.sb6 = wx.StaticBox(rPanel, -1, u'已读工资表清单')
        self.lc2 = ObjectListView(rPanel, style = wx.LC_REPORT|wx.SUNKEN_BORDER)
        self.lc2.SetColumns([
            ColumnDefn(u'编号','right',80,"GetId"),
            ColumnDefn(u'文件名','left',280,"filename"),
            ColumnDefn(u'工作表','right',60,"sheet")
            ])
        self.sbsz6 = wx.StaticBoxSizer(self.sb6, wx.VERTICAL)
        self.sbsz6.Add(self.lc2,6,wx.EXPAND|wx.ALL,0)
        self.bsz1 = wx.BoxSizer(wx.HORIZONTAL)
        self.bsz1.Add(self.readb,0,wx.ALL,1)
        self.bsz1.Add(self.saveb,0,wx.ALL,1)
        self.sbsz6.Add(self.bsz1,1,wx.CENTER|wx.ALL,0)
        
        lBox = wx.BoxSizer(wx.VERTICAL)
        rBox = wx.BoxSizer(wx.VERTICAL)

        lBox.Add(self.sbsz3, 3, wx.EXPAND)
        lBox.Add(self.sbsz2, 4, wx.EXPAND)
        lBox.Add(self.searchsz, 1, wx.EXPAND)
        lBox.Add(self.btsz1, 1, wx.EXPAND)
        lBox.Add(self.btsz2, 1, wx.EXPAND)
        rBox.Add(self.sbsz6, 2, wx.EXPAND)
        rBox.Add(self.sbsz5, 6, wx.EXPAND)
        rBox.Add(self.sbsz1, 4, wx.EXPAND)

        lPanel.SetSizer(lBox)
        rPanel.SetSizer(rBox)
        splitter.SplitVertically(lPanel, rPanel, 100)

        self.Bind(wx.EVT_BUTTON, self.OnClear, self.clearb)
        self.Bind(wx.EVT_BUTTON, self.OnDoSearch, self.calcb)
        self.Bind(wx.EVT_BUTTON, self.OnOpenListFile, self.ofb)
        self.Bind(wx.EVT_BUTTON, self.OnReadListFile, self.rfb)
        self.lc1.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnLc1ActiveItem)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnDoSearch, self.search)
        self.lc2.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnLc2ActiveItem)
        # self.lc2.Bind(wx.EVT_LIST_DELETE_ITEM, self.OnLc2ItemDelete)
        # self.lc2.Bind(wx.EVT_LIST_INSERT_ITEM, self.OnLc2ItemInsert)
        self.Bind(wx.EVT_BUTTON, self.OnReadLc2, self.readb)
        self.Bind(wx.EVT_BUTTON, self.OnSaveLc2, self.saveb)

    def InitLc2(self):
        #self.lc2.SetObjects(self.olv_items)
        self.lc2.SetEmptyListMsg(u"请选择工作表")

class App(wx.App):
    def OnInit(self):
        self.frame = SrzmFrame(None)
        self.frame.Show()
        return True

if __name__ == "__main__":
    app = App()
    app.MainLoop()
