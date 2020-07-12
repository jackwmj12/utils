import os
import fnmatch
from win32com.client import Dispatch
import win32com.client
import collections

#返回path目录以及其子目录下下的所有fnexp类型的文件，
#iterfindfiles(path, "*.xls"):返回所有xls文件
def iterfindfiles(path,fnexp):
    for root,dirs,files in os.walk(path):
        for filename in fnmatch.filter(files,fnexp):
            yield  os.path.join(root,filename)

class easyExcel(object):
    """A utility to make it easier to get at Excel.    Remembering
    to save the data is your problem, as is    error handling.
    Operates on one workbook at a time."""
    def __init__(self, filename=None):  # 打开文件或者新建文件（如果不存在的话）
        if os.name != "nt":
            print("Error System")
            return False
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible = 0  #0代表隐藏对象，但可以通过菜单再显示-1代表显示对象2代表隐藏对象，但不可以通过菜单显示，只能通过VBA修改为显示状态"""
        self.xlApp.DisplayAlerts = 0  # 后台运行，不显示，不警告
        if os.path.isfile(filename) is True:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):  # 保存文件
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):  # 关闭文件
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def get_rows(self,sheet = "Sheet1"):
        sht = self.xlBook.Worksheets(sheet)
        return sht.usedrange.rows.count

    def get_cols(self,sheet = "Sheet1"):
        sht = self.xlBook.Worksheets(sheet)
        return sht.usedrange.columns.count

    def getCell(self, sheet="Sheet1", row=1, col=1):  # 获取单元格的数据
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet="sheet1", row=1, col=1, value=""):  # 设置单元格的数据
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):  # 获得一块区域的数据，返回为一个二维元组
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  # 插入图片
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def cpSheet(self, before):  # 复制工作表
        "copy sheet"
        shts = self.xlBook.Worksheets
        shts(1).Copy(None, shts(1))

    #datalist 传入的字典列表
    #sheet工作表，默认为sheet1，
    # key_List 字典的keys,是一个列表，若有顺序，需要手动输入，如不设置，系统自动排序
    # header，excel头，如不设置，默认为key_List
    def setDict_list(self,data_list,sheet="Sheet1",key_list = None,header = None):
        if type(data_list)!=list:
            print("not list")
            return False
        elif type(data_list[0])!=dict:
            print("not dict")
            return False
        if key_list == None:
            key_list = list(data_list[0].keys())
        if header ==None:
            header = key_list
        cols = len(key_list)
        for col in range(0,cols):
            self.setCell(sheet=sheet,row=1,col=col+1,value=header[col])#设置头
        for row,item in enumerate(data_list):
            for col in range(0,cols):                    #设置内容
                self.setCell(sheet=sheet,row=row+2,col=col+1,value=item[key_list[col]])

    # 获取整个工作表内容，返回一个字典列表。key为header也就是excel首行内容。
    def get_content(self,sheet="Sheet1"):
        key_list =[];data_list =[]
        rows = self.get_rows(sheet)
        cols = self.get_cols(sheet)
        for col in range(0, cols):
            key_list.append((self.getCell(sheet=sheet,row=1, col=col + 1)))
        for row in range(rows):
            data = collections.OrderedDict()   #创建有序字典
            for col in range(cols):
                # self.__addWord(data,key_list[col],self.getCell(sheet=sheet,row=row+2,col=col+1))
                # print(type(self.getCell(sheet=sheet,row=row + 2, col=col + 1)))
                data[key_list[col]] = str(self.getCell(sheet=sheet,row=row + 2, col=col + 1))
            if data[key_list[0]] == None:
                break
            data_list.append(data)
        return key_list,data_list

    def get_header(self,sheet="Sheet1"):
        key_list = []
        cols = self.get_cols(sheet)
        for col in range(0, cols):
            content = self.getCell(sheet=sheet, row=1, col=col + 1)
            if content == None:
                break
            key_list.append(content)
        return key_list

    def row_delete(self,sheet = "Sheet1",row_start=0,row_stop=0):
        sht = self.xlBook.Worksheets(sheet)
        # for i in range(row_start,row_stop)
        avg = "{}:{}".format(str(row_start),str(row_stop))
        print(avg)
        sht.rows(avg).delete

    def __addWord(self,theIndex, word, pagenumber):
        theIndex.setdefault(word, []).append(pagenumber)  # 存在就在基础上加入列表，不存在就新建个字典

    def get_sheet(self):
        self.xlBook.Sheets(1).Name

    #如果输入的字典值全部为None，则返回False
    #如果不全为None，返回True
    def check_dict(self,data):
        print(data)
        print(type(data))
        content = data.values
        for item in content:
            if item != None:
                return True
        return False