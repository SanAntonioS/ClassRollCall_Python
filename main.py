import os
import openpyxl
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton,  QPlainTextEdit, QMessageBox, QFileDialog
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QDir
import win32com.client as win32

def xls2xlsx(filePath):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filePath)
    wb.SaveAs(filePath+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
    # wb.SaveAs(fname[:-1], FileFormat = 56)      #FileFormat = 56 is for .xls extension
    wb.Close()
    excel.Application.Quit()
    
def ReadData(filePath):
    workbook = openpyxl.load_workbook(filePath+"x")	# 返回一个workbook数据类型的值
    print(workbook.sheetnames)	# 打印Excel表中的所有表
    sheet = workbook.active
    sheet.title = 'sheet1'
    print(sheet)
    os.remove(filePath+"x")
    
class Stats():

    def __init__(self):
        self.ui = QUiLoader().load('main.ui')

        self.ui.RollCall.clicked.connect(self.RollCall)
        self.ui.Absenteeism.clicked.connect(self.Absenteeism)
        self.ui.ImportList.clicked.connect(self.ImportList)
        self.ui.OpenList.clicked.connect(self.OpenList)


    def RollCall(self):
        self.ui.name.append('hello')
    
    def Absenteeism(self):
        info = self.textEdit.toPlainText()
        
    def ImportList(self):
        filePath = QFileDialog.getOpenFileName(self.ui, "导入名单")[0]
        filePath = QDir.toNativeSeparators(filePath)
        print(filePath)
        xls2xlsx(filePath)
        ReadData(filePath)
    
    def OpenList(self):
        info = self.textEdit.toPlainText()
    

app = QApplication([])
stats = Stats()
stats.ui.show()
app.exec()