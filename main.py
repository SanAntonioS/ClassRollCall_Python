import os
import openpyxl
import random
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton,  QPlainTextEdit, QMessageBox, QFileDialog
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QDir
import win32com.client as win32

#将xls格式文件转换为xlsx格式，openpyxl包无法处理xls格式文件
def xls2xlsx(filePath):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filePath)
    wb.SaveAs(filePath[:-38]+"Log.xlsx", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
    # wb.SaveAs(fname[:-1], FileFormat = 56)      #FileFormat = 56 is for .xls extension
    wb.Close()
    excel.Application.Quit()
    
def ReadData():
    workbook = openpyxl.load_workbook("Log.xlsx")	# 返回一个workbook数据类型的值
    sheet = workbook.active
    student_id = sheet['A3:F49']
    student_id_list = list(student_id)
    random.shuffle(student_id_list)
    print(student_id_list[0][0].value)
    print(student_id_list[0][1].value)
    print(student_id_list[0][2].value)
    
    return student_id_list
    
class Stats():

    def __init__(self):
        self.ui = QUiLoader().load('main.ui')

        self.ui.RollCall.clicked.connect(self.RollCall)
        self.ui.Absenteeism.clicked.connect(self.Absenteeism)
        self.ui.ImportList.clicked.connect(self.ImportList)
        self.ui.OpenList.clicked.connect(self.OpenList)


    def RollCall(self):
        self.ui.studentName.append(student_id_list[0][2].value)
        self.ui.studentClass.append(student_id_list[0][0].value)
        self.ui.studentID.append(student_id_list[0][1].value)
    
    def Absenteeism(self):
        info = self.textEdit.toPlainText()
        
    def ImportList(self):
        filePath = QFileDialog.getOpenFileName(self.ui, "导入名单")[0]
        filePath = QDir.toNativeSeparators(filePath)
        print(filePath)
        xls2xlsx(filePath)
    
    def OpenList(self):
        info = self.textEdit.toPlainText()
    

app = QApplication([])
stats = Stats()
stats.ui.show()
student_id_list = ReadData()
app.exec()
