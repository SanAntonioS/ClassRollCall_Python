import os
import openpyxl
import random
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton,  QPlainTextEdit, QMessageBox, QFileDialog, QTableWidgetItem
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QDir, Qt
import win32com.client as win32
from datetime import datetime
import pyttsx3

studentNum = 0

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
    workbook.close()
    
    return student_id_list
    
class Stats():

    def __init__(self):
        self.ui = QUiLoader().load('main.ui')
        self.ui.tableWidget.setColumnWidth(0, 70)
        self.ui.tableWidget.setColumnWidth(1, 120)
        self.ui.tableWidget.setColumnWidth(2, 120)
        self.ui.RollCall.clicked.connect(self.RollCall)
        self.ui.Absenteeism.clicked.connect(self.Absenteeism)
        self.ui.ImportList.clicked.connect(self.ImportList)
        self.ui.Sort.clicked.connect(self.Sort)
        self.ui.SaveData.clicked.connect(self.SaveData)
        self.ui.OpenData.clicked.connect(self.OpenData)


    def RollCall(self):
        global studentNum
        
        #显示在左边三栏上
        self.ui.studentName.setPlainText(student_id_list[studentNum][2].value)
        self.ui.studentClass.setPlainText(student_id_list[studentNum][0].value)
        self.ui.studentID.setPlainText(student_id_list[studentNum][1].value)
        
        #播放被点到的人的姓名
        speaker.say(student_id_list[studentNum][2].value)
        speaker.runAndWait()
        
        #更新右边表格数据
        self.ui.tableWidget.insertRow(0)
        nameItem = QTableWidgetItem("%s" % student_id_list[studentNum][2].value)
        classItem = QTableWidgetItem("%s" % student_id_list[studentNum][0].value)
        idItem = QTableWidgetItem("%s" % student_id_list[studentNum][1].value)
        statusItem = QTableWidgetItem("%s" % '已到')
        self.ui.tableWidget.setItem(0,0,nameItem)
        self.ui.tableWidget.setItem(0,1,classItem)
        self.ui.tableWidget.setItem(0,2,idItem)
        self.ui.tableWidget.setItem(0,3,statusItem)
        studentNum += 1
        #还需要实现按学号排序
        
    def Sort(self):
        #bug:排序后再点旷课后，不是显示在应该显示的学生表格上
        self.ui.tableWidget.sortItems(2,Qt.AscendingOrder)
    
    def Absenteeism(self):
        statusItem = QTableWidgetItem("%s" % '旷课')
        self.ui.tableWidget.setItem(0,3,statusItem)
        
    def ImportList(self):
        filePath = QFileDialog.getOpenFileName(self.ui, "导入名单")[0]
        filePath = QDir.toNativeSeparators(filePath)
        print(filePath)
        xls2xlsx(filePath)
    
    def SaveData(self):
        workbook = openpyxl.load_workbook("Log.xlsx")	# 返回一个workbook数据类型的值
        sheet = workbook.active
        sheet.insert_cols(7)
        sheet['G2'] = datetime.now().strftime('%Y-%m-%d')
        
        tableRow = self.ui.tableWidget.rowCount()
        for i in range(tableRow):
            tableName = self.ui.tableWidget.item(i,0).text()
            for row in range(3,49):
                name = sheet.cell(row ,3).value
                if name == tableName:
                    sheet.cell(row, 7, self.ui.tableWidget.item(i,3).text())
        workbook.save("Log.xlsx")
        workbook.close()
        
    def OpenData(self):
        os.startfile("Log.xlsx")
    

speaker = pyttsx3.init()
msg = "你好"
speaker.say(msg)
speaker.runAndWait()
app = QApplication([])
stats = Stats()
stats.ui.show()
student_id_list = ReadData()
app.exec()
