import win32com.client as win32
def xls2xlsx
    fname = r"C:\Users\snow\Desktop\Python_homework\2022-2023-2_杭州_Python机器学习_学硕1-班级名单.xls"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
    # wb.SaveAs(fname[:-1], FileFormat = 56)      #FileFormat = 56 is for .xls extension
    wb.Close()
    excel.Application.Quit()