import win32com.client as win32
from win32com.client import Dispatch

path = "C:\\Users\\Yacine\\Desktop\\ProjetGenerator\\Classeur1.xlsx"

xlApp = Dispatch("Excel.Application")

xlApp.Visible = True

xlWb = xlApp.Workbooks.Open(path)

ValueRes = xlWb.ActiveSheet.Cells(1,1).Value 

print (ValueRes)

xlApp.Application.Quit()