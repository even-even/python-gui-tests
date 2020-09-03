from comtypes.client import CreateObject
import os

project_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
filename = os.path.join(project_dir, "data/groups.xlsx")

xl = CreateObject("Excel.Application")
xl.Visible = 1
wb = xl.Workbooks.Add()
for i in range(10):
    xl.Range["A%s" % (i+1)].Value[()] = "test group %s" % str(i+1)
wb.SaveAs(filename)
xl.Quit()
