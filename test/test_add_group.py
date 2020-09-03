import pytest
from comtypes.client import CreateObject
import os

# retrieving test data from excel file
project_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
filename = os.path.join(project_dir, "data/groups.xlsx")
xl = CreateObject("Excel.Application")
xl.Visible = 1
wb = xl.Workbooks.Open(filename)
ws = wb.Worksheets(1)
testdata = []
for row in range(1, 11):
    cell_data = ws.Cells[row, 1].Value()
    testdata.append(cell_data)
xl.Quit()


@pytest.mark.parametrize("group", testdata, ids=[repr(x) for x in testdata])
def test_add_group(app, group):
    old_list = app.groups.get_group_list()
    app.groups.add_new_group(group)
    new_list = app.groups.get_group_list()
    old_list.append(group)
    assert sorted(old_list) == sorted(new_list)
