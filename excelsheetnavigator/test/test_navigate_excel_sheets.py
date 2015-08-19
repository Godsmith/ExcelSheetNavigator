import unittest
import win32com.client
from excelsheetnavigator.navigate_excel_sheets import main
from excelsheetnavigator.navigate_excel_sheets import NavigatorModel
from excelsheetnavigator.navigate_excel_sheets import WindowMgr
from tempfile import gettempdir
import tempfile
from mock import Mock

class TestNavigatorModel(unittest.TestCase):
    
    def setUp(self):
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.workbook = self.excel.Workbooks.Add()
        self.excel.Application.DisplayAlerts = False
        self.workbook.SaveAs(gettempdir() + "\esn_workbook.xls")
    
    def tearDown(self):
        self.workbook.Close()
        self.excel.Application.DisplayAlerts = True
    
    def add_worksheets(self, names):
        for sheetname in names:
            sheet = self.workbook.Worksheets.Add()
            sheet.Name = sheetname
    
    def test_text_to_show_one_sheet(self):
        sheet = self.workbook.Worksheets.Add()
        sheet.Name = "esn_worksheet"
        model = NavigatorModel(WindowMgr())
        assert model.text_to_show('sn_wor') == 'esn_worksheet'
        
    def test_text_to_show_two_sheets(self):
        self.add_worksheets(["esn_worksheet1", "esn_worksheet2"])
        model = NavigatorModel(WindowMgr())
        assert model.text_to_show('sn_wor') == 'esn_worksheet1\nesn_worksheet2'
        
    def test_text_to_show_oneOfTwoSheets(self):
        self.add_worksheets(["esn_worksheet1", "esn_worksheet2"])
        model = NavigatorModel(WindowMgr())
        assert model.text_to_show('sheet1') == "esn_worksheet1"
    
    def test_switchToFirstWorksheetInList_oneSheetNoTextEntered(self):
        self.add_worksheets(["esn_worksheet1"])
        window_mgr_mock = Mock(spec=WindowMgr)
        model = NavigatorModel(window_mgr_mock)
        model.text_to_show('sheet1')
        
        model.switch_to_first_worksheet_in_list()
        
        window_mgr_mock.find_window_text.assert_called_with('esn_workbook.xls')
        
    def test_switchToFirstWorksheetInList_tenWorkbooksTwentySheetsEach(self):
        workbooks = []
        for i in range(10):
            workbooks.append(self.excel.Workbooks.Add())
        for workbook in workbooks:
            for i in range(20):
                workbook.Worksheets.Add()
        first_workbook_name = workbooks[0].Name
        workbooks[0].Sheets(1).Name = 'a1b2c3d4'
        window_mgr_mock = Mock(spec=WindowMgr)
        model = NavigatorModel(window_mgr_mock)
        model.text_to_show('2c3')
        
        model.switch_to_first_worksheet_in_list()
        
        window_mgr_mock.find_window_text.assert_called_with(first_workbook_name)
        
        
        for workbook in workbooks:
            workbook.Close()
        

if __name__ == '__main__':
    unittest.main()