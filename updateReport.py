import os
from datetime import datetime
import win32com.client as win32

def update_report(report_path, debug=False):
    # --- Excel setup ---
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Open(os.path.join(report_path, 'DAILY ORDER FULFILLMENT (F.U.D).xlsx'))

    sheetName = "SUMMARY"
    ws = wb.Worksheets(sheetName)

    summaryTable = ws.ListObjects.Item("SummaryTable")




report_path = 'C:\\Users\\koichik\\Documents\\fudFReport'

update_report(report_path = report_path)