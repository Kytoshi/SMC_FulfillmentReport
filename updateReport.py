import os
from datetime import datetime
import time
import win32com.client as win32

def update_report(report_path, debug=False):
    # --- Excel setup ---
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = excel.Workbooks.Open(os.path.join(report_path, 'DAILY ORDER FULFILLMENT (F.U.D).xlsx'))


    # --- Refresh each connection safely ---
    print("Refreshing connections...")
    for conn in wb.Connections:
        try:
            # Force synchronous refresh if possible
            if conn.Type in [1, 2]:  # OLEDB=1, ODBC=2
                try:
                    conn.OLEDBConnection.BackgroundQuery = False
                except AttributeError:
                    pass  # Not all OLEDB connections expose BackgroundQuery

            conn.Refresh()

            # Wait for connection to finish if it has 'Refreshing' attribute
            try:
                while conn.Refreshing:
                    time.sleep(1)
            except AttributeError:
                pass  # Connection type does not have Refreshing
            print(f"Refreshed: {conn.Name}")
        except Exception as e:
            print(f"Error refreshing {conn.Name}: {e}")

    # --- Refresh pivot tables ---
    print("Refreshing pivot tables...")
    for ws in wb.Worksheets:
        for pt in ws.PivotTables():
            try:
                pt.RefreshTable()
                print(f"Refreshed pivot table: {pt.Name} on {ws.Name}")
            except Exception as e:
                print(f"Error refreshing pivot table {pt.Name}: {e}")

    # --- Save workbook ---
    wb.Save()
    print("Workbook saved.")

    # --- Restore screen updating ---
    excel.ScreenUpdating = True

    # --- Leave Excel open ---
    print("Excel refresh complete. Workbook remains open.")
    return excel, wb


report_path = 'C:\\Users\\koichik\\Documents\\fudFReport'

update_report(report_path = report_path)