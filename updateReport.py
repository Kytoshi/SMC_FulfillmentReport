import os
from datetime import datetime
import time
import win32com.client as win32

def update_report(report_path, debug=False, progress_callback=None):
    """
    Updates the Excel report with a continuous progress bar (0-100%).
    """
    try:
        if progress_callback:
            progress_callback("start")  # Start progress bar

        # --- Step 0: Open workbook ---
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        wb = excel.Workbooks.Open(os.path.join(report_path, 'DAILY ORDER FULFILLMENT (F.U.D).xlsx'))

        # --- Step 1: Refresh connections ---
        print("Refreshing connections...")
        total_connections = wb.Connections.Count
        for i, conn in enumerate(wb.Connections):
            try:
                if conn.Type in [1, 2]:
                    try:
                        conn.OLEDBConnection.BackgroundQuery = False
                    except AttributeError:
                        pass
                conn.Refresh()
                try:
                    while conn.Refreshing:
                        time.sleep(1)
                except AttributeError:
                    pass
                print(f"Refreshed: {conn.Name}")
            except Exception as e:
                print(f"Error refreshing {conn.Name}: {e}")
            
            # Update progress continuously (~30% for connections)
            if progress_callback and total_connections > 0:
                progress = (i + 1) / total_connections * 30
                progress_callback("update", progress)

        # --- Step 2: Refresh pivot tables ---
        print("Refreshing pivot tables...")
        all_pts = [(ws, pt) for ws in wb.Worksheets for pt in ws.PivotTables()]
        total_pts = len(all_pts)
        for i, (ws, pt) in enumerate(all_pts):
            try:
                pt.RefreshTable()
                print(f"Refreshed pivot table: {pt.Name} on {ws.Name}")
            except Exception as e:
                print(f"Error refreshing pivot table {pt.Name}: {e}")
            
            # Update progress continuously (~40% for pivot tables)
            if progress_callback and total_pts > 0:
                progress = 30 + (i + 1) / total_pts * 40
                progress_callback("update", progress)

        # --- Step 3: Save workbook (~20%) ---
        wb.Save()
        print("Workbook saved.")
        if progress_callback:
            progress_callback("update", 90)

        # --- Step 4: Finish (~10%) ---
        wb.Activate()
        excel.Visible = True
        excel.ScreenUpdating = True
        print("Excel refresh complete. Workbook remains open.")
        if progress_callback:
            progress_callback("update", 100)

        return excel, wb

    finally:
        if progress_callback:
            progress_callback("stop", 100)
