from datetime import datetime
import os
import pythoncom
import win32com.client


def export_excel_to_pdf(excel_file_path, base_pdf_dir, num_blocks):
    pythoncom.CoInitialize()
    excel = None
    wb_win = None

    try:
        now = datetime.now()
        month_name = now.strftime("%b").upper()
        year_str = now.strftime("%Y")

        target_dir = os.path.join(base_pdf_dir, f"NNN {month_name}-{year_str}")
        os.makedirs(target_dir, exist_ok=True)

        date_str = now.strftime("%d-%m-%Y")
        pdf_name = f"NNN {date_str}.pdf"
        abs_pdf_path = os.path.abspath(os.path.join(target_dir, pdf_name))

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        abs_excel_path = os.path.abspath(excel_file_path)
        wb_win = excel.Workbooks.Open(abs_excel_path)
        ws_win = wb_win.ActiveSheet

        max_row = ws_win.Cells(ws_win.Rows.Count, 8).End(-4162).Row
        if max_row < 1:
            max_row = 1

        col_h_values = ws_win.Range(
            ws_win.Cells(1, 8), ws_win.Cells(max_row, 8)
        ).Value

        start_row = 1
        found_count = 0
        target_count = num_blocks + 1

        if col_h_values:
            for r in range(max_row, 0, -1):
                val = (
                    col_h_values[r - 1][0]
                    if isinstance(col_h_values, tuple)
                    else col_h_values
                )
                if val is not None and str(val).strip() != "":
                    found_count += 1
                    if found_count == target_count:
                        start_row = r
                        break

        ws_win.PageSetup.PrintArea = f"A{start_row}:H{max_row}"
        page_setup = ws_win.PageSetup
        page_setup.LeftMargin = excel.CentimetersToPoints(0.64)
        page_setup.RightMargin = excel.CentimetersToPoints(0.64)
        page_setup.TopMargin = excel.CentimetersToPoints(1.91)
        page_setup.BottomMargin = excel.CentimetersToPoints(1.91)
        page_setup.HeaderMargin = excel.CentimetersToPoints(0.76)
        page_setup.FooterMargin = excel.CentimetersToPoints(0.76)
        page_setup.PaperSize = 9
        page_setup.Orientation = 1
        page_setup.Zoom = False
        page_setup.FitToPagesWide = 1
        page_setup.FitToPagesTall = 2

        ws_win.ExportAsFixedFormat(0, abs_pdf_path)
        return abs_pdf_path

    finally:
        if wb_win:
            try:
                wb_win.Close(False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()