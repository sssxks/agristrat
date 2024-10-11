from win32com.client import Dispatch

def export_to_excel(df, workbook_path, sheet_name):
    """
    Exports a pandas DataFrame to an Excel sheet using win32com.

    Parameters:
    df (pd.DataFrame): The DataFrame to export.
    workbook_path (str): The full path to the Excel workbook.
    sheet_name (str): The name of the sheet where data will be written.

    Returns:
    None
    """
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = 1  # Make Excel visible

    xlApp.Workbooks.Open(workbook_path)

    sheet = xlApp.ActiveWorkbook.Sheets(sheet_name)
    sheet.select
    sheet.cells.clearcontents  # Clear existing contents in the sheet

    y, x = 1 + df.shape[0], df.shape[1]

    sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, x)).Value = list(df.columns)
    sheet.Range(sheet.Cells(2, 1), sheet.Cells(y, x)).Value = list(df.values)

    xlApp.ActiveWorkbook.RefreshAll  # Refresh all (e.g., pivot tables)
    xlApp.ActiveWorkbook.Save  # Save the workbook
