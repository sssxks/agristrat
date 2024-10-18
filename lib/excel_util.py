from win32com.client import Dispatch
import re


def excel_cell_to_indices(cell_ref):
    """
    Convert an Excel-style cell reference (e.g., 'B2') to row and column indices.

    Parameters:
    cell_ref (str): The Excel-style cell reference (e.g., 'B2').

    Returns:
    tuple: A tuple containing the row and column indices (row, col).
    """
    match = re.match(r"([A-Z]+)([0-9]+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")

    col_str, row_str = match.groups()

    # Convert column letters to a number (e.g., 'A' -> 1, 'B' -> 2, ..., 'AA' -> 27)
    col_num = 0
    for char in col_str:
        col_num = col_num * 26 + (ord(char.upper()) - ord("A") + 1)

    row_num = int(row_str)

    return row_num, col_num


def export_to_excel(
    df,
    workbook_path,
    sheet_name,
    include_columns=True,
    include_index=False,
    start_position="A1",
):
    """
    Exports a pandas DataFrame to an Excel sheet using win32com.
    会帮你打开Excel，如果你没打开的话。

    Parameters:
    df (pd.DataFrame): The DataFrame to export.
    workbook_path (str): The full path to the Excel workbook.
    sheet_name (str): The name of the sheet where data will be written.
    include_columns (bool): Whether to include the DataFrame's column headers. Default is True.
    include_index (bool): Whether to include the DataFrame's row indices. Default is False.
    start_position (str): The starting cell in Excel (e.g., 'B2'). Default is 'A1'.

    Returns:
    None
    """
    # Convert the start_position (e.g., 'B2') to row and column indices
    start_row, start_col = excel_cell_to_indices(start_position)

    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = 1  # Make Excel visible

    xlApp.Workbooks.Open(workbook_path)

    sheet = xlApp.ActiveWorkbook.Sheets(sheet_name)
    sheet.select
    sheet.cells.clearcontents  # Clear existing contents in the sheet

    # Adjust the DataFrame based on whether to include index
    if include_index:
        df_to_export = df.reset_index()
    else:
        df_to_export = df

    # Determine the number of rows and columns
    num_rows, num_cols = df_to_export.shape

    # Write column headers if specified
    if include_columns:
        col_headers = list(df_to_export.columns)
        sheet.Range(
            sheet.Cells(start_row, start_col),
            sheet.Cells(start_row, start_col + num_cols - 1),
        ).Value = col_headers
        start_row += 1  # Move to the next row for data

    # Write the DataFrame values
    sheet.Range(
        sheet.Cells(start_row, start_col),
        sheet.Cells(start_row + num_rows - 1, start_col + num_cols - 1),
    ).Value = list(df_to_export.values)

    xlApp.ActiveWorkbook.RefreshAll()  # Refresh all (e.g., pivot tables)
    xlApp.ActiveWorkbook.Save()  # Save the workbook
