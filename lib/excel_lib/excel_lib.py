import win32com.client


from ... import config
from .fusion_dataclasses import Body, Module

from exceltypes import Application, Workbook, Worksheet, ListObject, Range

BODY_TABLE_NAME = "IndividualParts"
MODULE_TABLE_NAME = "ModulesParts"
SHEET_NAME = "LinkFusion"

def open_excel_doc(excel_file_path: str) -> Workbook:
    """Opens Excel document at given path."""
    excel: Application = win32com.client.Dispatch('Excel.Application')
    excel.Visible = config.DEBUG
    excel.DisplayAlerts = False
    return excel.Workbooks.Open(excel_file_path)

def save_and_quit(workbook: Workbook, save_path: str):
    """Saves and closes given Excel workbook at given path."""
    workbook.SaveAs(save_path, ConflictResolution=2)
    workbook.Application.Quit()

def set_table_data(sheet: Worksheet, table: ListObject, data: list[list]):
    # TODO: Handle empty data
    data_n_cols = len(data[0]) if data else 0
    table_n_cols = table.Range.Columns.Count

    assert data_n_cols == table_n_cols, \
        f"Mismatched amount of columns in data and table: {data_n_cols=} != {table_n_cols=}"

    table.DataBodyRange.Clear()

    start_cell: Range = table.Range.Cells(1, 1)
    end_cell: Range = start_cell.Offset(len(data)+1, len(data[0]) if data else 0)
    new_range: Range = sheet.Range(start_cell, end_cell)

    table.Resize(new_range)
    table.DataBodyRange.Value = data

def write_bodies_to_table(workbook: Workbook, bodies: list[Body]):
    """Populates the Bodies table in the given workbook.

    @note Raises KeyError if either the expected sheet or the expected
          table cannot be found.
    
    @param workbook The workbook in which to populate the Bodies table.
    @param bodies A list containing the data to populate the table with.
    """
    # Convert list of bodies to 2d array
    body_table_data = []
    for body in bodies:
        body_table_data.append([body.name, body.count, body.material])

    sheet: Worksheet = workbook.Sheets(SHEET_NAME)
    table: ListObject = sheet.ListObjects(BODY_TABLE_NAME)
    set_table_data(sheet, table, body_table_data)


def write_modules_to_table(workbook: Workbook, modules: list[Module]):
    """Populates the Modules table in the given workbook.
    
    @note Raises KeyError if either the expected sheet or the expected
          table cannot be found.
    
    @param workbook The workbook in which to populate the Modules table.
    @param modules A list containing the data to populate the table with.
    """
    # Convert list of modules to 2d array
    module_table_data = []
    for id, module in enumerate(modules):
        for body in module.bodies:
            module_table_data.append([module.category, id+1, module.name, body.name, body.count])

    sheet: Worksheet = workbook.Sheets(SHEET_NAME)
    table: ListObject = sheet.ListObjects(MODULE_TABLE_NAME)
    set_table_data(sheet, table, module_table_data)