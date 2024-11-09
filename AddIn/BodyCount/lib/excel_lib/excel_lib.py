from openpyxl.workbook import Workbook

from .custom_table import CustomTable
from .fusion_dataclasses import Body, Module

BODY_TABLE_NAME = "IndividualParts"
MODULE_TABLE_NAME = "ModulesParts"
SHEET_NAME = "LinkFusion"

def write_bodies_to_table(workbook: Workbook, bodies: list[Body]):
    """Populates the Bodies table in the given workbook.

    @note Raises KeyError if either the expected sheet or the expected
          table cannot be found.
    
    @param workbook The workbook in which to populate the Bodies table.
    @param bodies A list containing the data to populate the table with.
    """
    sheet = workbook[SHEET_NAME]
    body_table = CustomTable(sheet, sheet.tables[BODY_TABLE_NAME])
    
    # Convert list of bodies to 2d array
    body_table_data = []
    for body in bodies:
        body_table_data.append([body.name, body.count, body.material])

    body_table.set_data(body_table_data)

def write_modules_to_table(workbook: Workbook, modules: list[Module]):
    """Populates the Modules table in the given workbook.
    
    @note Raises KeyError if either the expected sheet or the expected
          table cannot be found.
    
    @param workbook The workbook in which to populate the Modules table.
    @param modules A list containing the data to populate the table with.
    """
    sheet = workbook[SHEET_NAME]
    module_table = CustomTable(sheet, sheet.tables[MODULE_TABLE_NAME])

    # Convert list of modules to 2d array
    module_table_data = []
    for id, module in enumerate(modules):
        for body in module.bodies:
            module_table_data.append([module.category, id+1, module.name, body.name, body.count])

    module_table.set_data(module_table_data)