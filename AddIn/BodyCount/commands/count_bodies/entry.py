import adsk.core
import adsk.fusion
import os
import json

import openpyxl as op
from openpyxl.workbook import Workbook

from ...lib import fusionAddInUtils as futil
from ...lib import counting_lib
from ...lib import excel_lib
from ... import config

app = adsk.core.Application.get()
ui = app.userInterface

CMD_ID = f"{config.COMPANY_NAME}_{config.ADDIN_NAME}_count_bodies"
CMD_NAME = "Count Bodies"
CMD_DESCRIPTION = "Count Bodies"  # TODO: Better description

IS_PROMOTED = True

WORKSPACE_ID = "FusionSolidEnvironment"
PANEL_ID = "BodyCount"

ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources", "")

ATTR_GRP = f"{config.COMPANY_NAME}_{config.ADDIN_NAME}"


def start():
    futil.log("Hello from count_bodies")

    cmd_def = ui.commandDefinitions.addButtonDefinition(
        CMD_ID, CMD_NAME, CMD_DESCRIPTION, ICON_FOLDER
    )

    futil.add_handler(cmd_def.commandCreated, command_created)

    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Create panel if it doesn't already exist
    if (panel := workspace.toolbarPanels.itemById(PANEL_ID)) is None:
        panel = workspace.toolbarPanels.add(PANEL_ID, "BodyCount")

    control = panel.controls.addCommand(cmd_def)
    control.isPromoted = IS_PROMOTED


def stop():
    futil.log("Goodbye from count_bodies")

    cmd_def = ui.commandDefinitions.itemById(CMD_ID)

    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    control = panel.controls.itemById(CMD_ID)

    if cmd_def:
        cmd_def.deleteMe()

    if control:
        control.deleteMe()

    # Delete panel if this was the last command
    if panel and len(panel.controls) == 0:
        panel.deleteMe()


def get_input_file_path() -> str | None:
    open_dialog = ui.createFileDialog()
    open_dialog.title = "Select input excel file"
    open_dialog.isMultiSelectEnabled = False
    open_dialog.filter = "Excel file (*.xls;*.xlsx;*.xlsm);;All files (*)"
    open_dialog.showOpen()
    return open_dialog.filename


def get_output_file_path() -> str | None:
    save_dialog = ui.createFileDialog()
    save_dialog.title = "Select output excel file"
    save_dialog.showSave()
    return save_dialog.filename


def try_saving_workbook(workbook: Workbook, path: str):
    while True:
        try:
            workbook.save(path)
            return
        except PermissionError:
            button = ui.messageBox(
                "The file is probably open in excel. Please close the file and try again.",
                "Permission error",
                adsk.core.MessageBoxButtonTypes.RetryCancelButtonType,
            )

            # If user pressed Cancel, we stop the infinite loop
            # by returning.
            if button == adsk.core.DialogResults.DialogCancel:
                return


def command_created(args: adsk.core.CommandCreatedEventArgs):
    futil.log("creating count_bodies")

    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    rootComp = design.rootComponent

    bodies = counting_lib.collect_bodies_under(rootComp)
    modules = counting_lib.collect_modules_under(rootComp)

    input_file_path = get_input_file_path()
    if not input_file_path:
        return

    workbook = op.open(input_file_path, keep_vba=True, keep_links=True, rich_text=True)

    excel_lib.write_bodies_to_table(workbook, bodies)
    excel_lib.write_modules_to_table(workbook, modules)

    output_file_path = get_output_file_path()
    if not output_file_path:
        return

    try_saving_workbook(workbook, output_file_path)

    futil.log("Done")
