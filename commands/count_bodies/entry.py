import adsk.core
import adsk.fusion
import os
import json

import openpyxl as op
from openpyxl.workbook import Workbook

from ...lib import fusionAddInUtils as futil
from ...lib import counting_lib, excel_lib, settings_lib
from ... import config

from pathlib import Path

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

    design = adsk.fusion.Design.cast(app.activeProduct)
    root = design.rootComponent

    while root.customGraphicsGroups.count != 0:
        root.customGraphicsGroups.item(0).deleteMe()

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
    save_dialog.filter = "Excel file (*.xls;*.xlsx;*.xlsm);;All files (*)"
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

    inputs = args.command.commandInputs

    shared_data = settings_lib.load_shared_data()
    file_data = settings_lib.load_file_data()
    modules = counting_lib.collect_modules_under(rootComp)

    path_table = inputs.addTableCommandInput('', '', 3, '3:6:2')
    excel_file_path = inputs.addStringValueInput('excel_path', '', str(file_data.excel_path))
    select_button = inputs.addBoolValueInput('select_folder', 'Select Folder', False)
    path_table.addCommandInput(inputs.addTextBoxCommandInput('', '', '<h4>Path to excel file</h4>', 1, True), 0, 0)
    path_table.addCommandInput(excel_file_path, 0, 1)
    path_table.addCommandInput(select_button, 0, 2)
    path_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle
    path_table.minimumVisibleRows = 1
    path_table.maximumVisibleRows = 1

    module_table = inputs.addTableCommandInput('modules', 'Modules', 3, '1:1:1')
    module_table.addCommandInput(inputs.addTextBoxCommandInput('', '', '<h3>Category</h3>', 1, True), 0, 0)
    module_table.addCommandInput(inputs.addTextBoxCommandInput('', '', '<h3>Detail material</h3>', 1, True), 0, 1)
    module_table.addCommandInput(inputs.addTextBoxCommandInput('', '', '<h3>Wood material</h3>', 1, True), 0, 2)

    groups = {module.category for module in modules}
    for i, group in enumerate(groups):
        name_inp = inputs.addTextBoxCommandInput('', '', group, 1, True)
        detail_dropdown = inputs.addDropDownCommandInput('', '', adsk.core.DropDownStyles.TextListDropDownStyle)
        wood_dropdown = inputs.addDropDownCommandInput('', '', adsk.core.DropDownStyles.TextListDropDownStyle)

        selected_detail = None
        selected_wood = None
        if group in file_data.modules:
            module_data = file_data.modules[group]
            selected_detail = module_data.detail_material
            selected_wood = module_data.wood_material

        for detail_material in shared_data.detail_materials:
            detail_dropdown.listItems.add(detail_material, detail_material == selected_detail)

        for wood_material in shared_data.wood_materials:
            wood_dropdown.listItems.add(wood_material, wood_material == selected_wood)

        module_table.addCommandInput(name_inp, i+1, 0)
        module_table.addCommandInput(detail_dropdown, i+1, 1)
        module_table.addCommandInput(wood_dropdown, i+1, 2)

    futil.add_handler(args.command.execute, command_execute)
    futil.add_handler(args.command.inputChanged, input_changed)
    futil.add_handler(args.command.validateInputs, validate_inputs)

def validate_inputs(args: adsk.core.ValidateInputsEventArgs):
    inputs = args.inputs
    wood_table: adsk.core.TableCommandInput = inputs.itemById('modules')
    file_data = settings_lib.load_file_data()

    dropdowns_filled = True
    for i in range(1, wood_table.rowCount):
        category_name_inp: adsk.core.TextBoxCommandInput = wood_table.getInputAtPosition(i, 0)
        detail_dropdown: adsk.core.DropDownCommandInput = wood_table.getInputAtPosition(i, 1)
        wood_dropdown: adsk.core.DropDownCommandInput = wood_table.getInputAtPosition(i, 2)

        category_name = category_name_inp.formattedText
        if category_name in file_data.modules:
            detail_material = detail_dropdown.selectedItem
            wood_material = wood_dropdown.selectedItem
            file_data.modules[category_name].detail_material = detail_material.name if detail_material is not None else None
            file_data.modules[category_name].wood_material = wood_material.name if wood_material is not None else None
        else:
            file_data.modules[category_name] = settings_lib.ModuleSettings(
                category_name=category_name,
                detail_material=detail_dropdown.selectedItem,
                wood_material=wood_dropdown.selectedItem
            )

        if detail_dropdown.selectedItem is None or wood_dropdown.selectedItem is None:
            dropdowns_filled = False

    excel_path_inp: adsk.core.StringValueCommandInput = inputs.itemById('excel_path')
    excel_path = Path(excel_path_inp.value)
    is_valid_excel_path = excel_path.is_absolute() and excel_path.is_file()

    file_data.excel_path = excel_path
    settings_lib.save_file_data(file_data)
    
    args.areInputsValid = dropdowns_filled and is_valid_excel_path


def input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    inputs = args.inputs

    if changed_input.id == 'select_folder':
        new_path = get_input_file_path()
        if new_path is None:
            return
        excel_path_input: adsk.core.StringValueCommandInput = inputs.itemById('excel_path')
        excel_path_input.value = new_path


def command_execute(args: adsk.core.CommandEventArgs):
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    rootComp = design.rootComponent

    inputs = args.command.commandInputs

    bodies  = counting_lib.collect_bodies_under(rootComp)
    modules = counting_lib.collect_modules_under(rootComp)

    excel_path_input: adsk.core.StringValueCommandInput = inputs.itemById('excel_path')
    excel_path = excel_path_input.value

    workbook = op.open(excel_path, keep_vba=True, keep_links=True, rich_text=True)

    excel_lib.write_bodies_to_table(workbook, bodies)
    excel_lib.write_modules_to_table(workbook, modules)

    # output_file_path = get_output_file_path()
    # if not output_file_path:
    #     return

    try_saving_workbook(workbook, excel_path)
