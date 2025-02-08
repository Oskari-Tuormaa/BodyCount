import adsk.core
import adsk.fusion
import os
import re

from ...lib import fusionAddInUtils as futil
from ...lib import counting_lib, excel_lib, settings_lib
from ... import config

from pathlib import Path

from exceltypes import Workbook

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


def get_input_file_path() -> Path | None:
    open_dialog = ui.createFileDialog()
    open_dialog.title = "Select input excel file"
    open_dialog.isMultiSelectEnabled = False
    open_dialog.filter = "Excel file (*.xls;*.xlsx;*.xlsm);;All files (*)"
    if open_dialog.showOpen() == adsk.core.DialogResults.DialogOK:
        return Path(open_dialog.filename)


def command_created(args: adsk.core.CommandCreatedEventArgs):
    futil.log("creating count_bodies")

    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    rootComp = design.rootComponent

    args.command.setDialogMinimumSize(600, 100)

    inputs = args.command.commandInputs

    shared_data = settings_lib.load_shared_data()
    file_data = settings_lib.load_file_data()
    modules = counting_lib.collect_modules_under(rootComp)

    path_table = inputs.addTableCommandInput('', '', 3, '3:6:2')
    excel_file_path = inputs.addStringValueInput('excel_path', '', str(file_data.excel_path))
    select_button = inputs.addBoolValueInput('select_folder', 'Select Folder', False)
    path_table.addCommandInput(inputs.addTextBoxCommandInput('', '', '<b>Path to excel file</b>', 6, True), 0, 0)
    path_table.addCommandInput(excel_file_path, 0, 1)
    path_table.addCommandInput(select_button, 0, 2)
    path_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle
    path_table.minimumVisibleRows = 1
    path_table.maximumVisibleRows = 1

    module_table = inputs.addTableCommandInput('modules', 'Modules', 3, '1:1:1')
    module_table.addCommandInput(inputs.addTextBoxCommandInput('', '', '<b>Category</b>', 1, True), 0, 0)
    module_table.addCommandInput(inputs.addTextBoxCommandInput('', '', '<b>Detail material</b>', 1, True), 0, 1)
    module_table.addCommandInput(inputs.addTextBoxCommandInput('', '', '<b>Wood material</b>', 1, True), 0, 2)

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

    dropdowns_filled = True
    for i in range(1, wood_table.rowCount):
        detail_dropdown: adsk.core.DropDownCommandInput = wood_table.getInputAtPosition(i, 1)
        wood_dropdown: adsk.core.DropDownCommandInput = wood_table.getInputAtPosition(i, 2)

        if detail_dropdown.selectedItem is None or wood_dropdown.selectedItem is None:
            dropdowns_filled = False

    excel_path_inp: adsk.core.StringValueCommandInput = inputs.itemById('excel_path')
    excel_path = Path(excel_path_inp.value)
    is_valid_excel_path = excel_path.is_absolute() and excel_path.is_file()

    args.areInputsValid = dropdowns_filled and is_valid_excel_path


def update_file_data(inputs: adsk.core.CommandInputs):
    file_data = settings_lib.load_file_data()
    modules_table: adsk.core.TableCommandInput = inputs.itemById('modules')
    # TODO: Fix saving of file data

    for i in range(1, modules_table.rowCount):
        category_name_inp: adsk.core.TextBoxCommandInput = modules_table.getInputAtPosition(i, 0)
        detail_dropdown: adsk.core.DropDownCommandInput = modules_table.getInputAtPosition(i, 1)
        wood_dropdown: adsk.core.DropDownCommandInput = modules_table.getInputAtPosition(i, 2)

        detail_material = detail_dropdown.selectedItem.name if detail_dropdown.selectedItem is not None else None
        wood_material = wood_dropdown.selectedItem.name if wood_dropdown.selectedItem is not None else None

        category_name = category_name_inp.formattedText
        if category_name in file_data.modules:
            file_data.modules[category_name].detail_material = detail_material
            file_data.modules[category_name].wood_material = wood_material
        else:
            file_data.modules[category_name] = settings_lib.ModuleSettings(
                category_name=category_name,
                detail_material=detail_material,
                wood_material=wood_material
            )

    excel_path_inp: adsk.core.StringValueCommandInput = inputs.itemById('excel_path')
    file_data.excel_path = Path(excel_path_inp.value)
    settings_lib.save_file_data(file_data)


def input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    inputs = args.inputs

    if changed_input.id == 'select_folder':
        new_path = get_input_file_path()
        if new_path is None:
            return
        excel_path_input: adsk.core.StringValueCommandInput = inputs.itemById('excel_path')
        excel_path_input.value = str(new_path)


def command_execute(args: adsk.core.CommandEventArgs):
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    rootComp = design.rootComponent

    inputs = args.command.commandInputs
    
    update_file_data(inputs)

    modules_dict: dict[str, tuple[str, str]] = {}
    modules_table = inputs.itemById("modules")
    for i in range(1, modules_table.rowCount):
        category_name_inp: adsk.core.TextBoxCommandInput = modules_table.getInputAtPosition(i, 0)
        detail_dropdown: adsk.core.DropDownCommandInput = modules_table.getInputAtPosition(i, 1)
        wood_dropdown: adsk.core.DropDownCommandInput = modules_table.getInputAtPosition(i, 2)
        modules_dict[category_name_inp.formattedText] = (
            detail_dropdown.selectedItem.name,
            wood_dropdown.selectedItem.name
        )

    modules = counting_lib.collect_modules_under(rootComp)

    MATCH_WOOD_PATTERN = re.compile(r"^(9[7-9]\..*? )?([1-9]|12|1[4-6])\.")
    MATCH_DETAIL_PATTERN = re.compile(r"^(9[7-9]\..*? )?10\.")

    bodies_dict: dict[tuple[str, str], counting_lib.Body] = {}
    for module in modules:
        detail_type, wood_type = modules_dict[module.category]

        for body in module.bodies:
            material = body.material

            if MATCH_WOOD_PATTERN.match(body.name):
                material = wood_type
            elif MATCH_DETAIL_PATTERN.match(body.name):
                material = detail_type

            key = (body.name, material)

            if key not in bodies_dict:
                bodies_dict[key] = counting_lib.Body(
                    name=body.name,
                    count=0,
                    material=material,
                )
            bodies_dict[key].count += body.count
    
    bodies = list(bodies_dict.values())
    counting_lib.human_sort(bodies, key=lambda x: x.name)

    excel_path_input: adsk.core.StringValueCommandInput = inputs.itemById('excel_path')
    excel_path = excel_path_input.value

    try:
        workbook = excel_lib.open_excel_doc(excel_path)
    except Exception as e:
        if adsk.core.DialogResults.DialogCancel in e.args:
            return
        raise

    excel_lib.write_bodies_to_table(workbook, bodies)
    excel_lib.write_modules_to_table(workbook, modules)

    excel_lib.save(workbook, excel_path)
    excel_lib.close(workbook)
