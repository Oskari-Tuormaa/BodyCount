import adsk.core
import adsk.fusion
import os

from pathlib import Path

from ...lib import fusionAddInUtils as futil
from ...lib import settings_lib
from ... import config

app = adsk.core.Application.get()
ui = app.userInterface

CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_settings'
CMD_NAME = "Settings"
CMD_DESCRIPTION = "Settings for the KitchenX AddIn" # TODO: Better description

IS_PROMOTED = False

WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'BodyCount'

# RESOURCE_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', 'Icon', '')
RESOURCE_FOLDER = Path(__file__).parent/'resources'/''
ICON_FOLDER = RESOURCE_FOLDER/'Icon'/''
ADD_FOLDER = RESOURCE_FOLDER/'Add'/''
REMOVE_FOLDER = RESOURCE_FOLDER/'Remove'/''

def start():
    futil.log("Hello from settings")

    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESCRIPTION, str(ICON_FOLDER))

    futil.add_handler(cmd_def.commandCreated, command_created)

    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Create panel if it doesn't already exist
    if (panel := workspace.toolbarPanels.itemById(PANEL_ID)) is None:
        panel = workspace.toolbarPanels.add(PANEL_ID, "BodyCount")

    control = panel.controls.addCommand(cmd_def)
    control.isPromoted = IS_PROMOTED

def stop():
    futil.log("Goodbye from settings")

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

def command_created(args: adsk.core.CommandCreatedEventArgs):
    futil.log("Creating settings")

    user_data = settings_lib.load_user_data()

    inputs = args.command.commandInputs

    inputs.addTextBoxCommandInput('', '', '<h3><center>User Data</center></h3>', 1, True)
    inputs.addTextBoxCommandInput('', '', '<i><center>These settings are saved locally on your computer.</center></i>', 1, True)

    dropbox_path = inputs.addStringValueInput('dropbox_path', 'Path to Dropbox')
    inputs.addBoolValueInput('select_folder', 'Select Folder', False)

    if user_data.shared_data_path is not None:
        dropbox_path.value = user_data.shared_data_path

    overwrite_check = inputs.addBoolValueInput('overwrite', 'Overwrite excel document', True)
    overwrite_check.value = user_data.overwrite

    inputs.addSeparatorCommandInput('')
    inputs.addTextBoxCommandInput('', '', '<h3><center>Shared Data</center></h3>', 1, True)
    inputs.addTextBoxCommandInput('', '', '<i><center>These settings are saved in Dropbox.</center></i>', 1, True)

    if user_data.shared_data_path is not None and Path(user_data.shared_data_path).exists():
        shared_data = settings_lib.load_shared_data()

        # Wood types table
        wood_table = inputs.addTableCommandInput('wood_types', 'Wood types', 1, '1')
        wood_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle
        wood_title = inputs.addTextBoxCommandInput('', '', '<h4>Wood Types</h4>', 1, True)
        wood_table.addCommandInput(wood_title, 0, 0)

        add_btn = inputs.addBoolValueInput('wood_add', 'Add wood type', False, str(ADD_FOLDER))
        remove_btn = inputs.addBoolValueInput('wood_remove', 'Remove wood type', False, str(REMOVE_FOLDER))
        wood_table.addToolbarCommandInput(add_btn)
        wood_table.addToolbarCommandInput(remove_btn)

        for i, wood_type in enumerate(shared_data.wood_materials):
            wood_table.addCommandInput(
                inputs.addStringValueInput('', '',  wood_type),
                i+1, 0
            )

        # Detail materials table
        detail_table = inputs.addTableCommandInput('detail_materials', 'Detail materials', 1, '1')
        detail_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle
        detail_title = inputs.addTextBoxCommandInput('', '', '<h4>Details materials</h4>', 1, True)
        detail_table.addCommandInput(detail_title, 0, 0)

        add_btn = inputs.addBoolValueInput('detail_add', 'Add detail material', False, str(ADD_FOLDER))
        remove_btn = inputs.addBoolValueInput('detail_remove', 'Remove detail material', False, str(REMOVE_FOLDER))
        detail_table.addToolbarCommandInput(add_btn)
        detail_table.addToolbarCommandInput(remove_btn)

        for i, detail_material in enumerate(shared_data.detail_materials):
            detail_table.addCommandInput(
                inputs.addStringValueInput('', '', detail_material),
                i+1, 0
            )
    else:
        inputs.addTextBoxCommandInput('', '', '<i><center>Set <b>Path to Dropbox</b> to a valid path to access shared settings.<br>Note that command must be ok\'d and opened again for change to take effect.</center></i>', 2, True)

    futil.add_handler(args.command.execute, command_execute)
    futil.add_handler(args.command.inputChanged, input_changed)
    futil.add_handler(args.command.validateInputs, validate_inputs)

def validate_inputs(args: adsk.core.ValidateInputsEventArgs):
    # futil.log('validate')
    inputs = args.inputs

    dropbox_path: adsk.core.StringValueCommandInput = inputs.itemById('dropbox_path')
    path = Path(dropbox_path.value)
    args.areInputsValid = path.is_absolute() and path.is_dir()

def input_changed(args: adsk.core.InputChangedEventArgs):
    # futil.log('input_changed')
    changed_input = args.input
    inputs = args.inputs

    if changed_input.id == 'select_folder':
        folder_dialog = ui.createFolderDialog()
        folder_dialog.title = 'Select the DropBox root folder'
        result = folder_dialog.showDialog()

        if result == adsk.core.DialogResults.DialogOK:
            dropbox_path: adsk.core.StringValueCommandInput = inputs.itemById('dropbox_path')
            dropbox_path.value = folder_dialog.folder

    elif changed_input.id == 'wood_add':
        wood_table: adsk.core.TableCommandInput = args.inputs.itemById('wood_types')
        wood_table.addCommandInput(
            inputs.addStringValueInput('', ''), wood_table.rowCount, 0
        )

    elif changed_input.id == 'wood_remove':
        wood_table: adsk.core.TableCommandInput = args.inputs.itemById('wood_types')
        if wood_table.rowCount == 1:
            # Only title in table = empty table
            return
        if (i := wood_table.selectedRow) == -1 or i >= wood_table.rowCount:
            i = wood_table.rowCount - 1
        wood_table.deleteRow(i)
        if wood_table.selectedRow == 0:
            wood_table.selectedRow = 1

    elif changed_input.id == 'detail_add':
        detail_table: adsk.core.TableCommandInput = args.inputs.itemById('detail_materials')
        detail_table.addCommandInput(
            inputs.addStringValueInput('', ''), detail_table.rowCount, 0
        )

    elif changed_input.id == 'detail_remove':
        detail_table: adsk.core.TableCommandInput = args.inputs.itemById('detail_materials')
        if detail_table.rowCount == 1:
            # Only title in table = empty table
            return
        if (i := detail_table.selectedRow) == -1 or i >= detail_table.rowCount:
            i = detail_table.rowCount - 1
        detail_table.deleteRow(i)
        if detail_table.selectedRow == 0:
            detail_table.selectedRow = 1

def command_execute(args: adsk.core.CommandEventArgs):
    inputs = args.command.commandInputs
    dropbox_path: adsk.core.StringValueCommandInput = inputs.itemById('dropbox_path')
    overwrite: adsk.core.BoolValueCommandInput = inputs.itemById('overwrite')

    user_data = settings_lib.load_user_data()
    user_data.shared_data_path = dropbox_path.value
    user_data.overwrite = overwrite.value
    settings_lib.save_user_data(user_data)

    try:
        shared_data = settings_lib.load_shared_data()
    except RuntimeError:
        return

    wood_table: adsk.core.TableCommandInput | None = inputs.itemById('wood_types')
    if wood_table is not None:
        shared_data.wood_materials = list(
            wood_table.getInputAtPosition(i, 0).value
            for i in range(1, wood_table.rowCount)
        )

    detail_table: adsk.core.TableCommandInput | None = inputs.itemById('detail_materials')
    if detail_table is not None:
        shared_data.detail_materials = list(
            detail_table.getInputAtPosition(i, 0).value
            for i in range(1, detail_table.rowCount)
        )

    settings_lib.save_shared_data(shared_data)