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

RESOURCE_FOLDER = Path(__file__).parent/'resources'/''
ICON_FOLDER = RESOURCE_FOLDER/'Icon'/''
ADD_FOLDER = RESOURCE_FOLDER/'Add'/''
REMOVE_FOLDER = RESOURCE_FOLDER/'Remove'/''

# To avoid ID collision between inputs, this counter is simply
# incremented and used as ID when adding new inputs.
counter = 0

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
    global counter
    counter = 0

    futil.log("Creating settings")

    user_data = settings_lib.load_user_data()

    inputs = args.command.commandInputs

    inputs.addTextBoxCommandInput('', '', '<h2><center>User Data</center></h2>', 1, True)
    inputs.addTextBoxCommandInput('', '', '<i><center>These settings are saved locally on your computer.</center></i>', 4, True)

    inputs.addTextBoxCommandInput('', '', '<h3>Path to Dropbox</h3>', 1, True)
    path_table = inputs.addTableCommandInput('', '', 2, '4:1')
    dropbox_path = inputs.addStringValueInput('dropbox_path', 'Path to Dropbox')
    select_button = inputs.addBoolValueInput('select_folder', 'Select Folder', False)
    path_table.addCommandInput(dropbox_path, 0, 0)
    path_table.addCommandInput(select_button, 0, 1)
    path_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle
    path_table.minimumVisibleRows = 1
    path_table.maximumVisibleRows = 1

    if user_data.shared_data_path is not None:
        dropbox_path.value = str(user_data.shared_data_path)

    # Error message for path validation
    error_message = inputs.addTextBoxCommandInput('path_error', '', '', 2, True)
    error_message.isFullWidth = True

    inputs.addSeparatorCommandInput('')
    inputs.addTextBoxCommandInput('', '', '<h2><center>Shared Data</center></h2>', 1, True)
    inputs.addTextBoxCommandInput('', '', '<i><center>These settings are saved in Dropbox.</center></i>', 4, True)

    if user_data.shared_data_path is not None and Path(user_data.shared_data_path).exists() and settings_lib.is_directory_writable(Path(user_data.shared_data_path)):
        shared_data = settings_lib.load_shared_data()

        # Wood types table
        inputs.addTextBoxCommandInput('', '', '<h3>Wood Types</h3>', 1, True)
        wood_table = inputs.addTableCommandInput('wood_types', 'Wood types', 1, '1')
        wood_table.maximumVisibleRows = 10
        wood_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle

        add_btn = inputs.addBoolValueInput('wood_add', 'Add wood type', False, str(ADD_FOLDER))
        remove_btn = inputs.addBoolValueInput('wood_remove', 'Remove wood type', False, str(REMOVE_FOLDER))
        wood_table.addToolbarCommandInput(add_btn)
        wood_table.addToolbarCommandInput(remove_btn)

        for i, wood_type in enumerate(shared_data.wood_materials):
            wood_table.addCommandInput(inputs.addStringValueInput(f'{counter}', '',  wood_type), i, 0)
            counter += 1

        # Detail materials table
        inputs.addTextBoxCommandInput('', '', '<h3>Details materials</h3>', 1, True)
        detail_table = inputs.addTableCommandInput('detail_materials', 'Detail materials', 1, '1')
        detail_table.maximumVisibleRows = 10
        detail_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle

        add_btn = inputs.addBoolValueInput('detail_add', 'Add detail material', False, str(ADD_FOLDER))
        remove_btn = inputs.addBoolValueInput('detail_remove', 'Remove detail material', False, str(REMOVE_FOLDER))
        detail_table.addToolbarCommandInput(add_btn)
        detail_table.addToolbarCommandInput(remove_btn)

        for i, detail_material in enumerate(shared_data.detail_materials):
            detail_table.addCommandInput(inputs.addStringValueInput(f'{counter}', '', detail_material), i, 0)
            counter += 1
        
        # Steel <-> brass numbers table
        inputs.addTextBoxCommandInput('', '', '<h3>Steel/brass part numbers</h3>', 1, True)
        steel_brass_header_table = inputs.addTableCommandInput('', '', 2, '1:1')
        steel_brass_header_table.minimumVisibleRows = 1
        steel_brass_header_table.maximumVisibleRows = 1
        steel_brass_header_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle
        steel_brass_header_table.addCommandInput(inputs.addTextBoxCommandInput('', '', 'Steel', 1, True), 0, 0)
        steel_brass_header_table.addCommandInput(inputs.addTextBoxCommandInput('', '', 'Brass', 1, True), 0, 1)

        steel_brass_table = inputs.addTableCommandInput('steel_brass', 'Steel brass numbers', 2, '1:1')
        steel_brass_table.maximumVisibleRows = 10
        steel_brass_table.tablePresentationStyle = adsk.core.TablePresentationStyles.transparentBackgroundTablePresentationStyle

        add_btn = inputs.addBoolValueInput('numbers_add', 'Add new part number pair', False, str(ADD_FOLDER))
        remove_btn = inputs.addBoolValueInput('numbers_remove', 'Remove part number pair', False, str(REMOVE_FOLDER))
        steel_brass_table.addToolbarCommandInput(add_btn)
        steel_brass_table.addToolbarCommandInput(remove_btn)

        for i, (steelnum, brassnum) in enumerate(shared_data.steel_brass_numbers):
            steel_brass_table.addCommandInput(
                inputs.addIntegerSpinnerCommandInput(f'{counter}', '', 0, 1000, 1, steelnum),
                i, 0
            )
            counter += 1
            steel_brass_table.addCommandInput(
                inputs.addIntegerSpinnerCommandInput(f'{counter}', '', 0, 1000, 1, brassnum),
                i, 1
            )
            counter += 1
    else:
        inputs.addTextBoxCommandInput('', '', '<i><center>Set <b>Path to Dropbox</b> to a valid path to access shared settings.<br>Note that command must be ok\'d and opened again for change to take effect.</center></i>', 10, True)

    args.command.setDialogMinimumSize(300, 200)

    futil.add_handler(args.command.execute, command_execute)
    futil.add_handler(args.command.inputChanged, input_changed)
    futil.add_handler(args.command.validateInputs, validate_inputs)

def update_path_error_message(inputs: adsk.core.CommandInputs) -> bool:
    """Update the path error message and return whether the path is valid.
    
    @param inputs The command inputs.
    @return True if the path is valid and writable, False otherwise.
    """
    dropbox_path = adsk.core.StringValueCommandInput.cast(inputs.itemById('dropbox_path'))
    error_message = adsk.core.TextBoxCommandInput.cast(inputs.itemById('path_error'))
    
    path = Path(dropbox_path.value)
    
    # Check if path is absolute and is a directory
    if not path.is_absolute() or not path.is_dir():
        error_message.formattedText = '<span style="color: red;">✗ Invalid path</span>'
        return False
    
    # Check if directory is writable
    if not settings_lib.is_directory_writable(path):
        error_message.formattedText = '<span style="color: red;">✗ Directory not writable (permission denied)</span>'
        return False
    
    error_message.formattedText = '<span style="color: green;">✓ Path is valid</span>'
    return True

def validate_inputs(args: adsk.core.ValidateInputsEventArgs):
    inputs = args.inputs
    args.areInputsValid = update_path_error_message(inputs)

def input_changed(args: adsk.core.InputChangedEventArgs):
    # futil.log('input_changed')
    global counter
    changed_input = args.input
    inputs = args.inputs

    if changed_input.id == 'select_folder':
        folder_dialog = ui.createFolderDialog()
        folder_dialog.title = 'Select the DropBox root folder'
        result = folder_dialog.showDialog()

        if result == adsk.core.DialogResults.DialogOK:
            dropbox_path = adsk.core.StringValueCommandInput.cast(inputs.itemById('dropbox_path'))
            dropbox_path.value = folder_dialog.folder
            # Validation will be triggered automatically after path changes

    elif changed_input.id == 'wood_add':
        wood_table = adsk.core.TableCommandInput.cast(args.inputs.itemById('wood_types'))
        wood_table.addCommandInput(
            inputs.addStringValueInput(f'{counter}', ''), wood_table.rowCount, 0
        )
        counter += 1

    elif changed_input.id == 'wood_remove':
        wood_table = adsk.core.TableCommandInput.cast(args.inputs.itemById('wood_types'))
        if wood_table.rowCount == 0:
            return
        if (i := wood_table.selectedRow) == -1 or i >= wood_table.rowCount:
            i = wood_table.rowCount - 1
        wood_table.deleteRow(i)
        wood_table.selectedRow = -1

    elif changed_input.id == 'detail_add':
        detail_table = adsk.core.TableCommandInput.cast(args.inputs.itemById('detail_materials'))
        detail_table.addCommandInput(
            inputs.addStringValueInput(f'{counter}', ''), detail_table.rowCount, 0
        )
        counter += 1

    elif changed_input.id == 'detail_remove':
        detail_table = adsk.core.TableCommandInput.cast(args.inputs.itemById('detail_materials'))
        if detail_table.rowCount == 0:
            return
        if (i := detail_table.selectedRow) == -1 or i >= detail_table.rowCount:
            i = detail_table.rowCount - 1
        detail_table.deleteRow(i)
        detail_table.selectedRow = -1

    elif changed_input.id == 'numbers_add':
        steel_brass_table = adsk.core.TableCommandInput.cast(args.inputs.itemById('steel_brass'))
        row = steel_brass_table.rowCount
        steel_brass_table.addCommandInput(
            inputs.addIntegerSpinnerCommandInput(f'{counter}', '', 0, 1000, 1, 0), row, 0
        )
        counter += 1
        steel_brass_table.addCommandInput(
            inputs.addIntegerSpinnerCommandInput(f'{counter}', '', 0, 1000, 1, 0), row, 1
        )
        counter += 1

    elif changed_input.id == 'numbers_remove':
        steel_brass_table = adsk.core.TableCommandInput.cast(args.inputs.itemById('steel_brass'))
        if steel_brass_table.rowCount == 1:
            # Only column headers in table = empty table
            return
        if (i := steel_brass_table.selectedRow) == -1 or i >= steel_brass_table.rowCount:
            i = steel_brass_table.rowCount - 1
        steel_brass_table.deleteRow(i)
        steel_brass_table.selectedRow = -1

def command_execute(args: adsk.core.CommandEventArgs):
    inputs = args.command.commandInputs
    dropbox_path = adsk.core.StringValueCommandInput.cast(inputs.itemById('dropbox_path'))

    user_data = settings_lib.load_user_data()
    user_data.shared_data_path = Path(dropbox_path.value)
    settings_lib.save_user_data(user_data)

    try:
        shared_data = settings_lib.load_shared_data()
    except RuntimeError as e:
        ui.messageBox("Please set a valid Dropbox path", "Settings Error")
        return
    except (PermissionError, OSError) as e:
        ui.messageBox(
            f"Cannot write to directory: {str(e)}\n\nPlease select a different directory.",
            "Permission Error"
        )
        return

    cast_str = lambda x: adsk.core.StringValueCommandInput.cast(x)
    cast_int = lambda x: adsk.core.IntegerSpinnerCommandInput.cast(x)

    wood_table = adsk.core.TableCommandInput.cast(inputs.itemById('wood_types'))
    if wood_table is not None:
        shared_data.wood_materials = list(
            cast_str(wood_table.getInputAtPosition(i, 0)).value
            for i in range(wood_table.rowCount)
        )

    detail_table = adsk.core.TableCommandInput.cast(inputs.itemById('detail_materials'))
    if detail_table is not None:
        shared_data.detail_materials = list(
            cast_str(detail_table.getInputAtPosition(i, 0)).value
            for i in range(detail_table.rowCount)
        )

    steel_brass_table = adsk.core.TableCommandInput.cast(inputs.itemById('steel_brass'))
    if steel_brass_table is not None:
        shared_data.steel_brass_numbers = list(
            (cast_int(steel_brass_table.getInputAtPosition(i, 0)).value,
             cast_int(steel_brass_table.getInputAtPosition(i, 1)).value)
            for i in range(steel_brass_table.rowCount)
        )

    settings_lib.save_shared_data(shared_data)