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

ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

def start():
    futil.log("Hello from settings")

    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESCRIPTION, ICON_FOLDER)

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

    inputs.addTextBoxCommandInput('txt1', '', '<h3><center>User Data</center></h3>', 1, True)
    dropbox_path = inputs.addTextBoxCommandInput('dropbox_path', 'Path to Dropbox', '', 1, False)

    inputs.addBoolValueInput('select_folder', 'Select Folder', False)

    inputs.addSeparatorCommandInput('sep1')
    inputs.addTextBoxCommandInput('txt2', '', '<h3><center>Shared Data</center></h3>', 1, True)

    if user_data.shared_data_path is not None:
        dropbox_path.text = user_data.shared_data_path

    if user_data.shared_data_path is not None and Path(user_data.shared_data_path).exists():
        ...
    else:
        inputs.addTextBoxCommandInput('wrn1', '', '<i><center>Set <b>Path to Dropbox</b> to a valid path to access shared settings</center></i>', 1, True)
        inputs.addTextBoxCommandInput('wrn2', '', '<i><center>Note that command must be ok\'d and opened again for change to take effect.</center></i>', 2, True)

    futil.add_handler(args.command.execute, command_execute)
    futil.add_handler(args.command.inputChanged, input_changed)
    futil.add_handler(args.command.validateInputs, validate_inputs)

def validate_inputs(args: adsk.core.ValidateInputsEventArgs):
    inputs = args.inputs
    dropbox_path: adsk.core.TextBoxCommandInput = inputs.itemById('dropbox_path')
    path = Path(dropbox_path.text)
    args.areInputsValid = path.is_absolute() and path.is_dir()

def input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    if  changed_input.id == 'select_folder':
        folder_dialog = ui.createFolderDialog()
        folder_dialog.title = 'Select the DropBox root folder'
        result = folder_dialog.showDialog()

        if result == adsk.core.DialogResults.DialogOK:
            dropbox_path = args.inputs.itemById('dropbox_path')
            dropbox_path.text = folder_dialog.folder

def command_execute(args: adsk.core.CommandEventArgs):
    inputs = args.command.commandInputs
    dropbox_path: adsk.core.TextBoxCommandInput = inputs.itemById('dropbox_path')

    user_data = settings_lib.load_user_data()
    user_data.shared_data_path = dropbox_path.text
    settings_lib.save_user_data(user_data)