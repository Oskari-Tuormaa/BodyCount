import adsk.core
import adsk.fusion
import os

import json

from ...lib import fusionAddInUtils as futil
from ... import config

app = adsk.core.Application.get()
ui = app.userInterface

CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_define_strings'
CMD_NAME = "Define Strings"
CMD_DESCRIPTION = "Define strings" # TODO: Better description

IS_PROMOTED = True

WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'BodyCount'

ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

ATTR_GRP = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}'

def get_strings() -> list[str]:
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)

    if (attr := design.attributes.itemByName(ATTR_GRP, 'strings')):
        return json.loads(attr.value)
    return []

def set_strings(strings: list[str]):
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    design.attributes.add(ATTR_GRP, 'strings', json.dumps(strings))


def start():
    futil.log("Hello from define_strings")

    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESCRIPTION, ICON_FOLDER)

    futil.add_handler(cmd_def.commandCreated, command_created)

    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Create panel if it doesn't already exist
    if (panel := workspace.toolbarPanels.itemById(PANEL_ID)) is None:
        panel = workspace.toolbarPanels.add(PANEL_ID, "BodyCount")

    control = panel.controls.addCommand(cmd_def)
    control.isPromoted = IS_PROMOTED

def stop():
    futil.log("Goodbye from define_strings")

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
    futil.log("Creating define_strings!")

    inputs = args.command.commandInputs

    table = inputs.addTableCommandInput('stringsTable', 'Strings', 1, '1')

    addbutton = inputs.addBoolValueInput('addStringBtn', 'Add String', False, ICON_FOLDER)
    table.addToolbarCommandInput(addbutton)

    rmbutton = inputs.addBoolValueInput('rmStringBtn', 'Remove String', False, ICON_FOLDER)
    table.addToolbarCommandInput(rmbutton)

    strings = get_strings()

    for i, string in enumerate(strings):
        stringInput = inputs.addStringValueInput(f'string{i}', '', string)
        table.addCommandInput(stringInput, i, 0)

    selection = inputs.addSelectionInput('sel1', 'Selection', 'Select something')
    selection.addSelectionFilter('Occurrences')
    selection.setSelectionLimits(0, 0)

    futil.add_handler(args.command.execute, command_execute)
    futil.add_handler(args.command.inputChanged, command_input_changed)

def command_execute(args: adsk.core.CommandEventArgs):
    futil.log("Executing define_strings!")

    inputs = args.command.commandInputs
    table: adsk.core.TableCommandInput = inputs.itemById('stringsTable')

    nStrings = table.rowCount
    strings = [inputs.itemById(f'string{i}').value for i in range(nStrings)]
    set_strings(strings)

def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input: adsk.core.CommandInput = args.input
    inputs = args.inputs

    futil.log(f"Input changed define_strings {changed_input.id}")

    if changed_input.id == 'addStringBtn':
        table: adsk.core.TableCommandInput = inputs.itemById('stringsTable')
        n_strings = table.rowCount

        stringInput = inputs.addStringValueInput(f'string{n_strings}', '', 'String Name')
        table.addCommandInput(stringInput, n_strings, 0)

    elif changed_input.id == 'rmStringBtn':
        table: adsk.core.TableCommandInput = inputs.itemById('stringsTable')
        n_strings = table.rowCount
        if n_strings == 0:
            return
        table.deleteRow(n_strings - 1)
    
    elif changed_input.id == 'sel1':
        changed_input: adsk.core.SelectionCommandInput
        n_selected = changed_input.selectionCount

        to_keep = []
        for i in range(n_selected):
            selected_obj: adsk.fusion.Occurrence = changed_input.selection(i).entity
            
            should_remove = False
            next = selected_obj
            while (next := next.assemblyContext) and not should_remove:
                for j in range(n_selected):
                    other_obj: adsk.fusion.Occurrence = changed_input.selection(j).entity
                    if next == other_obj:
                        should_remove = True
                        break
            if not should_remove:
                to_keep.append(selected_obj)
        
        if len(to_keep) != n_selected:
            changed_input.clearSelection()
            for obj in to_keep:
                changed_input.addSelection(obj)