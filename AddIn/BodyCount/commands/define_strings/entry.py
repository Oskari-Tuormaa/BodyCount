import adsk.core
import adsk.fusion
import os
import re

from dataclasses import dataclass
from typing import cast

from ...lib import fusionAddInUtils as futil
from ...lib import attributes_lib
from ...lib import custom_graphics_lib
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
PLUS_ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', 'Plus', '')
MINUS_ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', 'Minus', '')

ATTR_GRP = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}'

selection_graphics = custom_graphics_lib.SelectionGraphicsGroups()

@dataclass
class CommandData:
    next_idx: int = 0
    currently_selected_idx: int | None = None
    currently_selected_string: str | None = None
    currently_selected_string_name: str | None = None
cmd_data = CommandData()

@dataclass
class Selection:
    selected_items: list[adsk.fusion.Occurrence]
selections: dict[int, Selection] = {}

MATCH_STRING_TXT_INPUT = re.compile("^string(\d+)$")


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

    selection_graphics.clear()

    inputs = args.command.commandInputs

    selector = inputs.addSelectionInput("selector", "", "")
    selector.isVisible = False
    selector.addSelectionFilter("Occurrences")
    selector.setSelectionLimits(0, 0)

    table = inputs.addTableCommandInput('stringsTable', 'Strings', 1, '1')

    addbutton = inputs.addBoolValueInput('addStringBtn', 'Add String', False, PLUS_ICON_FOLDER)
    table.addToolbarCommandInput(addbutton)

    rmbutton = inputs.addBoolValueInput('rmStringBtn', 'Remove String', False, MINUS_ICON_FOLDER)
    table.addToolbarCommandInput(rmbutton)

    strings = attributes_lib.get_file_data().strings
    for i, string in enumerate(strings):
        selection_graphics.create(f"grp{i}")
        stringInput = inputs.addStringValueInput(f'string{i}', '', string)
        table.addCommandInput(stringInput, i, 0)
        selections[i] = Selection([])
    cmd_data.next_idx = len(strings)

    futil.add_handler(args.command.execute, command_execute)
    futil.add_handler(args.command.destroy, command_destroy)
    futil.add_handler(args.command.inputChanged, command_input_changed)
    futil.add_handler(args.command.executePreview, command_execute_preview)

def command_execute(args: adsk.core.CommandEventArgs):
    futil.log("Executing define_strings!")

    inputs = args.command.commandInputs
    table: adsk.core.TableCommandInput = inputs.itemById('stringsTable')

    nStrings = table.rowCount
    strings = [table.getInputAtPosition(i, 0).value for i in range(nStrings)]
    file_data = attributes_lib.get_file_data()
    file_data.strings = strings
    attributes_lib.set_file_data(file_data)

    selection_graphics.clear()

def command_destroy(args: adsk.core.CommandEventArgs):
    selection_graphics.clear()

def command_execute_preview(args: adsk.core.CommandEventArgs):
    futil.log("Preview!")

    selection_graphics.clear()
    for i, sel in selections.items():
        graphics_group = selection_graphics.create(i)
        for obj in sel.selected_items:
            graphics_group.add_obj(obj)

def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = cast(adsk.core.CommandInput, args.input)
    inputs = args.inputs

    futil.log(f"Input changed define_strings {changed_input.id}")

    if changed_input.id == 'addStringBtn':
        table: adsk.core.TableCommandInput = inputs.itemById('stringsTable')
        n_strings = table.rowCount

        selection_graphics.create(f"grp{cmd_data.next_idx}")
        stringInput = inputs.addStringValueInput(f'string{cmd_data.next_idx}', '', 'String Name')
        table.addCommandInput(stringInput, n_strings, 0)
        selections[cmd_data.next_idx] = Selection([])

        cmd_data.next_idx += 1

    elif changed_input.id == 'rmStringBtn':
        table: adsk.core.TableCommandInput = inputs.itemById('stringsTable')
        n_strings = table.rowCount
        if n_strings == 0:
            return
        if (idx := table.selectedRow) != -1:
            table.deleteRow(idx)
        else:
            table.deleteRow(n_strings - 1)
    
    elif (match := MATCH_STRING_TXT_INPUT.match(changed_input.id)):
        changed_input = cast(adsk.core.StringValueCommandInput, changed_input)

        if changed_input.id != cmd_data.currently_selected_string:
            cmd_data.currently_selected_idx = int(match.group(1))
            cmd_data.currently_selected_string = changed_input.id
            cmd_data.currently_selected_string_name = changed_input.value
            return

    elif changed_input.id == "selector":
        changed_input = cast(adsk.core.SelectionCommandInput, changed_input)

        if cmd_data.currently_selected_idx is not None:
            obj = changed_input.selection(0).entity
            sel = selections[cmd_data.currently_selected_idx]
            if obj in sel.selected_items:
                sel.selected_items.remove(obj)
            else:
                sel.selected_items.append(obj)

        changed_input.clearSelection()
        app.activeViewport.refresh()