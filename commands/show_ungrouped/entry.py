import adsk.core
import adsk.fusion
import os
import re

from dataclasses import dataclass
from typing import cast

from ...lib import fusionAddInUtils as futil
from ...lib import attributes_lib
from ...lib import counting_lib
from ...lib import custom_graphics_lib
from ... import config

app = adsk.core.Application.get()
ui = app.userInterface

CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_show_ungrouped'
CMD_NAME = "Show Ungrouped Components"
CMD_DESCRIPTION = "Show Ungrouped Components" # TODO: Better description

IS_PROMOTED = True

WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'BodyCount'

ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

ATTR_GRP = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}'

selection_graphics = custom_graphics_lib.SelectionGraphicsGroups()

def start():
    futil.log("Hello from show_ungrouped")

    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESCRIPTION, ICON_FOLDER)

    futil.add_handler(cmd_def.commandCreated, command_created)

    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Create panel if it doesn't already exist
    if (panel := workspace.toolbarPanels.itemById(PANEL_ID)) is None:
        panel = workspace.toolbarPanels.add(PANEL_ID, "BodyCount")

    control = panel.controls.addCommand(cmd_def)
    control.isPromoted = IS_PROMOTED

def stop():
    futil.log("Goodbye from show_ungrouped")

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
    futil.log("Creating show_ungrouped!")
    design = adsk.fusion.Design.cast(app.activeProduct)
    root = design.rootComponent

    selection_graphics.clear()
    g = selection_graphics.create('ungrouped')
    g.set_color(adsk.core.Color.create(255, 0, 0, 100))

    inputs = args.command.commandInputs
    table = inputs.addTableCommandInput('table', 'Ungrouped Items', 1, '1')
    i_tbox = 0

    modules = counting_lib.collect_modules_under(root)

    for top_lvl_occ in counting_lib.traverse_occurrences(root, depth=0):
        if top_lvl_occ.name.startswith('G_'):
            continue

        g.add_occ(top_lvl_occ)

        str_inp = inputs.addStringValueInput(f"tbox{i_tbox}", "Tbox", top_lvl_occ.name)
        str_inp.isReadOnly = True
        table.addCommandInput(str_inp, i_tbox, 0)
        i_tbox += 1

            
    if len(root.bRepBodies) > 0:
        g.add_objs(*root.bRepBodies)

        for obj in root.bRepBodies:
            str_inp = inputs.addStringValueInput(f"tbox{i_tbox}", "Tbox", obj.name)
            str_inp.isReadOnly = True
            table.addCommandInput(str_inp, i_tbox, 0)
            i_tbox += 1

    futil.add_handler(args.command.execute, command_execute)
    futil.add_handler(args.command.destroy, command_destroy)
    futil.add_handler(args.command.executePreview, command_execute_preview)

def command_execute(args: adsk.core.CommandEventArgs):
    futil.log("Executing show_ungrouped!")

    selection_graphics.clear()

def command_destroy(args: adsk.core.CommandEventArgs):
    selection_graphics.clear()

def command_execute_preview(args: adsk.core.CommandEventArgs):
    futil.log("Preview!")