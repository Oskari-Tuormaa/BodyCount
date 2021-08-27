#Author-
#Description-

import adsk.core, adsk.fusion, adsk.cam, traceback
import importlib
import os
import re
from .packages import pylightxl

from typing import List, Dict


def count_unique(arr: List[object]) -> Dict[object, int]:
    """ Counts each unique occurrence in input list

    Args:
        arr (List[str]): Array to count unique occurrences in

    Returns: Dictionary containing the individual counts of
        each unique value in input array

    """
    res = dict()
    reg = re.compile("^.*?(?=$|\(.*?\))")

    #names = [x.name for x in arr if x.isVisible]
    #names.sort()

    for val in arr:
        if not val.isVisible:
            continue
        name = reg.findall(val.name)[0].strip()
        if name in res:
            res[name]["count"] += 1
        else:
            res[name] = dict(
                count = 1,
                material = val.material.name,
                mass = val.physicalProperties.mass,
            )
    return res

def save_xlsx(to_save: Dict[str, int], path: str):
    """ Saves dictionary to .xlsx format

    Args:
        to_save (Dict[str, int]): Dictionary to save
        path (str): Path in which to save the .xlsx file

    Returns: None

    """
    db = pylightxl.Database()
    db.add_ws("Sheet1")
    ws: pylightxl.pylightxl.Worksheet = db.ws("Sheet1")

    ws.update_index(2, 2, "Bodies")
    ws.update_index(2, 3, "Counts")
    ws.update_index(2, 4, "Material")
    ws.update_index(2, 5, "Mass")

    keys = list(to_save.keys())
    keys.sort()
    for idx, k in enumerate(keys):
        v = to_save[k]
        ws.update_index(3+idx, 2, k)
        ws.update_index(3+idx, 3, v["count"])
        ws.update_index(3+idx, 4, v["material"])
        ws.update_index(3+idx, 5, v["mass"]*v["count"])

    pylightxl.writexl(db, path)

def run(context):
    ui = None
    try:
        app: adsk.core.Application = adsk.core.Application.get()
        ui: adsk.core.UserInterface  = app.userInterface
        product: adsk.core.Product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        rootComp = design.rootComponent

        fileDialog = ui.createFileDialog()
        fileDialog.title = "Select output directory"
        fileDialog.initialFilename = f"{design.parentDocument.name}.xlsx"
        fileDialog.filter = "XLSX File (*.xlsx)"
        fileDialog.showSave()
        output_dir = fileDialog.filename

        if not output_dir:
            ui.messageBox("Failed: No output file specified")
            return

        bodies = rootComp.bRepBodies
        unique = count_unique(bodies)

        save_xlsx(unique, output_dir)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
