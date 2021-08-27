#Author-Oskari Tuormaa
#Description-

import adsk.core, adsk.fusion, adsk.cam, traceback
from .classes import Count

from typing import List, Dict

def run(context):
    ui = None
    try:
        app: adsk.core.Application = adsk.core.Application.get()
        ui: adsk.core.UserInterface  = app.userInterface
        product: adsk.core.Product = app.activeProduct
        design = adsk.fusion.Design.cast(product)

        fileDialog = ui.createFileDialog()
        fileDialog.title = "Select output directory"
        fileDialog.initialFilename = f"{design.parentDocument.name}.xlsx"
        fileDialog.filter = "XLSX File (*.xlsx)"
        fileDialog.showSave()
        output_dir = fileDialog.filename

        if not output_dir:
            return

        unique = Count()
        unique += design.allComponents
        unique.save_xlsx(output_dir)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
