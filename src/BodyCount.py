#Author-Oskari Tuormaa
#Description-

import adsk.core, adsk.fusion, adsk.cam, traceback
from .extras.classes import BodyCount, PriceCount
from .extras.packages import pylightxl

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
            return

        # Create DataBase and Worksheet
        db = pylightxl.Database()
        db.add_ws("Sheet1")
        ws: pylightxl.pylightxl.Worksheet = db.ws("Sheet1")

        # Count bodies
        unique = BodyCount()
        unique += rootComp.occurrences.asList
        unique.add_to_ws(ws)

        # Count prices
        prices = PriceCount()
        prices += rootComp.occurrences.asList
        prices.add_to_ws(ws)

        # Write file
        pylightxl.writexl(db, output_dir)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
