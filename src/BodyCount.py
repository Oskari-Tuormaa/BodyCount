#Author-Oskari Tuormaa
#Description-

from pathlib import Path
import traceback

import adsk.cam
import adsk.core
import adsk.fusion

from .extras.classes import BodyCount, PriceCount
from .extras.packages import pylightxl

log_dir = Path(Path(__file__).parent, "logs")

def run(context):
    ui = None
    try:
        app: adsk.core.Application = adsk.core.Application.get()
        ui: adsk.core.UserInterface | None = app.userInterface
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

        # Create log directory
        log_dir.mkdir(exist_ok=True)

        # Create DataBase and Worksheet
        db = pylightxl.Database()
        db.add_ws("Counts")
        db.add_ws("Calculations")
        ws: pylightxl.pylightxl.Worksheet = db.ws("Counts")

        # Count bodies
        unique = BodyCount()
        unique += rootComp.occurrences.asList
        unique.add_to_ws(ws)

        # Count prices
        prices = PriceCount()
        prices += rootComp.occurrences.asList
        prices.add_to_db(db)

        # Write file
        pylightxl.writexl(db, output_dir)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

        with open(Path(log_dir, "log.log"), "w+") as fd:
            fd.write(traceback.format_exc())
