#Author-Oskari Tuormaa
#Description-

from pathlib import Path
import traceback

import adsk.cam
import adsk.core
import adsk.fusion

from .extras.classes import BodyCount, PriceCount

# Setup proper path for installed packages
import importlib
import os, sys
packagepath = os.path.join(os.path.dirname(sys.argv[0]), 'Lib/site-packages/')
if packagepath not in sys.path:
    sys.path.append(packagepath)

# Install and import xlsxwriter
try:
    importlib.import_module("xlsxwriter")
except ImportError:
    import pip
    pip.main(["install", "xlsxwriter"])
finally:
    globals()["xw"] = importlib.import_module("xlsxwriter")

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
        wb = xw.Workbook(output_dir)
        ws_counts = wb.add_worksheet("Counts")
        ws_calc = wb.add_worksheet("Calculations")

        # Count bodies
        unique = BodyCount()
        unique += rootComp.occurrences.asList
        unique.add_to_ws(ws_counts)

        # Count prices
        prices = PriceCount()
        prices += rootComp.occurrences.asList
        prices.add_to_db(wb)

        # Write file
        # pylightxl.writexl(wb, output_dir)
        wb.close()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

        with open(Path(log_dir, "log.log"), "w+") as fd:
            fd.write(traceback.format_exc())
