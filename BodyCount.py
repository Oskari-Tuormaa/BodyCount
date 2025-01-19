import adsk.core
import importlib.metadata

import sys
import subprocess

from pathlib import Path

from .lib import fusionAddInUtils as futil

app = adsk.core.Application.get()
ui = app.userInterface

REQUIRED_PACKAGES = { 'openpyxl==3.1.0' }
INSTALLED_PACKAGES = {f'{pkg.metadata["Name"]}=={pkg.version}' for pkg in importlib.metadata.distributions()}
PACKAGES_TO_INSTALL = REQUIRED_PACKAGES - INSTALLED_PACKAGES

# Install Python packages if they're missing
try:
    if len(PACKAGES_TO_INSTALL) > 0:
        app = adsk.core.Application.get()
        ui = app.userInterface

        ui.messageBox("BodyCount needs to install some Python packages. This might cause a terminal window to open shortly. Please wait while packages install - it won't take too long.")

        python_path = Path(sys.executable)

        if python_path.stem != "python":
            python_path = python_path.parent/'Python'/'python.exe'

        subprocess.check_call([python_path, '-m', 'pip', 'install', *PACKAGES_TO_INSTALL])

        ui.messageBox("Package install done. Please restart Fusion.")
        sys.exit(0)
except:
    futil.handle_error('bodycount_setup')


from . import commands

def run(context):
    try:
        # This will run the start function in each of your commands as defined in commands/__init__.py
        commands.start()

    except:
        futil.handle_error('run')


def stop(context):
    try:
        # Remove all of the event handlers your app has created
        futil.clear_handlers()

        # This will run the start function in each of your commands as defined in commands/__init__.py
        commands.stop()

    except:
        futil.handle_error('stop')