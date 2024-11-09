import sys
import subprocess

# Try importing openpyxl, and install through pip if not found.
OPENPYXL_VERSION = "3.1.5"
try:
    import openpyxl
except ModuleNotFoundError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', f'openpyxl=={OPENPYXL_VERSION}'])

from . import commands
from .lib import fusionAddInUtils as futil


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