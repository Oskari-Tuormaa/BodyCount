# Try importing openpyxl, and install through pip if not found.
try:
    import openpyxl
except ModuleNotFoundError:
    python_path = sys.executable.split('/')
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])

from . import commands
from .lib import fusionAddInUtils as futil

import sys
import subprocess


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