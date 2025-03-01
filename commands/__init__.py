# Here you define the commands that will be added to your add-in.

# If you want to add an additional command, duplicate one of the existing directories and import it here.
# You need to use aliases (import "entry" as "my_module") assuming you have the default module named "entry".
from .show_ungrouped import entry as show_ungrouped
from .count_bodies import entry as count_bodies
from .settings import entry as settings

# Fusion will automatically call the start() and stop() functions.
commands = [
    show_ungrouped,
    count_bodies,
    settings
]


# Assumes you defined a "start" function in each of your modules.
# The start function will be run when the add-in is started.
def start():
    for command in commands:
        command.start()


# Assumes you defined a "stop" function in each of your modules.
# The stop function will be run when the add-in is stopped.
def stop():
    for command in commands:
        command.stop()