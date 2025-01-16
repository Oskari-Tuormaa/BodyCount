
# TODO

- Add per-user settings.
    - Add "Dropbox path" setting variable, that should point to Vermland folder in Dropbox.
- Add Vermland shared settings.
    - Can hold valid detail and wood materials, as well as defaults.
    - Path to shared settings is based on "Dropbox path": <dropbox_path/Vermland>/The Collection/BodyCount.
- Add command input box to "Count bodies" command.
    - Table for selection of wood type and detail type per grouping (fx. "G_String1" and "G_Island"):
        - Pull list of wood types and detail types from shared settings, as well as "overwrite indicators" (first number in name).
    - Should contain "excel file" file select input (this should probably be custom).
        - Save path on project basis, replacing part of path with per-user "Dropbox path".
        - On opening of the command, validate if the previous path is still valid, and notify user the file no longer exists.
    - For advanced users, add a "Overwrite" checkmark that is defaulted to true.
        - If the checkmark is false, display an output excel file file select input.

- Add assertions and nicer output to excel_lib.
    - Missing sheet/tables.
    - Wrong column count in tables.

- Add check for newer BodyCount release, and automatic update.

# WIP

- Investigate bug where new instances of Fusion opens instead of installing openpyxl on first run of Add-In.
    - It seems like that way I'm trying to install openpyxl doesn't play nice outside of the debug environment.
        - **So it turns out the `sys.executable` points to Fusion.exe instead of the python instance... So I've had to extract the path relative to Fusion.exe. This works on my machine, but might not work on others - have to test.**
        - **My fix probably also doesn't work at all on Mac's.**

# Done

- Ignore bodies with names matching "Body\d+" and "delete.*" with any case.
- Change "Define Strings" to "Display Ungrouped".
    - "Strings" will be defined by top-level component, with some prefix (fx. "G_String1" and "G_Island").
    - "Display Ungrouped" should mark all ungrouped bodies in transparent red, showing through other components.
