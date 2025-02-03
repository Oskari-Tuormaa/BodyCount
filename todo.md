
# TODO

- [ ] Updates to command input:
    - Pull list of "overwrite indicators" (first number in name).
    - Save "excel file" path on project basis, replacing part of path with per-user "Dropbox path".
- [ ] Fix textboxes in command menus being too small, and being hidden by scrollbars.

- [ ] Add assertions and nicer output to excel_lib.
    - Missing sheet/tables.
    - Wrong column count in tables.

- [ ] Add check for newer BodyCount release, and automatic update.

# WIP

# Done

- [x] Ignore bodies with names matching "Body\d+" and "delete.*" with any case.
- [x] Change "Define Strings" to "Display Ungrouped".
    - "Strings" will be defined by top-level component, with some prefix (fx. "G_String1" and "G_Island").
    - "Display Ungrouped" should mark all ungrouped bodies in transparent red, showing through other components.
- [x] Add per-user settings.
    - Add "Dropbox path" setting variable, that should point to Vermland folder in Dropbox.
- [x] Add Vermland shared settings.
    - Can hold valid detail and wood materials, as well as defaults.
    - Path to shared settings is based on "Dropbox path": <dropbox_path>/Vermland/The Collection/BodyCount.
- [x] Investigate bug where new instances of Fusion opens instead of installing openpyxl on first run of Add-In.
    - It seems like that way I'm trying to install openpyxl doesn't play nice outside of the debug environment.
        - **So it turns out the `sys.executable` points to Fusion.exe instead of the python instance... So I've had to extract the path relative to Fusion.exe. This works on my machine, but might not work on others - have to test.**
        - **My fix probably also doesn't work at all on Mac's.**
        - **Seems to work on Simons machine!**
- [X] Add command input box to "Count bodies" command.
    - Table for selection of wood type and detail type per grouping (fx. "G_String1" and "G_Island"):
    - Should contain "excel file" file select input (this should probably be custom).
    - For advanced users, add a "Overwrite" checkmark that is defaulted to true.
        - If the checkmark is false, display an output excel file file select input.
    - Body name to material conventions:
        - 1-9 + 12 + 14-16: Overwrite with chosen wood type
        - 10: Overwrite with chosen detail type
        - All others: Use internal Fusion material
        - OBS: Names starting with 97-98 should use number following the first space to determine material.
            - Ex: "98.1_Modified 5.2_Front_W592_H389" should use 5 and "98.1_Modified 10.2_DrawerBolt" should use 10.