# BodyCount Add-In Architecture

## System Overview

BodyCount is a Fusion 360 add-in with three main commands that help users count and categorize CAD bodies, export data to Excel, and visualize ungrouped components.

## Architecture Diagram - Overall Structure

```plantuml
@startuml BodyCount_Overview
skinparam componentStyle rectangle
skinparam backgroundColor #FEFEFE

rectangle "Fusion 360 Core API" #E8F4F8 {
  component "Design Model\n(Components, Bodies)" as fusion_model
  component "UI & Commands" as fusion_ui
  component "Custom Graphics" as fusion_graphics
}

rectangle "BodyCount Add-In Layer" {
  component "Count Bodies\nCommand" as cmd_count #FFE6CC
  component "Show Ungrouped\nCommand" as cmd_show #FFE6CC
  component "Settings\nCommand" as cmd_settings #FFE6CC
}

rectangle "Library Layer" {
  component "counting_lib\n(traverse, sort)" as lib_counting #E6F2FF
  component "excel_lib\n(workbook I/O)" as lib_excel #E6F2FF
  component "settings_lib\n(config mgmt)" as lib_settings #E6F2FF
  component "custom_graphics_lib\n(visualization)" as lib_gfx #E6F2FF
  component "fusionAddInUtils\n(logging, errors)" as lib_utils #F0F0F0
}

rectangle "External Systems" #F5F5F5 {
  database "Excel Files" as excel_files
  database "JSON Settings\n(Local & Dropbox)" as config_files
}

cmd_count --> lib_counting
cmd_count --> lib_excel
cmd_count --> lib_settings
cmd_count --> lib_utils

cmd_show --> lib_counting
cmd_show --> lib_gfx
cmd_show --> lib_utils

cmd_settings --> lib_settings
cmd_settings --> lib_utils

lib_counting --> fusion_model
lib_excel --> excel_files
lib_gfx --> fusion_graphics
lib_settings --> config_files
lib_utils --> fusion_ui
lib_utils --> fusion_graphics

@enduml
```

## Command-Specific Data Flows

### Count Bodies Command
```plantuml
@startuml CountBodies_Flow
skinparam componentStyle rectangle
skinparam backgroundColor #FEFEFE

actor "User" as user
rectangle "Count Bodies Command" #FFE6CC {
  component "UI Dialog" as ui_dialog
  component "Material Mapper" as material_mapper
}

rectangle "Libraries" {
  component "counting_lib" as lib_counting #E6F2FF
  component "excel_lib" as lib_excel #E6F2FF
  component "settings_lib" as lib_settings #E6F2FF
}

rectangle "Fusion 360" {
  component "Design Model" as fusion_model #E8F4F8
}

rectangle "Storage" {
  database "Excel File" as excel_db
  database "Settings JSON" as settings_db
}

user --> ui_dialog : selects file &\nmaterial types
ui_dialog --> lib_settings : load previous\nsettings
settings_db --> lib_settings
lib_settings --> ui_dialog : populate\ndropdowns

ui_dialog --> material_mapper : map selections
material_mapper --> lib_counting : traverse &\ncollect bodies
fusion_model --> lib_counting
lib_counting --> material_mapper : list of\nbodies

material_mapper --> lib_excel : write data
lib_excel --> excel_db : update tables

material_mapper --> lib_settings : save\nsettings
lib_settings --> settings_db

@enduml
```

### Show Ungrouped Command
```plantuml
@startuml ShowUngrouped_Flow
skinparam componentStyle rectangle
skinparam backgroundColor #FEFEFE

actor "User" as user
rectangle "Show Ungrouped Command" #FFE6CC {
  component "Handler" as handler
}

rectangle "Libraries" {
  component "counting_lib" as lib_counting #E6F2FF
  component "custom_graphics_lib" as lib_gfx #E6F2FF
}

rectangle "Fusion 360" {
  component "Design Model" as fusion_model #E8F4F8
  component "Custom Graphics" as fusion_graphics #E8F4F8
}

user --> handler : click button
handler --> lib_counting : find ungrouped\nbodies
fusion_model --> lib_counting
lib_counting --> handler : list of\nungrouped bodies

handler --> lib_gfx : create\nvisualization
lib_gfx --> fusion_graphics

@enduml
```

### Settings Command
```plantuml
@startuml Settings_Flow
skinparam componentStyle rectangle
skinparam backgroundColor #FEFEFE

actor "User" as user
rectangle "Settings Command" #FFE6CC {
  component "UI Dialog" as ui_dialog
}

rectangle "Libraries" {
  component "settings_lib" as lib_settings #E6F2FF
}

rectangle "Storage" {
  database "User Settings\n(.user-settings.json)" as user_settings_db
  database "Shared Settings" as shared_settings_db
}

user --> ui_dialog : enter Dropbox path
ui_dialog --> lib_settings : validate &\nsave path

lib_settings --> user_settings_db : save user\npreferences
lib_settings --> shared_settings_db : load available\nmaterials

shared_settings_db --> lib_settings
lib_settings --> ui_dialog : populate\nmaterial lists

@enduml
```

## Data Flow

### Count Bodies Command
```
User Input (Excel file, Material selection)
    ↓
→ Load File Settings (previous selections)
→ Load Shared Settings (available materials)
→ Traverse Design (counting_lib)
    └─ Iterate components recursively
    └─ Filter by visibility & naming rules
    └─ Collect bodies per module/category
→ Map Bodies to Materials
    └─ Apply naming conventions (1-9, 10, 12, etc.)
→ Write to Excel (excel_lib)
    └─ Open Excel workbook
    └─ Populate IndividualParts table
    └─ Populate ModulesParts table
    └─ Save file
→ Save File Settings (remember selections)
```

### Show Ungrouped Command
```
User clicks "Show Ungrouped"
    ↓
→ Traverse Design (counting_lib)
→ Find ungrouped bodies
→ Create Custom Graphics (custom_graphics_lib)
    └─ Render transparent red highlighting
```

### Settings Command
```
User enters Dropbox path
    ↓
→ Validate path (check writeability)
→ Save to User Settings (.user-settings.json)
→ Load/refresh Shared Settings from Dropbox
    └─ Available materials (wood, steel/brass numbers)
    └─ Update dropdown lists
```

## Key Design Patterns

1. **Modular Commands**: Each command is independent but shares common utilities
2. **Settings Hierarchy**: User settings → File settings → Shared settings (Dropbox)
3. **Generator-based Traversal**: Recursive generators for memory-efficient tree navigation
4. **Type-hinted Dataclasses**: Strong typing for Excel structures (Body, Module)
5. **Centralized Error Handling**: All errors routed through `futil.handle_error()`
6. **Event-driven UI**: Fusion 360 command handlers for user interaction
7. **Directory Writeability Validation**: Real file creation test to ensure write access

## File Organization

```
BodyCount/
├── BodyCount.py              # Add-in entry point (startup/shutdown)
├── config.py                 # Global configuration
├── commands/
│   ├── count_bodies/         # Main counting & export command
│   ├── show_ungrouped/       # Visual highlighting command
│   └── settings/             # User settings command
└── lib/
    ├── counting_lib/         # CAD model traversal
    ├── excel_lib/            # Excel I/O operations
    ├── settings_lib/         # Settings management (JSON serialization)
    ├── custom_graphics_lib/  # Visual rendering
    └── fusionAddInUtils/     # Logging, error handling, event management
```

## Material Support

### Wood Materials
- **Storage**: Shared settings (Dropbox)
- **Management**: Dynamic list editable in Settings dialog
- **Scope**: All modules can select wood material

### Steel/Brass Materials
- **Storage**: Shared settings (steel/brass part number mapping)
- **Management**: Hardcoded options [Steel, Brass]
- **Scope**: All modules can select Steel or Brass
- **Mapping**: Steel/Brass numbers in shared settings determine material selection

## Settings File Structure

### User Settings (`.user-settings.json`)
```json
{
    "shared_data_path": "/path/to/dropbox",
    "overwrite": true
}
```

### Shared Settings (`<Dropbox Path>/Vermland/The Collection/BodyCount/shared_settings.json`)
```json
{
    "wood_materials": ["Material1", "Material2", ...],
    "steel_brass_numbers": [[10, 20], [11, 21], ...]
}
```

### File Settings (Project-specific, stored in project)
Stores per-module material selections and Excel file path references.
