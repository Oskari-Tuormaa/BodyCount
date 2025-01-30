import adsk.fusion
import adsk.core

from ... import config

from serde import serde
from serde.json import to_json, from_json
from dataclasses import field
from pathlib import Path

app = adsk.core.Application.get()
ui = app.userInterface

ATTR_GRP = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}'
ATTR_NAME = "file_data"

@serde
class ModuleSettings:
    category_name: str
    detail_material: str | None
    wood_material: str | None

@serde
class FileData:
    excel_path: Path | None = None
    modules: dict[str, ModuleSettings] = field(default_factory=lambda: dict())

def load_file_data() -> FileData:
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)

    if (attr := design.attributes.itemByName(ATTR_GRP, ATTR_NAME)):
        return from_json(FileData, attr.value)
    return FileData()

def save_file_data(file_data: FileData):
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    design.attributes.add(ATTR_GRP, ATTR_NAME, to_json(file_data))