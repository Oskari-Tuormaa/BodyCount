import adsk.fusion
import adsk.core

from ... import config
from .user_settings import load_user_data

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
    is_excel_path_in_dropbox: bool = False
    modules: dict[str, ModuleSettings] = field(default_factory=lambda: dict())

def load_file_data() -> FileData:
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)

    if (attr := design.attributes.itemByName(ATTR_GRP, ATTR_NAME)):
        file_data = from_json(FileData, attr.value)

        # Restore Dropbox path
        if file_data.excel_path is not None and file_data.is_excel_path_in_dropbox:
            user_data = load_user_data()

            if user_data.shared_data_path is None:
                raise RuntimeError() # TODO: Warn user that shared_data_path is not set

            file_data.excel_path = user_data.shared_data_path/file_data.excel_path

    else:
        file_data = FileData()

    return file_data

def save_file_data(file_data: FileData):
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)

    user_data = load_user_data()

    # Replace Dropbox path
    if (
        user_data.shared_data_path is not None and
        file_data.excel_path is not None and
        file_data.excel_path.is_relative_to(user_data.shared_data_path)
    ):
        file_data.excel_path = file_data.excel_path.relative_to(user_data.shared_data_path)
        file_data.is_excel_path_in_dropbox = True
    else:
        file_data.is_excel_path_in_dropbox = False

    saved = design.attributes.add(ATTR_GRP, ATTR_NAME, to_json(file_data))