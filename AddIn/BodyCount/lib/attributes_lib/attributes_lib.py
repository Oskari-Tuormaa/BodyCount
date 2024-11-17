import adsk.fusion
import adsk.core
import json

from dataclasses import dataclass, asdict

from .. import fusionAddInUtils as futil
from ... import config

app = adsk.core.Application.get()
ui = app.userInterface

class Serializable:
    @classmethod
    def deserialize(cls, data: str):
        return cls(**json.loads(data))

    def serialize(self) -> str:
        return json.dumps(asdict(self))

@dataclass
class PerFileData(Serializable):
    strings: list[str]

@dataclass
class ModuleData(Serializable):
    string: str

ATTR_GRP = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}'
PER_FILE_ATTR_NAME = "per_file"

def get_file_data() -> PerFileData:
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)

    if (attr := design.attributes.itemByName(ATTR_GRP, PER_FILE_ATTR_NAME)):
        return PerFileData.deserialize(attr.value)
    return PerFileData([])

def set_file_data(data: PerFileData):
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    design.attributes.add(ATTR_GRP, PER_FILE_ATTR_NAME, data.serialize())