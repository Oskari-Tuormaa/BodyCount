import adsk.fusion
import adsk.core
import json

from dataclasses import dataclass, asdict
from typing import Protocol, Generator

from .. import counting_lib
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

class SupportsAttributes(Protocol):
    @property
    def attributes(self) -> adsk.core.Attributes:
        ...

@dataclass
class PerFileData(Serializable):
    strings: list[str]

@dataclass
class ModuleData(Serializable):
    string: str

ATTR_GRP = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}'
PER_FILE_ATTR_NAME = "per_file"
MODULE_ATTR_NAME = "module"

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

def get_module_data(obj: SupportsAttributes) -> ModuleData | None:
    if (attr := obj.attributes.itemByName(ATTR_GRP, MODULE_ATTR_NAME)):
        return ModuleData.deserialize(attr.value)

def set_module_data(obj: SupportsAttributes, data: ModuleData):
    obj.attributes.add(ATTR_GRP, MODULE_ATTR_NAME, data.serialize())

def traverse_module_data() -> Generator[tuple[adsk.fusion.Occurrence, ModuleData], None, None]:
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    rootComp = design.rootComponent

    for occ in counting_lib.traverse_occurrences(rootComp):
        if (data := get_module_data(occ)):
            yield (occ, data)