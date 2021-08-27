import adsk.core, adsk.fusion, adsk.cam, traceback
import re
from .packages import pylightxl

from typing import List, Dict


class CountObject(object):

    """Class that holds count of single body"""

    _count: int = 1
    _material: adsk.core.Material
    _mass: float
    
    def __init__(self, body: adsk.fusion.BRepBody):
        self._material = body.material.name
        self._mass = body.physicalProperties.mass
    
    def __add__(self, other):
        # Add two CountObject together
        if type(other) == type(self):
            self._count += other._count
            self._mass += other._mass
        
        # Add BRepBody to CountObject
        elif type(other) == adsk.fusion.BRepBody:
            self._count += 1
            self._mass += other.physicalProperties.mass

        else:
            raise ValueError(f"Invalid type {type(other)} of {other}")
        return self
    
    def __radd__(self, other):
        return self.__add__(other)


class Count(object):

    """Holds count of several bodies"""

    _counts: Dict[str, CountObject]
    _re_extract_name: re.Pattern
    
    def __init__(self):
        self._counts = dict()
        # Regex for removing duplicate count from body name
        self._re_extract_name = re.compile("^.*?(?=$|\(.*?\))")

    def save_xlsx(self, path: str):
        """Saves count to xlsx file specified by path"""
        db = pylightxl.Database()
        db.add_ws("Sheet1")
        ws: pylightxl.pylightxl.Worksheet = db.ws("Sheet1")

        # Populate headers
        ws.update_index(2, 2, "Bodies")
        ws.update_index(2, 3, "Counts")
        ws.update_index(2, 4, "Material")
        ws.update_index(2, 5, "Mass [kg]")

        # Add data per BRepBody
        keys = list(self._counts.keys())
        keys.sort()
        for idx, k in enumerate(keys):
            v = self._counts[k]
            ws.update_index(3+idx, 2, k)
            ws.update_index(3+idx, 3, v._count)
            ws.update_index(3+idx, 4, v._material)
            ws.update_index(3+idx, 5, v._mass)

        # Write file
        pylightxl.writexl(db, path)
    
    def __add__(self, other):
        # Add two Count together
        if type(other) == type(self):
            for k in other._counts.keys():
                if k in self._counts:
                    self._counts[k] += other._counts[k]
                else:
                    self._counts[k] = other._counts[k]

        # Add BRepBody to Count
        elif type(other) == adsk.fusion.BRepBody:
            if other.isVisible:
                name = self._re_extract_name.findall(other.name)[0].strip()
                if name in self._counts:
                    self._counts[name] += other
                else:
                    self._counts[name] = CountObject(other)

        # Add BRepBodies to Count
        elif type(other) == adsk.fusion.BRepBodies:
            for body in other:
                self += body

        # Add Component to Count
        elif type(other) == adsk.fusion.Component:
            if other.isBodiesFolderLightBulbOn:
                self += other.bRepBodies
        
        # Add Components to Count
        elif type(other) == adsk.fusion.Components:
            for comp in other:
                self += comp

        else:
            raise ValueError(f"Invalid type {type(other)} of {other}")
        return self
    
    def __radd__(self, other):
        return self.__add__(other)