import adsk.core, adsk.fusion, adsk.cam, traceback
import re
from .packages import pylightxl

from typing import List, Dict
from pathlib import Path

log_dir = Path(Path(__file__).parent.parent, "logs")


class CountObject(object):

    """Class that holds count of single body"""
    
    def __init__(self, body: adsk.fusion.BRepBody):
        self._count = 1
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


class BodyCount(object):

    """Holds count of several bodies"""
    
    def __init__(self):
        self._counts: Dict[str, CountObject] = dict()
        # Regex for removing duplicate count from body name
        self._re_extract_name: re.Pattern = re.compile("^.*?(?=$|\(.*?\))")

    def add_to_ws(self, ws: pylightxl.Worksheet):
        """Saves count to worksheet"""

        # Populate headers
        ws.update_index(2,  2, "Body")
        ws.update_index(2,  3, "Count")
        ws.update_index(2,  4, "Material")
        ws.update_index(2,  5, "Mass [kg]")

        # Add data per BRepBody
        keys = list(self._counts.keys())
        keys.sort()
        for idx, k in enumerate(keys):
            v = self._counts[k]
            ws.update_index(3+idx, 2, k)
            ws.update_index(3+idx, 3, v._count)
            ws.update_index(3+idx, 4, v._material)
            ws.update_index(3+idx, 5, v._mass)
    
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
                
        # Add OccurrenceList to Count
        elif type(other) == adsk.fusion.OccurrenceList:
            for occ in other:
                self += occ
        
        # Add Occurrence to Count
        elif type(other) == adsk.fusion.Occurrence:
            self += other.childOccurrences
            if other.isVisible:
                self += other.component

        # Add Component to Count
        elif type(other) == adsk.fusion.Component:
            if other.isBodiesFolderLightBulbOn:
                self += other.bRepBodies

        else:
            raise ValueError(f"Invalid type {type(other)} of {other}")
        return self
    
    def __radd__(self, other):
        return self.__add__(other)


class PriceObject(object):

    """Holds price info for single component"""

    def __init__(self, other, depth = [0]):
        with open(Path(log_dir, "output.log"), "a+") as fd:
            fd.write("\t"*depth[0] + str(type(other)) + " " + str(other.name) + "\n")
            if type(other) == adsk.fusion.Component:
                occurrences = other.occurrences.asList
                if len(occurrences) != 0:
                    co = occurrences
                    fd.write("\t"*depth[0] + str(co) + str(len(co)) + "\n")
            fd.write("\n")

        self._count: int = 1
        if type(other) == adsk.fusion.Component:
            try:
                self._price_per = float(other.description)
            except (ValueError, AttributeError):
                other: adsk.fusion.Component = other
                prices = PriceCount(include_bodies=True)
                depth[0] += 1
                prices += other.bRepBodies
                prices += other.occurrences.asList
                depth[0] -= 1
                
                with open(Path(log_dir, "output.log"), "a+") as fd:
                    for k, v in prices._prices.items():
                        fd.write("\t"*depth[0] + k + " " + str(v._count) + " " + str(v._price_per) + "\n")
                    fd.write("\n")

                self._price_per = sum(x._count * x._price_per for x in prices._prices.values())
        else:
            try:
                self._price_per = float(other.description)
            except (ValueError, AttributeError):
                # Description isn't valid price
                self._price_per = 0
    
    def __add__(self, other):
        # Add two PriceObject together
        if type(other) == type(self):
            self._count += other._count
        else:
            self._count += 1
        return self
    
    def __radd__(self, other):
        return self.__add__(other)


class PriceCount(object):

    def __init__(self, include_bodies: bool = False):
        self.include_bodies = include_bodies
        self._prices: Dict[str, PriceObject] = dict()
        # Regex for removing version number from name
        self._re_remove_version: re.Pattern = re.compile("^.*?(?=$| v\d+)")
    
    def add_to_ws(self, ws: pylightxl.Worksheet):
        """Saves prices to worksheet"""

        # Populate headers
        ws.update_index(2,  8, "Components")
        ws.update_index(2,  9, "Price per")
        ws.update_index(2, 10, "Count")
        ws.update_index(2, 11, "Total price")
        
        # Add data per Price
        keys = list(self._prices.keys())
        keys.sort()
        for idx, k in enumerate(keys):
            v = self._prices[k]
            ws.update_index(3+idx,  8, k)
            ws.update_index(3+idx,  9, v._price_per)
            ws.update_index(3+idx, 10, v._count)
            ws.update_index(3+idx, 11, v._price_per*v._count)

    def __add__(self, other):
        if self.include_bodies and type(other) == adsk.fusion.BRepBody:
            if other.isVisible:
                # Add to PriceObject
                name = self._re_remove_version.findall(other.name)[0].strip()
                if name in self._prices:
                    self._prices[name] += other
                else:
                    self._prices[name] = PriceObject(other)

        elif self.include_bodies and type(other) == adsk.fusion.BRepBodies:
            for body in other:
                self += body

        elif type(other) == adsk.fusion.OccurrenceList:
            for occ in other:
                self += occ
        
        elif type(other) == adsk.fusion.Occurrence:
            if other.isVisible:
                self += other.component
                if self.include_bodies:
                    self += other.childOccurrences
        
        elif type(other) == adsk.fusion.Component:
            if other.isBodiesFolderLightBulbOn:
                # Add to PriceObject
                name = self._re_remove_version.findall(other.name)[0].strip()
                if name in self._prices:
                    self._prices[name] += other
                else:
                    self._prices[name] = PriceObject(other)

        else:
            raise ValueError(f"Invalid type {type(other)} of {other}")
        return self
    
    def __radd__(self, other):
        return self.__add__(other)