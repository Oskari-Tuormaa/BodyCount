import adsk.core, adsk.fusion, adsk.cam, traceback
import re
from .packages import pylightxl

from typing import List, Dict


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

    def __init__(self, comp: adsk.fusion.Component):
        self._count: int = 1
        try:
            self._price_per = float(comp.description)
        except (ValueError, AttributeError):
            # Description isn't valid price
            self._price_per = 0
    
    def __add__(self, other):
        # Add two PriceObject together
        if type(other) == type(self):
            self._count += other._count
        
        # Add Component to PriceObject
        elif type(other) == adsk.fusion.Component:
            self._count += 1

        else:
            raise ValueError(f"Invalid type {type(other)} of {other}")
        return self
    
    def __radd__(self, other):
        return self.__add__(other)


class PriceCount(object):

    def __init__(self):
        self._prices: Dict[str, PriceObject] = dict()
    
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
        if type(other) == adsk.fusion.OccurrenceList:
            for occ in other:
                self += occ
        
        elif type(other) == adsk.fusion.Occurrence:
            self += other.component
        
        elif type(other) == adsk.fusion.Component:
            # Add to PriceObject
            if other.name in self._prices:
                self._prices[other.name] += other
            else:
                self._prices[other.name] = PriceObject(other)

        else:
            raise ValueError(f"Invalid type {type(other)} of {other}")
        return self
    
    def __radd__(self, other):
        return self.__add__(other)