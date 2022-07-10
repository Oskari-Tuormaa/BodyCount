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

    def __init__(self, other: adsk.fusion.Occurrence, count=1):
        self._count: int = count
        self._price_per = 0

        try:  # Try to use description of component as price
            self._price_per = float(other.component.description)
        except (ValueError, AttributeError):  # Description isn't valid price
            prices = PriceCount()
            prices += other.childOccurrences
            self._price_per = prices.total()

    def __add__(self, other):
        # Add two PriceObject together
        if type(other) == type(self):
            self._count += other._count
        return self

    def __radd__(self, other):
        return self.__add__(other)


class PriceCount(object):

    def __init__(self):
        self._prices: Dict[str, PriceObject] = dict()
        # Regex for removing version number from name
        self._re_remove_version: re.Pattern = re.compile(r"^.*?(?=$| v\d+)")

    def add_to_db(self, ws: pylightxl.Worksheet):
        """Saves prices to worksheet"""

        # Populate headers
        ws.update_index(2,  8, "Modules")
        ws.update_index(2,  9, "Category")
        ws.update_index(2, 10, "Price per")
        ws.update_index(2, 11, "Count")
        ws.update_index(2, 12, "Cost")
        ws.update_index(2, 13, "Price ex. VAT")
        ws.update_index(2, 14, "Price incl. VAT")

        # Add data per Price
        keys = list(self._prices.keys())
        keys.sort()
        for idx, k in enumerate(keys):
            v = self._prices[k]
            ws.update_index(3+idx,  8, k)
            ws.update_index(3+idx, 10, v._price_per)
            ws.update_index(3+idx, 11, v._count)
            ws.update_index(3+idx, 12, f"=J{3+idx}*K{3+idx}")
            ws.update_index(3+idx, 13, f"=L{3+idx}*1.4")
            ws.update_index(3+idx, 14, f"=M{3+idx}*1.25")

        n_modules = len(keys)
        categories = ["String 1", "String 2", "String 3", "Island", "Custom"]
        for idx, c in enumerate(categories):
            de_row = 2+n_modules # Data End row
            row = 3+idx+n_modules # Current row
            ws.update_index(row, 11, c)
            ws.update_index(row, 12, fr"=SUMIF(I3:I{de_row},K{row},L3:L{de_row})")
            ws.update_index(row, 13, fr"=SUMIF(I3:I{de_row},K{row},M3:M{de_row})")
            ws.update_index(row, 14, fr"=SUMIF(I3:I{de_row},K{row},N3:N{de_row})")

        row = 3+n_modules+len(categories)
        ws.update_index(row, 11, "Total")
        ws.update_index(row, 12, f"=SUM(L{row-len(categories)}:L{row-1})")
        ws.update_index(row, 13, f"=SUM(M{row-len(categories)}:M{row-1})")
        ws.update_index(row, 14, f"=SUM(N{row-len(categories)}:N{row-1})")

    def __add__(self, other):
        if type(other) == adsk.fusion.OccurrenceList:
            for occ in other:
                if not occ.isVisible:
                    continue

                name = self._re_remove_version.findall(occ.component.name)[0].strip()
                if name in self._prices:  # Skip if component has already been indexed
                    continue

                n_occ = len([x for x in other if (self._re_remove_version.findall(x.component.name)[0].strip() == name and x.isVisible)])
                self._prices[name] = PriceObject(occ, n_occ)
        return self

    def __radd__(self, other):
        return self.__add__(other)

    def total(self):
        res = 0
        for k, v in self._prices.items():
            res += v._count * v._price_per
        return res
