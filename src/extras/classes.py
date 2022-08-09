import adsk.core, adsk.fusion, adsk.cam, traceback
import re

from typing import List, Dict
from pathlib import Path

import importlib
import os, sys
packagepath = os.path.join(os.path.dirname(sys.argv[0]), 'Lib/site-packages/')
if packagepath not in sys.path:
    sys.path.append(packagepath)

# Install and import xlsxwriter
try:
    importlib.import_module("xlsxwriter")
except ImportError:
    import pip
    pip.main(["install", "xlsxwriter"])
finally:
    globals()["xw"] = importlib.import_module("xlsxwriter")

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

    def add_to_ws(self, ws: xw.workbook.Worksheet):
        """Saves count to worksheet"""

        # Populate headers
        ws.write(1,  1, "Body")
        ws.write(1,  2, "Count")
        ws.write(1,  3, "Material")
        ws.write(1,  4, "Mass [kg]")

        # Add data per BRepBody
        keys = list(self._counts.keys())
        keys.sort()
        for idx, k in enumerate(keys):
            v = self._counts[k]
            ws.write(2+idx, 1, k)
            ws.write(2+idx, 2, v._count)
            ws.write(2+idx, 3, v._material)
            ws.write(2+idx, 4, v._mass)

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

    def add_to_db(self, wb: xw.Workbook):
        """Saves prices to worksheet"""
        count_ws: xw.workbook.Worksheet = wb.get_worksheet_by_name("Counts")
        calc_ws : xw.workbook.Worksheet = wb.get_worksheet_by_name("Calculations")

        # Populate headers
        count_ws.write(1,  7, "Modules")
        count_ws.write(1,  8, "Category")
        count_ws.write(1,  9, "Price per")
        count_ws.write(1, 10, "Count")
        count_ws.write(1, 11, "Cost")
        count_ws.write(1, 12, "Price ex. VAT")
        count_ws.write(1, 13, "Price incl. VAT")
        count_ws.write(1, 15, "Installation type")
        count_ws.write(1, 16, "Price per")
        count_ws.write(1, 18, "Category")
        count_ws.write(1, 19, "Abbreviation")
        calc_ws .write(1,  1, "Module")
        calc_ws .write(1,  2, "Count")
        calc_ws .write(1,  3, "Installation type")
        calc_ws .write(1,  4, "Price")

        # Set installation types and headers
        count_ws.write(2, 15, "CA")
        count_ws.write(2, 16, "2200")
        count_ws.write(3, 15, "Li")
        count_ws.write(3, 16, "1500")
        count_ws.write(4, 15, "Is")
        count_ws.write(4, 16, "1800")
        count_ws.write(5, 15, "default")
        count_ws.write(5, 16, "0")

        # Set Category abbreviations
        abbreviations = {
            "String 1": "S1",
            "String 2": "S2",
            "String 3": "S3",
            "Island": "I",
            "Custom": "C",
        }
        for i, (cat, abbr) in enumerate(abbreviations.items()):
            count_ws.write(2+i, 18, cat)
            count_ws.write(2+i, 19, abbr)

        # Add data per Price
        keys = list(self._prices.keys())
        keys.sort()
        for idx, k in enumerate(keys):
            v = self._prices[k]
            count_ws.write(2+idx,  7, k)
            count_ws.write(2+idx,  9, int(v._price_per))
            count_ws.write(2+idx, 10, int(v._count))
            count_ws.write(2+idx, 11, f"=J{3+idx}*K{3+idx}")
            count_ws.write(2+idx, 12, f"=ROUND(L{3+idx}*1.4)")
            count_ws.write(2+idx, 13, f"=ROUND(M{3+idx}*1.25)")

            row = 2+idx
            calc_ws.write(row, 1, f"=Counts!H{row+1}")
            calc_ws.write(row, 2, f"=Counts!K{row+1}")
            calc_ws.write(row, 3, f"=MID(B{row+1}, FIND(\"-\", B{row+1})+1, 2)")
            calc_ws.write(row, 4, f"=LOOKUP(D{row+1}, Counts!P3:P100, Counts!Q3:Q100)*C{row+1}")

        n_modules = len(keys)
        categories = ["String 1", "String 2", "String 3", "Island", "Custom"]
        for idx, c in enumerate(categories):
            de_row = 2+n_modules # Data End row
            abbr_row = 3+idx
            row = 2+idx+n_modules # Current row

            cat = f"K{row+1}"
            abbr = f"T{abbr_row}"
            cats = f"I3:I{de_row}"
            costs = f"L3:L{de_row}"
            costs_excl_vat = f"M3:M{de_row}"
            costs_incl_vat = f"N3:N{de_row}"

            count_ws.write(row, 10, c)
            count_ws.write_array_formula(row, 11, row, 11,
                    fr"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs}, 0))}}")
            count_ws.write_array_formula(row, 12, row, 12,
                    fr"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs_excl_vat}, 0))}}")
            count_ws.write_array_formula(row, 13, row, 13,
                    fr"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs_incl_vat}, 0))}}")

        bold_format = wb.add_format()
        bold_format.set_bold()
        row = 2+n_modules+len(categories)
        count_ws.write(row, 10, "Total")
        count_ws.write(row, 11, f"=SUM(L{row-len(categories)}:L{row-1})")
        count_ws.write(row, 12, f"=SUM(M{row-len(categories)}:M{row-1})")
        count_ws.write(row, 13, f"=SUM(N{row-len(categories)}:N{row-1})")
        row += 1
        count_ws.write(row, 10, "Total Installation")
        count_ws.write(row, 11, f"=SUM(Calculations!E3:E{3+n_modules})")

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
