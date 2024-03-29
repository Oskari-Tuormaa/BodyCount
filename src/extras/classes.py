import adsk.core, adsk.fusion, adsk.cam, traceback
import re

from typing import List, Dict
from pathlib import Path
from .human_sort import human_sort

import importlib
import os, sys
packagepath = os.path.join(os.path.dirname(sys.argv[0]), 'Lib', 'site-packages')
if packagepath not in sys.path:
    sys.path.append(packagepath)

# Install and import xlsxwriter
try:
    import xlsxwriter as xw
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
        human_sort(keys)
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
        indesign_ws : xw.workbook.Worksheet = wb.get_worksheet_by_name("InDesign")

        # Populate headers
        count_ws.write(1,  7, "Modules")
        count_ws.write(1,  8, "Category")
        count_ws.write(1,  9, "Price per")
        count_ws.write(1, 10, "Price ex. VAT")
        count_ws.write(1, 11, "Price incl. VAT")
        count_ws.write(1, 13, "Installation type")
        count_ws.write(1, 14, "Price per")
        count_ws.write(1, 16, "Category")
        count_ws.write(1, 17, "Abbreviation")
        calc_ws .write(1,  1, "Module")
        calc_ws .write(1,  2, "Installation type")
        calc_ws .write(1,  3, "Price")

        # Set installation types and headers
        count_ws.write(2, 13, "CA")
        count_ws.write(2, 14, 2200)
        count_ws.write(3, 13, "Li")
        count_ws.write(3, 14, 1500)
        count_ws.write(4, 13, "Is")
        count_ws.write(4, 14, 1800)

        # Set Category abbreviations
        categories = {
            "String 1": "S1",
            "String 2": "S2",
            "String 3": "S3",
            "Island": "I",
            "Custom": "C",
        }
        for i, (cat, abbr) in enumerate(categories.items()):
            count_ws.write(2+i, 16, cat)
            count_ws.write(2+i, 17, abbr)

        n_extra = 5

        # Add data per Price
        modules = list(self._prices.keys())
        human_sort(modules)
        row_modules_end = 2
        for mod in modules + [""]*n_extra:
            v = self._prices.get(mod)

            n = 1 if v is None else v._count
            per = 0 if v is None else v._price_per

            for i in range(n):
                count_ws.write(row_modules_end,  7, mod)
                count_ws.data_validation(row_modules_end, 8, row_modules_end, 8, {'validate': 'list', 'source': '=$R$3:$R$7'})
                count_ws.write(row_modules_end,  9, int(per))
                count_ws.write(row_modules_end, 10, f"=ROUND(J{row_modules_end+1}*1.4 + IF(ISBLANK(H{row_modules_end+1}), 0, 2206), 0)")
                count_ws.write(row_modules_end, 11, f"=ROUND(K{row_modules_end+1}*1.25, 0)")

                calc_ws.write(row_modules_end, 1, f"=Counts!H{row_modules_end+1}")
                calc_ws.write(row_modules_end, 2, f"=IFERROR(MID(B{row_modules_end+1}, FIND(\"-\", B{row_modules_end+1})+1, 2),\"\")")
                calc_ws.write(row_modules_end, 3, f"=IFERROR(INDEX(Counts!O3:O100, MATCH(C{row_modules_end+1}, Counts!N3:N100)), 0)")
                row_modules_end += 1

        n_modules = len(modules)+n_extra
        for idx, c in enumerate(categories.keys()):
            de_row = row_modules_end
            abbr_row = 3+idx
            row = idx+de_row

            cat = f"I{row+1}"
            abbr = f"R{abbr_row}"
            cats = f"I3:I{de_row}"
            costs = f"J3:J{de_row}"
            costs_excl_vat = f"K3:K{de_row}"
            costs_incl_vat = f"L3:L{de_row}"

            count_ws.write(row, 8, c)
            count_ws.write_array_formula(row, 9, row, 9,
                    fr"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs}, 0))}}")
            count_ws.write_array_formula(row, 10, row, 10,
                    fr"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs_excl_vat}, 0))}}")
            count_ws.write_array_formula(row, 11, row, 11,
                    fr"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs_incl_vat}, 0))}}")

        row = row_modules_end+len(categories)
        count_ws.write(row, 8, "Total")
        count_ws.write(row, 9, f"=SUM(J{row-len(categories)+1}:J{row})")
        count_ws.write(row, 10, f"=SUM(K{row-len(categories)+1}:K{row})")
        count_ws.write(row, 11, f"=SUM(L{row-len(categories)+1}:L{row})")

        # Write category counts
        ix = 6
        iy = 1
        n_raw_modules = row_modules_end - 3
        for i, cat in enumerate(categories.keys()):
            calc_ws.write(iy, ix+i, cat)

            abbr = f"Counts!R{3+i}"
            cats = f"Counts!I3:I{row_modules_end}"
            mods = f"Counts!H3:H{row_modules_end}"
            cat_i = f"{chr(71+i)}2"
            calc_ws.write_array_formula(iy+1, ix+i, iy+1+n_raw_modules, ix+i, fr'{{=IF(({cats}={cat_i})+({cats}={abbr}),{mods},"")}}')

        iy += n_raw_modules + 2
        ix -= 1
        off = 0
        for i, k in enumerate(modules + [""] * n_extra):
            if k == "":
                k = f"=Counts!H{3+off}"
                off += 1
            else:
                off += x._count if (x := self._prices.get(k)) else 0
            calc_ws.write(iy+i, ix, k)
            for j, cat in enumerate(categories.keys()):
                col = chr(71+j)
                col_range = f"{col}3:{col}{n_raw_modules+3}"
                mod = f"F{iy+i+1}"
                calc_ws.write(iy+i, ix+1+j, f"=COUNTIF({col_range},{mod})")

        # Setup InDesign formats
        format_pcs: xw.workbook.Format = wb.add_format({"num_format": "#,##0 [$pcs.]"})
        format_price: xw.workbook.Format = wb.add_format({"num_format": "#,##0.-"})
        format_dkk: xw.workbook.Format = wb.add_format({"num_format": "#,##0 [$DKK]"})

        # Write InDesign sheet
        stride = n_modules + 2
        for i, cat in enumerate(categories.keys()):
            y_top = i * stride
            indesign_ws.write(y_top, 0, cat)

            col = chr(71+i)
            mod_range = f"Calculations!F{n_raw_modules+4}:F{n_raw_modules+n_modules+3}"
            mod_range_2 = f"Counts!H3:H{row_modules_end}"
            price_per_range = f"Counts!K3:K{row_modules_end}"
            counts_range = f"Calculations!{col}{n_raw_modules+4}:{col}{n_raw_modules+n_modules+3}"

            indesign_ws.write_array_formula(y_top, 2, y_top, 2, f"=SUM(IFERROR(C{y_top+2}:C{y_top+n_modules+1}*B{y_top+2}:B{y_top+n_modules+1},0))", format_dkk)
            indesign_ws.write_array_formula(y_top+1, 0, y_top+n_modules, 0, f'{{=IFERROR(INDEX({mod_range},SMALL(IF({counts_range}<>0,ROW({mod_range})-{n_raw_modules+3}),ROW(1:{n_modules}))), "")}}')
            for j in range(n_modules):
                indesign_ws.write(y_top+j+1, 1, f'=IFERROR(INDEX({counts_range}, MATCH(A{y_top+j+2}, {mod_range}, 0)), "")', format_pcs)
                indesign_ws.write(y_top+j+1, 2, f'=IFERROR(INDEX({price_per_range}, MATCH(A{y_top+j+2}, {mod_range_2}, 0)), "")', format_price)

        row = stride * len(categories) + 2
        indesign_ws.write(row, 0, "Total Installation")
        indesign_ws.write_array_formula(row, 2, row, 2, f"{{=SUM(Calculations!D3:D{3+n_modules}+0)}}", format_dkk)

        row += 2
        row_totals = row_modules_end+len(categories)+1
        indesign_ws.write(row, 0, "in total ex VAT.")
        indesign_ws.write(row+1, 0, "VAT")
        indesign_ws.write(row+2, 0, "in total incl. VAT")
        indesign_ws.write(row, 2, f"=Counts!K{row_totals} + C{row-1}", format_dkk)
        indesign_ws.write(row+1, 2, f"=C{row+3}-C{row+1}", format_dkk)
        indesign_ws.write(row+2, 2, f"=Counts!L{row_totals} + C{row-1}", format_dkk)


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
