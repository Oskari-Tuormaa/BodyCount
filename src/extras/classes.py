import adsk.core, adsk.fusion, adsk.cam, traceback
import re

from typing import Callable, Generator, List, Dict, Union
from pathlib import Path
from .human_sort import human_sort

import importlib
import os, sys

packagepath = os.path.join(os.path.dirname(sys.argv[0]), "Lib", "site-packages")
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

def filter_name(name: str) -> str:
    OCC_NAME_FILTERS = [
        (r"^(.*) v\d+$", r"\1"),
    ]
    for pat, repl in OCC_NAME_FILTERS:
        name = re.sub(pat, repl, name)
    return name


def add_bodies_to_worksheet(root: adsk.fusion.Component, wb: xw.Workbook):
    # Get sheets
    if (ws := wb.get_worksheet_by_name("Counts")) is None:
        ws = wb.add_worksheet("Counts")

    # Count BRepBodies
    counts = {}
    for body in traverse_brepbodies(root):
        name = filter_name(body.name)
        if name in counts:
            counts[name][0] += 1
        else:
            counts[name] = [1, body]

    # Count FSA
    for occ in traverse_occurrences(root, predicate=lambda x: "FSA" in x.name):
        name = filter_name(occ.component.name)
        if name in counts:
            counts[name][0] += 1
        else:
            counts[name] = [1, occ]

    # Add headers
    ws.write(1, 1, "Body")
    ws.write(1, 2, "Count")
    ws.write(1, 3, "Material")
    ws.write(1, 4, "Price")

    # Add counts
    keys = list(counts.keys())
    human_sort(keys)
    for i, name in enumerate(keys):
        count, body = counts[name]
        ws.write(2 + i, 1, name)
        ws.write(2 + i, 2, count)
        if hasattr(body, "material"):
            ws.write(2 + i, 3, body.material.name)
        if hasattr(body, "description") and body.desciption != "":
            try:
                price = float(body.description)
            except (ValueError, AttributeError):
                price = "ERROR IN DESCRIPTION"
            ws.write(2 + i, 4, price)


def add_components_to_worksheet(root: adsk.fusion.Component, wb: xw.Workbook):
    # Get sheets
    if (count_ws := wb.get_worksheet_by_name("Counts")) is None:
        count_ws = wb.add_worksheet("Counts")
    if (calc_ws := wb.get_worksheet_by_name("Calculations")) is None:
        calc_ws = wb.add_worksheet("Calculations")
    if (indesign_ws := wb.get_worksheet_by_name("InDesign")) is None:
        indesign_ws = wb.add_worksheet("InDesign")

    def get_price(occ: adsk.fusion.Occurrence) -> float:
        try:
            price = float(occ.component.description)
        except (ValueError, AttributeError):
            price = 0
        return price + sum(get_price(other) for other in occ.childOccurrences if other.isLightBulbOn)

    # Get prices
    prices = {}
    for occ in traverse_occurrences(root, depth=0):
        name = filter_name(occ.component.name)
        if name in prices:
            prices[name][0] += occ.isLightBulbOn
        else:
            prices[name] = [1, get_price(occ)]

    # Populate headers
    count_ws.write(1, 7, "Modules")
    count_ws.write(1, 8, "Category")
    count_ws.write(1, 9, "Price per")
    count_ws.write(1, 10, "Price ex. VAT")
    count_ws.write(1, 11, "Price incl. VAT")
    count_ws.write(1, 13, "Installation type")
    count_ws.write(1, 14, "Price per")
    count_ws.write(1, 16, "Category")
    count_ws.write(1, 17, "Abbreviation")
    calc_ws.write(1, 1, "Module")
    calc_ws.write(1, 2, "Installation type")
    calc_ws.write(1, 3, "Price")

    # Set installation types and headers
    installation_types = {
        "CA": 2091,
        "Li": 1776,
        "Is": 3552,
        "Be": 1776,
        "CAA": 2973,
        "LIA": 2658,
        "CAAA": 3855,
        "ISA": 4434,
        "ISAA": 5274,
        "JA": 579,
        "BA": 1776,
        "BAA": 2658,
        "WD": 2091,
        "T1": 1000,
        "T2": 1000,
        "T3": 1000,
        "T4": 500,
        "HA": 2000,
    }
    keys = list(installation_types.keys())
    human_sort(keys)
    for i, k in enumerate(keys):
        cost = installation_types.get(k)
        y = 2 + i
        count_ws.write(y, 13, k)
        count_ws.write(y, 14, cost)

    # Set Category abbreviations
    categories = [
        "String 1",
        "String 2",
        "String 3",
        "String 4",
        "String 5",
        "Island",
        "Wardrobe 1",
        "Wardrobe 2",
        "Wardrobe 3",
        "Wardrobe 4",
        "Wardrobe 5",
        "Bathroom 1",
        "Bathroom 2",
        "Bathroom 3",
        "Bathroom 4",
        "Bathroom 5",
        ]
    for i, cat in enumerate(categories):
        count_ws.write(2 + i, 16, cat)
        count_ws.write(2 + i, 17, cat)

    n_extra = 5

    # Add data per Price
    modules = list(prices.keys())
    human_sort(modules)
    row_modules_end = 2
    for mod in modules + [""] * n_extra:
        v = prices.get(mod)

        n = 1 if v is None else v[0]
        per = 0 if v is None else v[1]

        for i in range(n):
            count_ws.write(row_modules_end, 7, mod)
            count_ws.data_validation(
                row_modules_end,
                8,
                row_modules_end,
                8,
                {"validate": "list", "source": "=$R$3:$R$90"},
            )
            count_ws.write(row_modules_end, 9, int(per))
            count_ws.write(
                row_modules_end,
                10,
                f"=ROUND(J{row_modules_end+1}*2, 0)",
            )
            count_ws.write(row_modules_end, 11, f"=ROUND(K{row_modules_end+1}*1.25, 0)")

            calc_ws.write(row_modules_end, 1, f"=Counts!H{row_modules_end+1}")
            calc_ws.write(
                row_modules_end,
                2,
                f'=IFERROR(LEFT(RIGHT(B{row_modules_end+1}, LEN(B{row_modules_end+1}) - SEARCH("-", B{row_modules_end+1})), SEARCH("-", RIGHT(B{row_modules_end+1}, LEN(B{row_modules_end+1}) - SEARCH("-", B{row_modules_end+1})))-1), "")',
                #f'=IFERROR(MID(B{row_modules_end+1}, FIND("-", B{row_modules_end+1})+1, 2),"")',
            )
            calc_ws.write(
                row_modules_end,
                3,
                f"=IFERROR(INDEX(Counts!O3:O100, MATCH(C{row_modules_end+1}, Counts!N3:N100)), 0)",
            )
            row_modules_end += 1

    n_modules = len(modules) + n_extra
    for idx, c in enumerate(categories):
        de_row = row_modules_end
        abbr_row = 3 + idx
        row = idx + de_row

        cat = f"I{row+1}"
        abbr = f"R{abbr_row}"
        cats = f"I3:I{de_row}"
        costs = f"J3:J{de_row}"
        costs_excl_vat = f"K3:K{de_row}"
        costs_incl_vat = f"L3:L{de_row}"

        count_ws.write(row, 8, c)
        count_ws.write_array_formula(
            row,
            9,
            row,
            9,
            rf"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs}, 0))}}",
        )
        count_ws.write_array_formula(
            row,
            10,
            row,
            10,
            rf"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs_excl_vat}, 0))}}",
        )
        count_ws.write_array_formula(
            row,
            11,
            row,
            11,
            rf"{{=SUM(IF(({cats} = {cat}) + ({cats} = {abbr}), {costs_incl_vat}, 0))}}",
        )

    row = row_modules_end + len(categories)
    count_ws.write(row, 8, "Total")
    count_ws.write(row, 9, f"=SUM(J{row-len(categories)+1}:J{row})")
    count_ws.write(row, 10, f"=SUM(K{row-len(categories)+1}:K{row})")
    count_ws.write(row, 11, f"=SUM(L{row-len(categories)+1}:L{row})")

    # Write category counts
    ix = 6
    iy = 1
    n_raw_modules = row_modules_end - 3
    for i, cat in enumerate(categories):
        calc_ws.write(iy, ix + i, cat)

        abbr = f"Counts!R{3+i}"
        cats = f"Counts!I3:I{row_modules_end}"
        mods = f"Counts!H3:H{row_modules_end}"
        cat_i = f"{chr(71+i)}2"
        calc_ws.write_array_formula(
            iy + 1,
            ix + i,
            iy + 1 + n_raw_modules,
            ix + i,
            rf'{{=IF(({cats}={cat_i})+({cats}={abbr}),{mods},"")}}',
        )

    iy += n_raw_modules + 2
    ix -= 1
    off = 0
    for i, k in enumerate(modules + [""] * n_extra):
        if k == "":
            k = f"=Counts!H{3+off}"
            off += 1
        else:
            off += x[0] if (x := prices.get(k)) else 0
        calc_ws.write(iy + i, ix, k)
        for j, cat in enumerate(categories):
            col = chr(71 + j)
            col_range = f"{col}3:{col}{n_raw_modules+3}"
            mod = f"F{iy+i+1}"
            calc_ws.write(iy + i, ix + 1 + j, f"=COUNTIF({col_range},{mod})")

    # Setup InDesign formats
    format_pcs: xw.workbook.Format = wb.add_format({"num_format": "#,##0 [$pcs.]"})
    format_price: xw.workbook.Format = wb.add_format({"num_format": "#,##0.-"})
    format_dkk: xw.workbook.Format = wb.add_format({"num_format": "#,##0 [$DKK]"})

    # Write InDesign sheet
    stride = n_modules + 2
    for i, cat in enumerate(categories):
        y_top = i * stride
        indesign_ws.write(y_top, 0, cat)

        col = chr(71 + i)
        mod_range = f"Calculations!F{n_raw_modules+4}:F{n_raw_modules+n_modules+3}"
        mod_range_2 = f"Counts!H3:H{row_modules_end}"
        price_per_range = f"Counts!K3:K{row_modules_end}"
        counts_range = (
            f"Calculations!{col}{n_raw_modules+4}:{col}{n_raw_modules+n_modules+3}"
        )

        indesign_ws.write_array_formula(
            y_top,
            2,
            y_top,
            2,
            f"=SUM(IFERROR(C{y_top+2}:C{y_top+n_modules+1}*B{y_top+2}:B{y_top+n_modules+1},0))",
            format_dkk,
        )
        indesign_ws.write_array_formula(
            y_top + 1,
            0,
            y_top + n_modules,
            0,
            f'{{=IFERROR(INDEX({mod_range},SMALL(IF({counts_range}<>0,ROW({mod_range})-{n_raw_modules+3}),ROW(1:{n_modules}))), "")}}',
        )
        for j in range(n_modules):
            indesign_ws.write(
                y_top + j + 1,
                1,
                f'=IFERROR(INDEX({counts_range}, MATCH(A{y_top+j+2}, {mod_range}, 0)), "")',
                format_pcs,
            )
            indesign_ws.write(
                y_top + j + 1,
                2,
                f'=IFERROR(INDEX({price_per_range}, MATCH(A{y_top+j+2}, {mod_range_2}, 0)), "")',
                format_price,
            )

    row = stride * len(categories) + 2
    indesign_ws.write(row, 0, "Total Installation")
    indesign_ws.write_array_formula(
        row, 2, row, 2, f"{{=SUM(Calculations!D3:D{3+n_modules}+0)}}", format_dkk
    )

    row += 2
    row_totals = row_modules_end + len(categories) + 1
    indesign_ws.write(row, 0, "in total ex VAT.")
    indesign_ws.write(row + 1, 0, "VAT")
    indesign_ws.write(row + 2, 0, "in total incl. VAT")
    indesign_ws.write(row, 2, f"=Counts!K{row_totals} + C{row-1}", format_dkk)
    indesign_ws.write(row + 1, 2, f"=C{row+3}-C{row+1}", format_dkk)
    indesign_ws.write(row + 2, 2, f"=Counts!L{row_totals} + C{row-1}", format_dkk)


def traverse_occurrences(
    root: Union[adsk.fusion.Occurrence, adsk.fusion.Component],
    predicate: Union[Callable[[adsk.fusion.Occurrence], bool], None] = None,
    depth: Union[int, None] = None,
) -> Generator[adsk.fusion.Occurrence, None, None]:
    """Traverses and yields every visible occurrence under root"""

    iter = (
        root.childOccurrences
        if isinstance(root, adsk.fusion.Occurrence)
        else root.occurrences
    )
    for occ in iter:
        if depth is None:
            yield from traverse_occurrences(occ, predicate=predicate, depth=None)
        elif not depth <= 0:
            yield from traverse_occurrences(occ, predicate=predicate, depth=depth - 1)

        if not occ.isVisible:
            continue

        if predicate is None or predicate(occ):
            yield occ


def traverse_brepbodies(
    root: adsk.fusion.Component,
) -> Generator[adsk.fusion.BRepBody, None, None]:
    """Traverses and yields every visible bRepBody under root"""
    yield from [body for body in root.bRepBodies if body.isVisible]
    for occ in traverse_occurrences(root):
        yield from [body for body in occ.component.bRepBodies]
