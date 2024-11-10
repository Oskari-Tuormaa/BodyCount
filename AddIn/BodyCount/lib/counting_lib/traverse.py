import adsk.fusion
import re

from typing import Generator
from collections.abc import Callable

from .human_sort import human_sort
from ..excel_lib import Body, Module

def traverse_occurrences(
    root: adsk.fusion.Occurrence | adsk.fusion.Component,
    predicate: Callable[[adsk.fusion.Occurrence], bool] | None = None,
    depth: int | None = None,
) -> Generator[list[adsk.fusion.Occurrence | adsk.fusion.Component], None, None]:
    """Traverses and yields every visible occurrence under root.
    
    @param root The root component or occurrence, from which to start traversing.
    @param predicate Only returns matches where predicate returns True.
    @param depth Maximum depth of recursion. If not given, function will recurse as deep as possible.

    @return A generator yielding all visible Occurences under `root` for which `predicate` returns true,
            up to a recursion depth of `depth`. The returned value is a list containing the full path
            of the Occurrence, with the root object at the first index, and the found Occurrence as the last.
    """
    def _traverse_occurrences_inner(
        tree: list[adsk.fusion.Occurrence | adsk.fusion.Component],
        root: adsk.fusion.Occurrence | adsk.fusion.Component,
        predicate: Callable[[adsk.fusion.Occurrence], bool] | None = None,
        depth: int | None = None,
    ) -> Generator[list[adsk.fusion.Occurrence | adsk.fusion.Component], None, None]:
        iter = (
            root.childOccurrences
            if isinstance(root, adsk.fusion.Occurrence)
            else root.occurrences
        )
        for occ in iter:
            if not occ.isVisible:
                continue

            if depth is None:
                yield from _traverse_occurrences_inner([*tree, occ], occ, predicate=predicate, depth=None)
            elif not depth <= 0:
                yield from _traverse_occurrences_inner([*tree, occ], occ, predicate=predicate, depth=depth - 1)

            if predicate is None or predicate(occ):
                yield [*tree, occ]
    return _traverse_occurrences_inner([root], root, predicate, depth)


def traverse_brepbodies(
    root: adsk.fusion.Occurrence | adsk.fusion.Component,
) -> Generator[tuple[list[adsk.fusion.Occurrence | adsk.fusion.Component], adsk.fusion.BRepBody], None, None]:
    """Traverses and yields every visible bRepBody under root.
    
    @param root The root component or occurrence, from which to start traversing.

    @return A generator yielding tuples of all visible bRepBodies under root, and it's path
            in the component tree.
    """
    yield from [([root], body) for body in root.bRepBodies if body.isVisible]
    for branch in traverse_occurrences(root):
        yield from [([*branch], body) for body in branch[-1].component.bRepBodies if body.isVisible]

def filter_name(name: str) -> str:
    OCC_NAME_FILTERS = [
        (r"^(.*) v\d+$", r"\1"),
        (r"^(.*) v\d+:\d+$", r"\1"),
    ]
    for pat, repl in OCC_NAME_FILTERS:
        name = re.sub(pat, repl, name)
    return name

def collect_bodies_under(root: adsk.fusion.Component | adsk.fusion.Occurrence) -> list[Body]:
    bodies_dict: dict[tuple[str, str], Body] = {}
    for (branch, body) in traverse_brepbodies(root):
        name = filter_name(body.name)
        if hasattr(body, "material"):
            material = body.material.name
        else:
            material = ""
        key = (name, material)

        if key not in bodies_dict:
            bodies_dict[key] = Body(name, 0, material)
        bodies_dict[key].count += 1
    
    # Sort Body list based on the name
    bodies_list = list(bodies_dict.values())
    human_sort(bodies_list, key=lambda x: x.name)
    return bodies_list


def collect_modules_under(root: adsk.fusion.Component | adsk.fusion.Occurrence) -> list[Module]:
    modules: list[Module] = []

    for branch in traverse_occurrences(root, depth=0):
        modules.append(Module(
            "",
            filter_name(branch[-1].name),
            collect_bodies_under(branch[-1])
        ))

    return modules
