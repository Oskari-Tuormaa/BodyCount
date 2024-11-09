import adsk.fusion

from typing import Generator
from collections.abc import Callable

from ..excel_lib import Body, Module

def traverse_occurrences(
    root: adsk.fusion.Occurrence | adsk.fusion.Component,
    predicate: Callable[[adsk.fusion.Occurrence], bool] | None = None,
    depth: int | None = None,
) -> Generator[adsk.fusion.Occurrence, None, None]:
    """Traverses and yields every visible occurrence under root.
    
    @param root The root component or occurrence, from which to start traversing.
    @param predicate Only returns matches where predicate returns True.
    @param depth Maximum depth of recursion. If not given, function will recurse as deep as possible.

    @return A generator yielding all visible Occurences under `root` for which `predicate` returns true, up to a recursion depth of `depth`.
    """

    iter = (
        root.childOccurrences
        if isinstance(root, adsk.fusion.Occurrence)
        else root.occurrences
    )
    for occ in iter:
        if not occ.isVisible:
            continue

        if depth is None:
            yield from traverse_occurrences(occ, predicate=predicate, depth=None)
        elif not depth <= 0:
            yield from traverse_occurrences(occ, predicate=predicate, depth=depth - 1)

        if predicate is None or predicate(occ):
            yield occ


def traverse_brepbodies(
    root: adsk.fusion.Occurrence | adsk.fusion.Component,
) -> Generator[adsk.fusion.BRepBody, None, None]:
    """Traverses and yields every visible bRepBody under root.
    
    @param root The root component or occurrence, from which to start traversing.

    @return A generator yielding all visible bRepBodies under root.
    """
    yield from [body for body in root.bRepBodies if body.isVisible]
    for occ in traverse_occurrences(root):
        yield from [body for body in occ.component.bRepBodies if body.isVisible]


def collect_bodies_under(root: adsk.fusion.Component) -> list[Body]:
    bodies_dict: dict[tuple[str, str], Body] = {}
    for body in traverse_brepbodies(root):
        name = body.name
        if hasattr(body, "material"):
            material = body.material.name
        else:
            material = ""
        key = (name, material)

        if key not in bodies_dict:
            bodies_dict[key] = Body(name, 0, material)
        bodies_dict[key].count += 1
    return list(bodies_dict.values())