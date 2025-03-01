import adsk.fusion
import adsk.core

from dataclasses import dataclass

from .. import counting_lib


app = adsk.core.Application.get()

@dataclass
class Mesh:
    vertices: list[float]
    indices: list[int]

    def add_mesh(self, other: "Mesh"):
        self.indices += list([i + len(self.vertices) // 3 for i in other.indices])
        self.vertices += other.vertices


# fmt: off
COLOR_OPACITY = 100
COLORS = [
    adsk.core.Color.create(0xe6, 0x19, 0x4b, COLOR_OPACITY),
    adsk.core.Color.create(0x3c, 0xb4, 0x4b, COLOR_OPACITY),
    adsk.core.Color.create(0xff, 0xe1, 0x19, COLOR_OPACITY),
    adsk.core.Color.create(0x43, 0x63, 0xd8, COLOR_OPACITY),
    adsk.core.Color.create(0xf5, 0x82, 0x31, COLOR_OPACITY),
    adsk.core.Color.create(0x91, 0x1e, 0xb4, COLOR_OPACITY),
    adsk.core.Color.create(0x46, 0xf0, 0xf0, COLOR_OPACITY),
    adsk.core.Color.create(0xf0, 0x32, 0xe6, COLOR_OPACITY),
    adsk.core.Color.create(0xbc, 0xf6, 0x0c, COLOR_OPACITY),
    adsk.core.Color.create(0xfa, 0xbe, 0xbe, COLOR_OPACITY),
    adsk.core.Color.create(0x00, 0x80, 0x80, COLOR_OPACITY),
    adsk.core.Color.create(0xe6, 0xbe, 0xff, COLOR_OPACITY),
    adsk.core.Color.create(0x9a, 0x63, 0x24, COLOR_OPACITY),
    adsk.core.Color.create(0xff, 0xfa, 0xc8, COLOR_OPACITY),
    adsk.core.Color.create(0x80, 0x00, 0x00, COLOR_OPACITY),
    adsk.core.Color.create(0xaa, 0xff, 0xc3, COLOR_OPACITY),
    adsk.core.Color.create(0x80, 0x80, 0x00, COLOR_OPACITY),
    adsk.core.Color.create(0xff, 0xd8, 0xb1, COLOR_OPACITY),
    adsk.core.Color.create(0x00, 0x00, 0x75, COLOR_OPACITY),
]
# ft: on


class SelectionGraphics:
    def __init__(self, color: adsk.core.Color):
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        rootComp = design.rootComponent

        self._cgGroup = rootComp.customGraphicsGroups.add()
        self._mesh: adsk.fusion.CustomGraphicsMesh | None = None
        self._colorEffect = adsk.fusion.CustomGraphicsShowThroughColorEffect.create(color, 0.4)

        self._cgGroup.isSelectable = False

    def __del__(self):
        self.deleteMe()

    def deleteMe(self):
        if self._cgGroup:
            self._cgGroup.deleteMe()

    def add_occ(self, occ: adsk.fusion.Occurrence):
        self.add_objs(*counting_lib.traverse_brepbodies(occ))

    def add_objs(self, *objs: adsk.fusion.BRepBody):
        added_geometry = Mesh([], [])

        for obj in objs:
            obj_mesh = obj.meshManager.displayMeshes.bestMesh
            
            added_geometry.add_mesh(
                Mesh(
                    obj_mesh.nodeCoordinatesAsFloat,
                    obj_mesh.nodeIndices,
                )
            )

        if self._mesh is not None:
            # We already have a mesh = Append the new mesh data and replace the old one
            original_mesh = Mesh(
                list(self._mesh.coordinates.coordinates).copy(),
                list(self._mesh.vertexIndexList).copy(),
            )
            added_geometry.add_mesh(original_mesh)

        # No mesh yet = Create a new one
        self._mesh = self._cgGroup.addMesh(
            adsk.fusion.CustomGraphicsCoordinates.create(added_geometry.vertices),
            added_geometry.indices,
            [], []
        )
        
        self._mesh.color = self._colorEffect

    def set_color(self, color: adsk.core.Color):
        self._colorEffect = adsk.fusion.CustomGraphicsShowThroughColorEffect.create(color, 0.4)
        if self._mesh is not None: 
            self._mesh.color = self._colorEffect


class SelectionGraphicsGroups:
    def __init__(self):
        self._groups: dict[str, SelectionGraphics] = {}
        self._color_usage = [[c, 0] for c in COLORS]
    
    def __del__(self):
        self.deleteMe()
    
    def _get_next_color(self) -> adsk.core.Color:
        for i in range(1000):
            for val in self._color_usage:
                if val[1] <= i:
                    val[1] += 1
                    return val[0]
        return COLORS[0]

    def deleteMe(self):
        for sGroup in self._groups.values():
            sGroup.deleteMe()
    
    def get(self, group: str) -> SelectionGraphics | None:
        if group in self._groups:
            return self._groups[group]
    
    def create(self, group: str) -> SelectionGraphics:
        if group in self._groups:
            self.delete(group)
        
        color = self._get_next_color()
        self._groups[group] = SelectionGraphics(color)
        return self._groups[group]
    
    def delete(self, group: str):
        if group in self._groups:
            self._groups.pop(group).deleteMe()
    
    def clear(self):
        self._color_usage = [[c, 0] for c in COLORS]
        for group in list(self._groups):
            self.delete(group)
    
    def rename(self, orig: str, new: str):
        if orig in self._groups:
            self._groups[new] = self._groups.pop(orig)
    
    @property
    def ngroups(self) -> int:
        return len(self._groups)
