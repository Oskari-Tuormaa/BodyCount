import adsk.fusion
import adsk.core

from .generate_cube import generate_cube_bbox, SupportsBoundingBox


app = adsk.core.Application.get()


class SelectionGraphics:
    def __init__(self, color: adsk.core.Color):
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        rootComp = design.rootComponent

        self._cgGroup = rootComp.customGraphicsGroups.add()
        self._meshes: list[adsk.fusion.CustomGraphicsMesh] = []
        self._colorEffect = adsk.fusion.CustomGraphicsSolidColorEffect.create(color)

        self._cgGroup.isSelectable = False

    def __del__(self):
        self.deleteMe()

    def deleteMe(self):
        if self._cgGroup:
            self._cgGroup.deleteMe()

    def add_obj(self, obj: SupportsBoundingBox):
        cube = generate_cube_bbox(obj)
        mesh = self._cgGroup.addMesh(
            cube.vertices, cube.indices, cube.normals, cube.normalIndices
        )
        mesh.color = self._colorEffect
        self._meshes.append(mesh)

    def set_color(self, color: adsk.core.Color):
        self._colorEffect = adsk.fusion.CustomGraphicsSolidColorEffect.create(color)
        for mesh in self._meshes:
            mesh.color = self._colorEffect


class SelectionGraphicsGroups:
    def __init__(self):
        self._groups: dict[str, SelectionGraphics] = {}
    
    def __del__(self):
        self.deleteMe()

    def deleteMe(self):
        for sGroup in self._groups.values():
            sGroup.deleteMe()
    
    def get(self, group: str) -> SelectionGraphics | None:
        if group in self._groups:
            return self._groups[group]
    
    def create(self, group: str, *args, **kwargs) -> SelectionGraphics:
        if group in self._groups:
            self.delete(group)

        self._groups[group] = SelectionGraphics(*args, **kwargs)
        return self._groups[group]
    
    def delete(self, group: str):
        if group in self._groups:
            self._groups.pop(group).deleteMe()
    
    def clear(self):
        for group in list(self._groups):
            self.delete(group)
    
    @property
    def ngroups(self) -> int:
        return len(self._groups)
