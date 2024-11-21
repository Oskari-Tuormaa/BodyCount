import adsk.fusion
import adsk.core

from adsk.fusion import BoundingBoxEntityTypes as bb_ent_types

from dataclasses import dataclass
from typing import Protocol


class SupportsBoundingBox2(Protocol):
    def boundingBox2(self, entityTypes: adsk.fusion.BoundingBoxEntityTypes) -> adsk.core.BoundingBox3D:
        ...


@dataclass
class Cube:
    vertices: adsk.fusion.CustomGraphicsCoordinates
    indices: list[int]
    normals: list[float]
    normalIndices: list[int]


def generate_cube(corner1: adsk.core.Point3D, corner2: adsk.core.Point3D) -> Cube:
    """Generates vertices, indices and normals for a cube going from corner1 to corner2."""
    # fmt: off
    vertices = adsk.fusion.CustomGraphicsCoordinates.create(
        [corner1.x, corner1.y, corner1.z,
         corner2.x, corner1.y, corner1.z,
         corner2.x, corner2.y, corner1.z,
         corner1.x, corner2.y, corner1.z,
         corner1.x, corner1.y, corner2.z,
         corner2.x, corner1.y, corner2.z,
         corner2.x, corner2.y, corner2.z,
         corner1.x, corner2.y, corner2.z]
    )

    # Define normals for each face
    normals = [
        0,  0, -1,  # Bottom face (pointing down)
        0,  0,  1,  # Top face (pointing up)
        0, -1,  0,  # Front face (pointing forward)
        0,  1,  0,  # Back face (pointing backward)
       -1,  0,  0,  # Left face (pointing left)
        1,  0,  0   # Right face (pointing right)
    ]

    indices = [
        # Bottom face
        0, 1, 2,
        0, 2, 3,

        # Top face
        4, 5, 6,
        4, 6, 7,

        # Front face
        0, 1, 5,
        0, 5, 4,

        # Back face
        3, 2, 6,
        3, 6, 7,

        # Left face
        0, 3, 7,
        0, 7, 4,

        # Right face
        1, 2, 6,
        1, 6, 5
    ]

    normalIndices = [
        # Bottom face (all use normal 0)
        0, 0, 0,
        0, 0, 0,

        # Top face (all use normal 1)
        1, 1, 1,
        1, 1, 1,

        # Front face (all use normal 2)
        2, 2, 2,
        2, 2, 2,

        # Back face (all use normal 3)
        3, 3, 3,
        3, 3, 3,

        # Left face (all use normal 4)
        4, 4, 4,
        4, 4, 4,

        # Right face (all use normal 5)
        5, 5, 5,
        5, 5, 5
    ]
    # fmt: on

    return Cube(vertices, indices, normals, normalIndices)


def generate_cube_bbox(comp: SupportsBoundingBox2) -> Cube:
    """Generates vertices, indices and normals of a cube, given an object that supports the boundingBox2 method."""
    bbox = comp.boundingBox2(
        bb_ent_types.MeshBodyBoundingBoxEntityType |
        bb_ent_types.SolidBRepBodyBoundingBoxEntityType |
        bb_ent_types.SurfaceBodyBoundingBoxEntityType
    )
    minPt = bbox.minPoint
    maxPt = bbox.maxPoint
    return generate_cube(minPt.copy(), maxPt.copy())