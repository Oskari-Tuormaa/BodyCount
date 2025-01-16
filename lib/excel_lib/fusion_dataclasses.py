from dataclasses import dataclass

@dataclass
class Body:
    name: str
    count: int
    material: str

@dataclass
class Module:
    category: str
    name: str
    bodies: list[Body]