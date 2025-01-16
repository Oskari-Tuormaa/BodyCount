import json

from dataclasses import asdict

class Serializable:
    @classmethod
    def deserialize(cls, data: str):
        return cls(**json.loads(data))

    def serialize(self) -> str:
        return json.dumps(asdict(self), indent=4)