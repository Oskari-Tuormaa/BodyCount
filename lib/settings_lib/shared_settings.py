from .user_settings import load_user_data, UserData

from serde import serde
from serde.json import to_json, from_json
from dataclasses import field
from pathlib import Path

@serde
class SharedData:
    detail_materials: list[str] = field(default_factory=lambda: [])
    wood_materials: list[str] = field(default_factory=lambda: [])

cached_shared_data: SharedData | None = None
cached_shared_data_time: float = 0

def get_shared_data_path() -> Path:
    user_data = load_user_data()
    if user_data.shared_data_path is None:
        raise RuntimeError('Cannot load shared data, since shared data path is not set in user settings!')
    return Path(user_data.shared_data_path)/'Vermland'/'The Collection'/'BodyCount'/'shared_settings.json'

def load_shared_data() -> SharedData:
    global cached_shared_data, cached_shared_data_time
    shared_data_path = get_shared_data_path()

    if not shared_data_path.exists():
        save_shared_data(cached_shared_data := SharedData())
        cached_shared_data_time = shared_data_path.stat().st_mtime
        return cached_shared_data

    if (
        cached_shared_data is None or
        cached_shared_data_time != shared_data_path.stat().st_mtime
    ):
        with shared_data_path.open('r') as fd:
            cached_shared_data = from_json(SharedData, fd.read())
            cached_shared_data_time = shared_data_path.stat().st_mtime
    return cached_shared_data

def save_shared_data(shared_data: SharedData):
    shared_data_path = get_shared_data_path()
    shared_data_path.parent.mkdir(parents=True, exist_ok=True)
    with shared_data_path.open('w') as fd:
        fd.write(to_json(shared_data, indent=4))