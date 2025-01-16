from .user_settings import load_user_data, UserData
from .util import Serializable

from dataclasses import dataclass, field
from pathlib import Path

@dataclass
class SharedData(Serializable):
    detail_materials: list[str] = field(default_factory=lambda: [])

cached_shared_data: SharedData | None = None
cached_shared_data_time: float = 0

def get_shared_data_path() -> Path:
    user_data = load_user_data()
    if user_data.shared_data_path is None:
        raise RuntimeError('Cannot load shared data, since shared data path is not set in user settings!')
    return Path(user_data.shared_data_path)

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
            cached_shared_data = SharedData.deserialize(fd.read())
            cached_shared_data_time = shared_data_path.stat().st_mtime
    return cached_shared_data

def save_shared_data(shared_data: SharedData):
    shared_data_path = get_shared_data_path()
    with shared_data_path.open('w') as fd:
        fd.write(shared_data.serialize())