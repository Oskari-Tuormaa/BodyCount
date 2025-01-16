from .util import Serializable

from dataclasses import dataclass
from pathlib import Path

USER_SETTINGS_FILE = Path(__file__).parent.parent.parent/'.user-settings.json'

@dataclass
class UserData(Serializable):
    shared_data_path: str | None = None
    advanced_mode: bool = False

cached_user_data: UserData | None = None
cached_user_data_time: float = 0

def load_user_data() -> UserData:
    global cached_user_data, cached_user_data_time
    if not USER_SETTINGS_FILE.exists():
        save_user_data(cached_user_data := UserData())
        cached_user_data_time = USER_SETTINGS_FILE.stat().st_mtime
        return cached_user_data

    if (
        cached_user_data is None or
        cached_user_data_time != USER_SETTINGS_FILE.stat().st_mtime
    ):
        with USER_SETTINGS_FILE.open('r') as fd:
            cached_user_data = UserData.deserialize(fd.read())
            cached_user_data_time = USER_SETTINGS_FILE.stat().st_mtime
    return cached_user_data

def save_user_data(user_data: UserData):
    with USER_SETTINGS_FILE.open('w') as fd:
        fd.write(user_data.serialize())
