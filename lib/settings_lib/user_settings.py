from pathlib import Path
from serde import serde
from serde.json import to_json, from_json

USER_SETTINGS_FILE = Path(__file__).parent.parent.parent/'.user-settings.json'

@serde
class UserData:
    shared_data_path: Path | None = None
    overwrite: bool = True

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
            cached_user_data = from_json(UserData, fd.read())
            cached_user_data_time = USER_SETTINGS_FILE.stat().st_mtime
    return cached_user_data

def save_user_data(user_data: UserData):
    with USER_SETTINGS_FILE.open('w') as fd:
        fd.write(to_json(user_data, indent=4))
