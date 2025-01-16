import re

from typing import TypeVar
from collections.abc import Callable


def tryfloat(s: str) -> str | float:
    """
    Return an float if possible, or `s` unchanged.
    """
    try:
        return float(s)
    except ValueError:
        return s


def alphanum_key(s: str) -> list[str | float]:
    """
    Turn a string into a list of string and number chunks.

    >>> alphanum_key("z23a")
    ["z", 23, "a"]

    """
    if isinstance(s, list) or isinstance(s, tuple):
        return [alphanum_key(x) for x in s]
    
    if isinstance(s, int) or isinstance(s, float):
        return s

    return [tryfloat(c) for c in re.split("([0-9]+)", s)]


T = TypeVar('T')
def human_sort(l: list[T], key: Callable[[T], str] | None = None):
    """
    Sort a list in the way that humans expect.
    """
    if not key:
        l.sort(key=alphanum_key)
    else:
        l.sort(key=lambda x: alphanum_key(key(x)))
