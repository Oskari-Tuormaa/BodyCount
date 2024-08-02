import re


def tryfloat(s):
    """
    Return an float if possible, or `s` unchanged.
    """
    try:
        return float(s)
    except ValueError:
        return s


def alphanum_key(s):
    """
    Turn a string into a list of string and number chunks.

    >>> alphanum_key("z23a")
    ["z", 23, "a"]

    """
    if isinstance(s, list) or isinstance(s, tuple):
        return [alphanum_key(x) for x in s]
    
    if isinstance(s, int) or isinstance(s, float):
        return s

    key = [tryfloat(c) for c in re.split("([0-9]+)", s)]

    return key


def human_sort(l):
    """
    Sort a list in the way that humans expect.
    """
    l.sort(key=alphanum_key)
