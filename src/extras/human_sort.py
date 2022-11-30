import re


def tryint(s):
    """
    Return an int if possible, or `s` unchanged.
    """
    try:
        return int(s)
    except ValueError:
        return s


def alphanum_key(s):
    """
    Turn a string into a list of string and number chunks.

    >>> alphanum_key("z23a")
    ["z", 23, "a"]

    """
    key = [tryint(c) for c in re.split("([0-9]+)", s)]

    if len(key) >= 2 and isinstance(key[1], int) and key[1] == 9:
        key[0] = chr(1000)

    return key


def human_sort(l):
    """
    Sort a list in the way that humans expect.
    """
    l.sort(key=alphanum_key)
