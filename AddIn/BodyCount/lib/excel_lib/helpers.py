def column_to_index(col: str) -> int:
    """Converts column names to indices, starting from 1 (fx. A -> 1 or AB -> 28)."""
    idx = 0
    mul = 1
    for ch in col[::-1]:
        idx += mul * (ord(ch.upper()) - ord('A') + 1)
        mul *= 26
    return idx

def index_to_column(index: int) -> str:
    """Converts indices to column names, starting from 1 (fx. 1 -> A or 28 -> AB)."""
    col = ""
    while index > 0:
        nxt = (index-1) % 26 + 1
        col = chr(nxt + ord('A') - 1) + col
        index = (index - nxt) // 26
    return col
