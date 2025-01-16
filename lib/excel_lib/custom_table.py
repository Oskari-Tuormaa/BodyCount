import re

from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet
from dataclasses import dataclass

from .helpers import column_to_index, index_to_column

@dataclass
class TableRef:
    """A more flexible representation of the table reference.
    
    Column and row data stored as integers.
    Has helper methods for easier manipulation of the table range.
    """
    col_start: int
    row_start: int
    col_end: int
    row_end: int

    header_row_count: int
    totals_row_count: int

    @classmethod
    def parse(cls, table: Table) -> "TableRef":
        """Creates a TableRef object from the given Table.
        
        @return The generated `TableRef` object.
        """
        pattern = re.compile(r"([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)")
        match = pattern.match(table.ref)
        col_start, row_start, col_end, row_end = match.group(1), match.group(2), match.group(3), match.group(4)
        return TableRef(
            column_to_index(col_start),
            int(row_start),
            column_to_index(col_end),
            int(row_end),
            table.headerRowCount if table.headerRowCount else 0,
            table.totalsRowCount if table.totalsRowCount else 0,
        )

    @property
    def n_rows(self) -> int:
        """The amount of rows in the table, excluding header and total rows."""
        return self.row_end - self.row_start - self.header_row_count - self.totals_row_count

    @property
    def n_cols(self) -> int:
        return self.col_end - self.col_start + 1

    def totals_row_cell_range(self) -> str | None:
        """Generates the cell range string for the "totals" row (eg. H1:I4).

        @return A string containing the range, or None if the table doesn't have a "totals" row.
        """
        if self.totals_row_count == 0:
            return None
        totals_row_start = self.row_end - self.totals_row_count + 1
        totals_row_end = self.row_end
        totals_col_start = index_to_column(self.col_start)
        totals_col_end = index_to_column(self.col_end)
        return f"{totals_col_start}{totals_row_start}:{totals_col_end}{totals_row_end}"

    def range_data_cols(self):
        """Returns range over columns of table.
        
        @return Range over the column indices referenced by the table.
        """
        return range(self.col_start, self.col_end+1)

    def range_data_rows(self):
        """Returns range over rows of table.
        
        @return Range over the row indices referenced by the table.
        """
        data_row_start = self.row_start + self.header_row_count
        data_row_end = self.row_end - self.totals_row_count
        return range(data_row_start, data_row_end + 1)

    def set_n_rows(self, n_rows: int):
        """Sets how many rows of data is contained in the table, keeping header and totals rows.
        
        @param n_rows How many rows of data the Table should contain.
        """
        self.row_end = self.row_start + self.header_row_count + self.totals_row_count + n_rows - 1

    def __str__(self) -> str:
        start_index = index_to_column(self.col_start)
        end_index = index_to_column(self.col_end)
        return f"{start_index}{self.row_start}:{end_index}{self.row_end}"


class CustomTable:
    """Wrapper for openpyxl Table objects, adding some custom functionality."""

    def __init__(self, worksheet: Worksheet, table: Table):
        self._worksheet = worksheet
        self._table = table
        self._ref = TableRef.parse(table)
    
    def clear_table_data(self):
        """Clears all data from the table, keeping header and total rows."""
        for i_row in self._ref.range_data_rows():
            for i_col in self._ref.range_data_cols():
                self._worksheet.cell(i_row, i_col).value = None
    
    def set_n_rows(self, n_rows: int):
        """Sets how many rows of data the table contains.

        @note This method automatically moves the "totals" row, such that it is
              preserved in it's new location following the resize of the table.
        
        @param n_rows How many rows of data the table should contain.
        """
        if self._ref.totals_row_count != 0:
            totals_row_cells = self._worksheet[self._ref.totals_row_cell_range()]

        self._ref.set_n_rows(n_rows)
        self._table.ref = str(self._ref)

        if self._ref.totals_row_count != 0:
            new_totals_row_cells = self._worksheet[self._ref.totals_row_cell_range()]
            for old_row, new_row in zip(totals_row_cells, new_totals_row_cells):
                for old_cell, new_cell in zip(old_row, new_row):
                    new_cell.value = old_cell.value
                    old_cell.value = None
    
    def set_data(self, data: list[list[str | int | None]]):
        """Fills the table with the given data.

        @note Data must be formatted, such that the first index is row and
              second index is column (eg. data[row][column]).

        @note The amount of columns in the given data must match the
              amount of columns in the table.
        
        @param data A 2d list containing the data to fill the table with.
        """
        self.clear_table_data()
        self.set_n_rows(len(data))
        for i_row, d_row in zip(self._ref.range_data_rows(), data):
            assert len(d_row) == self._ref.n_cols,\
                f"Amount of columns in data does not match amount of columns in table. {len(d_row)} != {self._ref.n_cols}"
            for i_col, d in zip(self._ref.range_data_cols(), d_row):
                self._worksheet.cell(i_row, i_col).value = d