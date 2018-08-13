import collections
import warnings
from copy import copy

from cellbase.helper import DAO


def set_cell_value(cell, value):
    cell.value = value


class Celltable:
    """
    Celltable is equivalent to :class:`openpyxl.worksheet.Worksheet` which store the :class:`openpyxl.cell.Cell`
    """
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.col_ids = [col_id for col_id in worksheet[1]
                        if col_id.value is not None]
        self.cols = {col.value: [] for col in self.col_ids}
        self.rows = collections.OrderedDict()
        for row in worksheet.iter_rows(min_row=2):
            row_idx = row[0].row
            cells_in_row = {}
            for col_id in self.col_ids:
                cell = row[col_id.col_idx - 1]  # -1 as row is list(0 indexed)
                self.cols[col_id.value].append(cell)
                cells_in_row[col_id.value] = cell
            self.rows[row_idx] = cells_in_row

    def col_idx_to_col_id(self, col_idx):
        """
        Get column id cell with column index

        :param col_idx: Column index
        :type col_idx: int
        :return: Column id cell
        :rtype: openpyxl.cell.Cell
        """
        return self.worksheet[1][col_idx - 1]

    def safe_append(self, iterable, first_row=False):
        """
        Ensure new row appended on last row by setting worksheet._current_row,
        while preserving the original value of worksheet._current_row.

        .. note:: Set first_row to true to explicitly append to first row as worksheet.max_row always return 1

        :param iterable: Columns of data to append
        :param first_row: Explicitly append to first row
        :type first_row: bool
        """
        orig_current_row = self.worksheet._current_row
        # row_idx = worksheet._current_row + 1, see worksheet.append
        self.worksheet._current_row = self.worksheet.max_row if not first_row else 0
        self.worksheet.append(iterable)
        self.worksheet._current_row = orig_current_row

    def row_idxs_where(self, where=None):
        """
        Find the row indexes where any of the conditions match

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Row indexes where conditions match
        :rtype: list
        """
        if where is None:
            return [row_idx for row_idx in self.rows]
        row_idxs = []
        if DAO.COL_ROW_IDX in where:
            cond = where[DAO.COL_ROW_IDX]
            if callable(where[DAO.COL_ROW_IDX]):
                for row_idx in self.rows:
                    if cond(row_idx):
                        row_idxs.append(row_idx)
            else:
                row_idx = int(cond)
                if row_idx in self.rows:
                    row_idxs.append(row_idx)
        for col_name, cond in where.items():
            if col_name == DAO.COL_ROW_IDX:
                continue
            for cell in self.cols[col_name]:
                if cell.row not in row_idxs and cond(cell.value) if callable(cond) else cell.value == cond:
                    row_idxs.append(cell.row)

        return row_idxs

    def col_names_where(self, row_idx, where=None):
        """
        Find the column names where conditions match from a specific row

        :param row_idx: Row index to inspect
        :type row_idx: int
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Column id cell values where condition match
        :rtype: list
        """
        if where is None:
            return [col_name for col_name in self.cols]
        col_names = []
        for col_name, cond in where.items():
            if col_name == DAO.COL_ROW_IDX:
                if cond(row_idx) if callable(cond) else row_idx == int(cond):
                    col_names.append(col_name)
                continue
            cell = self.rows[row_idx][col_name]
            if cond(cell.value) if callable(cond) else cell.value == cond:
                col_names.append(col_name)
        return col_names

    def row_and_col_where(self, where=None):
        """
        Find row indexes where all conditions match by combining row_idx_where and col_names_where

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Row indexes where all conditions match
        :rtype: list
        """
        row_idxs_where = self.row_idxs_where(where)
        if where is None:
            return row_idxs_where
        row_idxs = []
        for row_idx in row_idxs_where:
            if len(self.col_names_where(row_idx, where)) == len(where):
                row_idxs.append(row_idx)
        return row_idxs

    def query(self, where=None):
        """
        Query data where conditions match

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: List of rows
        :rtype: list
        """
        rows_to_return = []
        for row_idx in self.row_and_col_where(where):
            values = {DAO.COL_ROW_IDX: row_idx}
            for key, cell in self.rows[row_idx].items():
                values[key] = cell.value
            rows_to_return.append(values)
        return rows_to_return

    def insert(self, value_in_dict):
        """
        Insert new row of data

        :param value_in_dict: Value of row in dict corresponding to col_names
        :type value_in_dict: dict
        :return: New row index
        :rtype: int
        """
        self.safe_append({col_id.col_idx: value_in_dict[col_id.value] for col_id in self.col_ids})
        new_row_idx = self.worksheet.max_row
        self.rows[new_row_idx] = {}
        for col_id in self.col_ids:
            new_cell = self.worksheet._cells[new_row_idx, col_id.col_idx]
            self.rows[new_row_idx][col_id.value] = new_cell
            self.cols[col_id.value].append(new_cell)
        return new_row_idx

    def update(self, value_in_dict, where=None):
        """
        Update row(s) where conditions match

        :param value_in_dict:
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Number of rows updated
        :rtype: int
        """
        if where is None:
            row = self.rows[value_in_dict[DAO.COL_ROW_IDX]]
            for cell in list(row.values())[1:]:
                cell.value = value_in_dict[self.col_idx_to_col_id(cell.col_idx).value]
            return 1
        return self.traverse(
            lambda cell: set_cell_value(cell, value_in_dict[self.col_idx_to_col_id(cell.col_idx).value]),
            where=where,
            select=[col_id.value for col_id in self.col_ids if col_id.value in value_in_dict]
        )

    def delete(self, where=None):
        """
        Delete row(s) of data where conditions match

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Number of rows deleted
        :rtype:int
        """
        row_idxs_where = self.row_and_col_where(where)
        affected_row_count = len(row_idxs_where)
        if affected_row_count is 0:
            return
        begin_max_row = self.worksheet.max_row  # max_row before delete
        # Pop row where condition matched
        for row_idx in row_idxs_where:
            self.rows.pop(row_idx)
        # Fill the gap, by changing key of rows starting from first popped row id
        first_popped_row_id = min(row_idxs_where)
        index_range = range(first_popped_row_id, begin_max_row + 1)  # +1 for range exclusive
        rows_after_first_popped_row = list(self.rows.values())[
                                      first_popped_row_id - 2:]  # -1 for col_id -1 for 0 indexed list
        for new_row_idx, row in zip(index_range, rows_after_first_popped_row):
            for col_id in self.col_ids:
                row[col_id.value].row = new_row_idx
            self.rows[new_row_idx] = row
        # Pop the last nth rows as changing key of dict may left old entry remains
        for last_row_idx in [begin_max_row - i for i in range(affected_row_count) if
                             begin_max_row - i not in row_idxs_where]:
            self.rows.pop(last_row_idx)
        # Sort dict by key as changing of
        # old key to empty(deleted) key may be treated as putting new entry
        # while delete() highly dependant on the sequence
        self.rows = collections.OrderedDict(sorted(self.rows.items()))
        # Update cols as the reference of cell is broken &
        # coordinate of cells to worksheet as worksheet._cells is not OrderedDict
        for col_id in self.col_ids:
            self.cols[col_id.value].clear()
        self.worksheet._cells.clear()
        self.safe_append({col_id.col_idx: col_id.value for col_id in self.col_ids}, first_row=True)  # Set col_ids
        for row_idx in self.rows:
            for col_id in self.col_ids:
                copied_cell = copy(self.rows[row_idx][col_id.value])
                self.worksheet._cells[row_idx, col_id.col_idx] = copied_cell
                self.cols[col_id.value].append(copied_cell)
        return affected_row_count

    def traverse(self, fn, where=None, select=None):
        """
        Access cells directly from rows where condition match

        :param fn:
            function(:class:`openpyxl.cell.Cell`) to allow accessing the cell.
            For example, lambda cell: cell.fill = PatternFill(fill_type="solid", fgColor="00FFFF00").
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :param select:
            The columns of the row to update.
            For example, ["id"], where only column under "id" will be accessed
        :type select: list
        :return: Number of rows traversed
        :rtype: int
        """
        if callable(fn) is False:
            raise TypeError("Expected callable for argument fn(cell)")
        row_idxs_where = self.row_idxs_where(where)
        for row_idx in row_idxs_where:
            select = [col_id.value for col_id in self.col_ids] if select is None else select
            for matched_col_id in [col_id for col_id in self.col_ids if col_id.value in select]:
                cell = self.rows[row_idx][matched_col_id.value]
                fn(cell)  # Expect callable to modify cell
                # Update value to worksheet
                self.worksheet._cells[row_idx, matched_col_id.col_idx] = cell
                # No need to update cols as it share same reference with row
        return len(row_idxs_where)

    def format(self, formatter, where=None, select=None):
        """
        Convenience method that built on top of traverse to format cell(s).
        If formatter is given, all other formats will be ignored.

        :param formatter:
            CellFormatter that hold all formats. When this is not None other formats will be ignored.
        :type formatter: CellFormatter
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :param select:
            The columns of the row to update.
            For example, ["id"], where only column under "id" will be formatted
        :type select: list
        :return: Number of rows formatted
        :rtype: int
        """
        if where is None and formatter.is_empty():
            return 0
        return self.traverse(lambda cell: formatter.format(cell), where=where, select=select)

    def __len__(self):
        """
        :return: Length of rows doesn't include header
        """
        return len(self.rows)

    def __getitem__(self, row_idx):
        """
        Get rows with row index

        :param row_idx: Row index or callable
        :return: Rows
        """
        return self.query({DAO.COL_ROW_IDX: row_idx})

    def __setitem__(self, row_idx, value):
        """
        Update if contains row_idx else insert.
        Insert will raise UserWarning when row_idx is callable

        :param row_idx: Row index or callable
        :raise UserWarning: When row_idx is callable and row_idx is not exists
        """
        if row_idx in self:
            self.update(value, {DAO.COL_ROW_IDX: row_idx})
        elif not callable(row_idx):
            self.insert(value)
        else:
            warnings.warn("Insertion with callable is not supported, please use Cellbase/DAO.insert() instead."
                          "Ignore this warning, if you are trying to update rows", UserWarning)

    def __delitem__(self, row_idx):
        """
        Delete with row index

        :param row_idx: Row index or callable
        """
        if row_idx in self:
            self.delete({DAO.COL_ROW_IDX: row_idx})

    def __contains__(self, row_idx):
        """
        Check if row index exists in Celltable

        :param row_idx: Row index or callable
        :return: If row exist
        :rtype: bool
        """
        return len(self.row_and_col_where(where={DAO.COL_ROW_IDX: row_idx})) > 0
