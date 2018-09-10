import warnings
from abc import ABC, abstractmethod

from pygsheets import Cell

from cellbase.formatter import LocalCellFormatter, GoogleCellFormatter
from cellbase.helper import DAO


class Celltable(ABC):
    """
    Celltable is equivalent to Worksheet from Workbook
    """
    DEFAULT_CELL_ATTRS = {'value': 'value', 'row': 'row', 'col': 'col'}

    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.rows = {}
        self.cols = {}
        self.col_ids = []
        self._max_row = 0

    def query(self, where=None):
        """
        Query data where conditions match

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: List of rows
        :rtype: list
        """
        rows_to_return = []
        for row_idx in self._row_and_col_where(where):
            values = {DAO.COL_ROW_IDX: row_idx}
            for key, cell in self.rows[row_idx].items():
                values[key] = self._get_cell(cell, 'value')
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
        if not isinstance(value_in_dict, dict):
            raise TypeError("Expecting dict given %s" % type(value_in_dict))
        new_row_idx = self._max_row + 2  # +2 = +1 as _max_row is size and +1 for new row
        new_row = self._on_insert(value_in_dict, new_row_idx)
        self.rows[new_row_idx] = {}
        for col_id in self.col_ids:
            new_cell = new_row[self._get_cell(col_id, 'col') - 1]
            col_id_value = self._get_cell(col_id, 'value')
            self.rows[new_row_idx][col_id_value] = new_cell
            self.cols[col_id_value].append(new_cell)
        self._max_row += 1
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
            try:
                row = self.rows[value_in_dict[DAO.COL_ROW_IDX]]
            except KeyError:
                raise ValueError("row_idx not found, it must be provided if 'where' is omitted")
            for cell in list(row.values())[1:]:
                self._set_cell(cell, 'value', value_in_dict[self._col_idx_to_col_id(self._get_cell(cell, 'col')).value])
            return 1
        return self.traverse(
            lambda cell: self._set_cell(
                cell, 'value', value_in_dict[self._col_idx_to_col_id(self._get_cell(cell, 'col')).value]), where
        )

    def delete(self, where=None):
        """
        Delete row(s) of data where conditions match

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Number of rows deleted
        :rtype:int
        """
        row_idxs_to_delete = self._row_and_col_where(where)
        num_rows_deleted = len(row_idxs_to_delete)
        if num_rows_deleted == 0:
            return 0
        self._on_delete(*self._pop_rows(row_idxs_to_delete))
        self._max_row -= num_rows_deleted
        return num_rows_deleted

    def traverse(self, fn, where=None, select=None):
        """
        Access cells directly from rows where condition match

        :param fn:
            function(cell) to allow accessing the cell.
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
        if not callable(fn):
            raise TypeError("Expected callable for argument fn(cell)")
        row_idxs_to_traverse = self._row_idxs_where(where)
        num_rows_traversed = len(row_idxs_to_traverse)
        if num_rows_traversed == 0:
            return 0
        select = [self._get_cell(col_id, 'value') for col_id in self.col_ids] if select is None else select
        traversed_cells = []
        for row_idx in row_idxs_to_traverse:
            for matched_col_id in [col_id for col_id in self.col_ids if col_id.value in select]:
                cell = self.rows[row_idx][matched_col_id.value]
                fn(cell)
                traversed_cells.append(cell)
        self._on_traverse(traversed_cells)
        return num_rows_traversed

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
        if len(formatter) == 0:
            return 0
        formatter = self._formatter_cls()(**formatter) if isinstance(formatter, dict) else formatter
        return self.traverse(lambda cell: formatter.format(cell), where=where, select=select)

    @abstractmethod
    def _on_insert(self, value_in_dict, new_row_idx):
        pass

    @abstractmethod
    def _on_delete(self, shifted_cells, popped_cells):
        pass

    @abstractmethod
    def _on_traverse(self, cells):
        pass

    @abstractmethod
    def _formatter_cls(self):
        pass

    def _parse(self, first_row, content_row, on_parse_cell=None):
        self.col_ids = [col_id for col_id in first_row if self._get_cell(col_id, 'value')]  # Ignore cols with no value
        self.cols = {self._get_cell(col_id, 'value'): [] for col_id in self.col_ids}
        for row in content_row:
            row_idx = self._get_cell(row[0], 'row')
            for col_id in self.col_ids:
                cell = row[self._get_cell(col_id, 'col') - 1]  # -1 as row is list(0 indexed)
                if self._get_cell(cell, 'value'):
                    if self._max_row < row_idx:
                        self._max_row = row_idx - 1
                    if on_parse_cell:
                        on_parse_cell(cell)
                    col_id_value = self._get_cell(col_id, 'value')
                    self.cols[col_id_value].append(cell)
                    if row_idx not in self.rows:
                        self.rows[row_idx] = {}
                    self.rows[row_idx][col_id_value] = cell

    def _cell_attrs(self):
        return Celltable.DEFAULT_CELL_ATTRS

    def _get_cell(self, cell, attr):
        return getattr(cell, self._cell_attrs()[attr])

    def _set_cell(self, cell, attr, value):
        setattr(cell, self._cell_attrs()[attr], value)

    def _pop_rows(self, row_idxs):
        # +2 = +1 as max_row is size and +1 for range exclusive
        row_idxs_affected = list(range(row_idxs[0], self._max_row + 2))
        row_idxs_remain = [row_idx for row_idx in row_idxs_affected if row_idx not in row_idxs]
        shifted_cells = []
        popped_cells = []
        for row_idx in row_idxs_affected:
            if row_idxs_remain:
                row_idx_remain = row_idxs_remain.pop(0)
                # Shift cell to overwrite "deleted" cell
                for col_id in self.col_ids:
                    col_id_value = self._get_cell(col_id, 'value')
                    cell = self.rows[row_idx_remain][col_id_value]
                    self._set_cell(cell, 'row', row_idx)
                    self.rows[row_idx][col_id_value] = cell
                    self.cols[col_id_value][row_idx - 2] = cell
                    shifted_cells.append(cell)
            else:
                # Pop cell that already shifted and left to be empty
                del self.rows[row_idx]
                for col_id in self.col_ids:
                    cell = self.cols[self._get_cell(col_id, 'value')].pop()
                    popped_cells.append(cell)
        return shifted_cells, popped_cells

    def _col_idx_to_col_id(self, col_idx):
        """
        Get column id cell with column index

        :param col_idx: Column index
        :type col_idx: int
        :return: Column id cell
        """
        return self.col_ids[col_idx - 1]

    def _row_idxs_where(self, where=None):
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
                row = self._get_cell(cell, 'row')
                value = self._get_cell(cell, 'value')
                if row not in row_idxs and cond(value) if callable(cond) else value == cond:
                    row_idxs.append(row)

        return row_idxs

    def _col_names_where(self, row_idx, where=None):
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
            value = self._get_cell(cell, 'value')
            if cond(value) if callable(cond) else value == cond:
                col_names.append(col_name)
        return col_names

    def _row_and_col_where(self, where=None):
        """
        Find row indexes where all conditions match by combining row_idx_where and col_names_where

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Row indexes where all conditions match
        :rtype: list
        """
        row_idxs_where = self._row_idxs_where(where)
        if where is None:
            return row_idxs_where
        row_idxs = []
        for row_idx in row_idxs_where:
            if len(self._col_names_where(row_idx, where)) == len(where):
                row_idxs.append(row_idx)
        return row_idxs

    def __len__(self):
        """
        :return: Length of rows doesn't include header
        """
        return self._max_row

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
        return len(self._row_and_col_where(where={DAO.COL_ROW_IDX: row_idx})) > 0


class LocalCelltable(Celltable):
    LOCAL_CELL_ATTRS = {'value': 'value', 'row': 'row', 'col': 'col_idx'}

    def __init__(self, worksheet):
        super().__init__(worksheet)
        self._parse(worksheet[1], worksheet.iter_rows(min_row=2))

    def _on_insert(self, value_in_dict, new_row_idx):
        # Make sure openpyxl actualy append at last row
        orig_current_row = self.worksheet._current_row
        self.worksheet._current_row = self._max_row + 1  # row_idx = worksheet._current_row + 1, see worksheet.append
        self.worksheet.append({col_id.col_idx: value_in_dict[col_id.value] for col_id in self.col_ids})
        self.worksheet._current_row = orig_current_row
        return list(self.worksheet.rows)[new_row_idx - 1]

    def _on_delete(self, shifted_cells, popped_cells):
        for cell in shifted_cells:
            self.worksheet._cells[(cell.row, cell.col_idx)] = self.rows[cell.row][self._col_idx_to_col_id(cell.col_idx).value]
        for cell in popped_cells:
            del self.worksheet._cells[(cell.row, cell.col_idx)]

    def _on_traverse(self, cells):
        for cell in cells:
            self.worksheet._cells[cell.row, cell.col_idx] = cell

    def _formatter_cls(self):
        return LocalCellFormatter

    def _cell_attrs(self):
        return LocalCelltable.LOCAL_CELL_ATTRS


class RemoteCelltable(Celltable):
    def __init__(self, worksheet, fetch=False):
        super().__init__(worksheet)
        self._has_fetched = False
        if fetch:
            self.fetch()

    def fetch(self):
        self._on_fetch()
        self._has_fetched = True

    @abstractmethod
    def _on_fetch(self):
        pass

    def query(self, where=None):
        if not self._has_fetched:
            self.fetch()
        return super().query(where)

    def insert(self, value_in_dict):
        if not self._has_fetched:
            self.fetch()
        return super().insert(value_in_dict)

    def delete(self, where=None):
        if not self._has_fetched:
            self.fetch()
        return super().delete(where)

    def traverse(self, fn, where=None, select=None):
        if not self._has_fetched:
            self.fetch()
        return super().traverse(fn, where, select)


class GoogleCelltable(RemoteCelltable):
    def __init__(self, worksheet, fetch=False):
        super().__init__(worksheet, fetch)

    def _on_insert(self, value_in_dict, new_row_idx):
        values = self._value_in_dict_to_row_value(value_in_dict)
        if self.worksheet.rows - 1 > self._max_row:
            self.worksheet.update_row(new_row_idx, values)
        else:
            # new_row_idx - 1 as it is inserted below the target row index
            self.worksheet.insert_rows(new_row_idx - 1, values=values)
        return self.worksheet.get_row(new_row_idx, 'cell')

    def _on_delete(self, shifted_cells, popped_cells):
        new_cells = []
        for cell in popped_cells:
            new_cells.append(Cell((cell.row, cell.col)))
        self.worksheet.update_cells(shifted_cells + new_cells)

    def _on_traverse(self, cells):
        self.worksheet.update_cells(cells)

    def _on_fetch(self):
        all_rows = self.worksheet.get_all_values('cell')
        self._parse(all_rows[0], all_rows[1:], lambda cell: cell.unlink())

    def _formatter_cls(self):
        return GoogleCellFormatter

    def _value_in_dict_to_row_value(self, value_in_dict):
        """ Convert dictionary to list according the sequence of col_ids """
        col_seqs = [col_id.col for col_id in self.col_ids]
        values = []
        for col_idx in range(1, self.col_ids[-1].col + 1):
            if col_idx in col_seqs:
                values.append(value_in_dict[self.col_ids[col_seqs.index(col_idx)].value])
            else:
                values.append('')
        return values
