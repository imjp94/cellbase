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
        self._worksheet = worksheet
        self._rows = []
        self._col_ids = []
        self._size = 0

    @property
    def row_idxs(self):
        return (i for i, _ in enumerate(self._rows, 2))

    @property
    def col_names(self):
        return (self._get_cell_attr(col_id, 'value') for col_id in self._col_ids)

    @property
    def size(self):
        return self._size

    @property
    def last_row_idx(self):
        return self._size + 1

    def query(self, where=None):
        """
        Query data where conditions match

        :return: List of rows
        """
        rows_to_return = []
        for row_idx in self._cross_where(where):
            values = {DAO.COL_ROW_IDX: row_idx}
            for col_name, cell in self._get_row(row_idx).items():
                values[col_name] = self._get_cell_attr(cell, 'value')
            rows_to_return.append(values)
        return rows_to_return

    def insert(self, data):
        """
        Insert new row of data, expect value of row in dict corresponding to col_names

        :return: New row index
        """
        if not isinstance(data, dict):
            raise TypeError("Expecting dict given %s" % type(data))
        new_row_idx = self.last_row_idx + 1
        new_row = self._on_insert(data, new_row_idx)
        self._rows.append({})
        for col_id in self._col_ids:
            new_cell = new_row[self._get_col_idx(col_id) - 1]
            col_name = self._get_col_name(col_id)
            self._set_cell(new_row_idx, col_name, new_cell)
        self._size += 1
        return new_row_idx

    def update(self, data, where=None):
        """
        Update row(s) where conditions match

        :return: Number of rows updated
        """
        if DAO.COL_ROW_IDX in data:
            where = {DAO.COL_ROW_IDX: data[DAO.COL_ROW_IDX]}
        elif not where:
            raise KeyError("row_idx not found, it must be provided if 'where' is omitted")
        return self.traverse(lambda cell: self._set_cell_attr(cell, 'value', data[self._get_col_name(cell)]),
                             where)

    def delete(self, where=None):
        """
        Delete row(s) of data where conditions match

        :return: Number of rows deleted
        """
        row_idxs_to_delete = self._cross_where(where)
        num_rows_deleted = len(row_idxs_to_delete)
        if num_rows_deleted == 0:
            return 0
        self._on_delete(*self._pop_rows(row_idxs_to_delete))
        self._size -= num_rows_deleted
        return num_rows_deleted

    def traverse(self, fn, where=None, select=None):
        """
        Access cells directly from rows where condition match

        Select all column if select omitted

        :return: Number of rows traversed
        """
        if not callable(fn):
            raise TypeError("Expected callable for argument fn(cell)")
        row_idxs_to_traverse = self._cross_where(where)
        num_rows_traversed = len(row_idxs_to_traverse)
        if num_rows_traversed == 0:
            return 0
        select = select or list(self.col_names)
        traversed_cells = []
        for row_idx in row_idxs_to_traverse:
            for col_name in [col_name for col_name in self.col_names if col_name in select]:
                cell = self._get_cell(row_idx, col_name)
                fn(cell)
                traversed_cells.append(cell)
        self._on_traverse(traversed_cells)
        return num_rows_traversed

    def format(self, formatter, where=None, select=None):
        """
        Convenience method that built on top of traverse to format cell(s).

        formatter can be :class:`cellbase.formatter.CellFormatter` or dict

        :return: Number of formatted rows
        """
        if len(formatter) == 0:
            return 0
        formatter = self._formatter_cls()(**formatter) if isinstance(formatter, dict) else formatter
        return self.traverse(lambda cell: formatter.format(cell), where=where, select=select)

    @abstractmethod
    def _on_insert(self, data, new_row_idx):
        pass

    @abstractmethod
    def _on_delete(self, shifted_coords, popped_coords):
        pass

    @abstractmethod
    def _on_traverse(self, cells):
        pass

    @abstractmethod
    def _formatter_cls(self):
        pass

    def _parse(self, first_row, content_row, on_parse_cell=None):
        self._col_ids = [col_id for col_id in first_row if self._get_cell_attr(col_id, 'value')]
        for row in content_row:
            row_idx = self._get_cell_attr(row[0], 'row')
            for col_id in self._col_ids:
                cell = row[self._get_col_idx(col_id) - 1]  # -1 as row is list(0 indexed)
                if self._get_cell_attr(cell, 'value'):
                    if self._size < row_idx:
                        self._size = row_idx - 1
                if on_parse_cell:
                    on_parse_cell(cell)
                if row_idx not in self.row_idxs:
                    self._rows.append({})
                self._set_cell(row_idx, self._get_col_name(col_id), cell)
        # Opt out empty rows after size
        self._rows = self._rows[:self.size]

    def _pop_rows(self, row_idxs):
        row_idxs_affected = list(range(row_idxs[0], self.last_row_idx + 1))
        row_idxs_remain = [row_idx for row_idx in row_idxs_affected if row_idx not in row_idxs]
        shifted_coords = []
        popped_coords = []
        for row_idx in row_idxs_affected:
            if row_idxs_remain:
                row_idx_remain = row_idxs_remain.pop(0)
                # Shift cell to overwrite "deleted" cell
                for col_name in self.col_names:
                    cell = self._get_cell(row_idx_remain, col_name)
                    self._set_cell_attr(cell, 'row', row_idx)
                    self._set_cell(row_idx, col_name, cell)
                    shifted_coords.append((row_idx, self._get_col_idx(col_name)))
            else:
                # Pop cell that already shifted and left to be empty
                self._rows.pop()
                for col_name in self.col_names:
                    popped_coords.append((row_idx, self._get_col_idx(col_name)))
        return shifted_coords, popped_coords

    def _vertical_where(self, where=None):
        """ Find the row indexes where any of the conditions match """
        if where is None:
            return list(self.row_idxs)
        row_idxs = []
        if DAO.COL_ROW_IDX in where:
            cond = where[DAO.COL_ROW_IDX]
            if callable(where[DAO.COL_ROW_IDX]):
                for row_idx in self.row_idxs:
                    if cond(row_idx):
                        row_idxs.append(row_idx)
            else:
                row_idx = int(cond)
                if row_idx in self.row_idxs:
                    row_idxs.append(row_idx)
        for col_name, cond in where.items():
            if col_name == DAO.COL_ROW_IDX:
                continue
            for row_idx in self.row_idxs:
                cell = self._get_cell(row_idx, col_name)
                value = self._get_cell_attr(cell, 'value')
                if row_idx not in row_idxs and cond(value) if callable(cond) else value == cond:
                    row_idxs.append(row_idx)
        return row_idxs

    def _horizontal_where(self, row_idx, where=None):
        """ Find the column names where conditions match from a specific row """
        if where is None:
            return list(self.col_names)
        col_names = []
        for col_name, cond in where.items():
            if col_name == DAO.COL_ROW_IDX:
                if cond(row_idx) if callable(cond) else row_idx == int(cond):
                    col_names.append(col_name)
                continue
            cell = self._get_cell(row_idx, col_name)
            value = self._get_cell_attr(cell, 'value')
            if cond(value) if callable(cond) else value == cond:
                col_names.append(col_name)
        return col_names

    def _cross_where(self, where=None):
        """ Find row indexes where all conditions match by combining row_idx_where and col_names_where """
        row_idxs_where = self._vertical_where(where)
        if where is None:
            return row_idxs_where
        row_idxs = []
        for row_idx in row_idxs_where:
            if len(self._horizontal_where(row_idx, where)) == len(where):
                row_idxs.append(row_idx)
        return row_idxs

    def _get_row(self, idx):
        return self._rows[idx - 2]

    def _set_row(self, idx, row):
        self._rows[idx - 2] = row

    def _get_cell(self, idx, name):
        return self._get_row(idx)[name]

    def _set_cell(self, idx, name, cell):
        self._get_row(idx)[name] = cell

    def _cell_attrs(self):
        return Celltable.DEFAULT_CELL_ATTRS

    def _get_cell_attr(self, cell, attr):
        return getattr(cell, self._cell_attrs()[attr])

    def _set_cell_attr(self, cell, attr, value):
        setattr(cell, self._cell_attrs()[attr], value)

    def _get_col_idx(self, cell_or_name):
        if not isinstance(cell_or_name, str):
            name = self._get_cell_attr(cell_or_name, 'value')
        else:
            name = cell_or_name
        for i, col_name in enumerate(self.col_names):
            if col_name == name:
                return self._get_cell_attr(self._col_ids[i], 'col')
        raise KeyError("Failed to get column index, no column name %s" % name)

    def _get_col_name(self, cell_or_idx):
        """ Get column name from cell or column index """
        if not isinstance(cell_or_idx, int):
            col_idx = self._get_cell_attr(cell_or_idx, 'col')
        else:
            col_idx = cell_or_idx
        return self._get_cell_attr(self._col_ids[col_idx - 1], 'value')

    def __len__(self):
        """ Length of rows doesn't include header """
        return self._size

    def __getitem__(self, row_idx):
        """ Get rows with row index or callable """
        return self.query({DAO.COL_ROW_IDX: row_idx})

    def __setitem__(self, row_idx, data):
        """ Update if contains row_idx else insert. Insert will raise UserWarning when row_idx is callable """
        if row_idx in self:
            self.update(data, {DAO.COL_ROW_IDX: row_idx})
        elif not callable(row_idx):
            self.insert(data)
        else:
            warnings.warn("Insertion with callable is not supported, please use Cellbase/DAO.insert() instead."
                          "Ignore this warning, if you are trying to update rows", UserWarning)

    def __delitem__(self, row_idx):
        """ Delete with row index """
        if row_idx in self:
            self.delete({DAO.COL_ROW_IDX: row_idx})

    def __contains__(self, row_idx):
        """ Check if row index exists in Celltable """
        return len(self._cross_where(where={DAO.COL_ROW_IDX: row_idx})) > 0


class LocalCelltable(Celltable):
    LOCAL_CELL_ATTRS = {'value': 'value', 'row': 'row', 'col': 'col_idx'}

    def __init__(self, worksheet):
        super().__init__(worksheet)
        self._parse(worksheet[1], worksheet.iter_rows(min_row=2))

    def _on_insert(self, data, new_row_idx):
        # Make sure openpyxl actualy append at last row
        orig_current_row = self._worksheet._current_row
        self._worksheet._current_row = self._size + 1  # row_idx = worksheet._current_row + 1, see worksheet.append
        self._worksheet.append({col_id.col_idx: data[col_id.value] for col_id in self._col_ids})
        self._worksheet._current_row = orig_current_row
        return list(self._worksheet.rows)[new_row_idx - 1]

    def _on_delete(self, shifted_coords, popped_coords):
        for row_idx, col_idx in shifted_coords:
            self._worksheet._cells[(row_idx, col_idx)] = self._get_cell(row_idx, self._get_col_name(col_idx))
        for row_idx, col_idx in popped_coords:
            del self._worksheet._cells[(row_idx, col_idx)]

    def _on_traverse(self, cells):
        for cell in cells:
            self._worksheet._cells[cell.row, cell.col_idx] = cell

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

    @property
    def has_fetched(self):
        return self._has_fetched

    def fetch(self):
        self._on_fetch()
        self._has_fetched = True

    @abstractmethod
    def _on_fetch(self):
        pass

    def _fetch_if_havent(self):
        if not self._has_fetched:
            self.fetch()

    def query(self, where=None):
        self._fetch_if_havent()
        return super().query(where)

    def insert(self, data):
        self._fetch_if_havent()
        return super().insert(data)

    def delete(self, where=None):
        self._fetch_if_havent()
        return super().delete(where)

    def traverse(self, fn, where=None, select=None):
        self._fetch_if_havent()
        return super().traverse(fn, where, select)


class GoogleCelltable(RemoteCelltable):
    def __init__(self, worksheet, fetch=False):
        super().__init__(worksheet, fetch)

    def _on_insert(self, data, new_row_idx):
        values = self._data_to_row_value(data)
        if self._worksheet.rows - 1 > self._size:
            self._worksheet.update_row(new_row_idx, values)
        else:
            # new_row_idx - 1 as it is inserted below the target row index
            self._worksheet.insert_rows(new_row_idx - 1, values=values)
        return self._worksheet.get_row(new_row_idx, 'cell')

    def _on_delete(self, shifted_coords, popped_coords):
        cells = []
        for row_idx, col_idx in shifted_coords:
            cells.append(self._get_cell(row_idx, self._get_col_name(col_idx)))
        for row_idx, col_idx in popped_coords:
            cells.append(Cell((row_idx, col_idx)))
        self._worksheet.update_cells(cells)

    def _on_traverse(self, cells):
        self._worksheet.update_cells(cells)

    def _on_fetch(self):
        all_rows = self._worksheet.get_all_values('cell')
        self._parse(all_rows[0], all_rows[1:], lambda cell: cell.unlink())

    def _formatter_cls(self):
        return GoogleCellFormatter

    def _data_to_row_value(self, data):
        """ Convert dictionary to list according the sequence of col_ids """
        col_seqs = [col_id.col for col_id in self._col_ids]
        values = []
        for col_idx in range(1, self._col_ids[-1].col + 1):
            if col_idx in col_seqs:
                values.append(data[self._col_ids[col_seqs.index(col_idx)].value])
            else:
                values.append('')
        return values
