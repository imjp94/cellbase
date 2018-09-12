import os
from abc import ABC, abstractmethod

import pygsheets
from openpyxl import Workbook, load_workbook
from pygsheets import ExportType, SpreadsheetNotFound

from cellbase.formatter import CellFormatter
from cellbase.celltable import Celltable, LocalCelltable, GoogleCelltable


class Cellbase(ABC):
    """
    Cellbase is equivalent to :class:`Workbook` which stores :class:`Celltable`
    """
    DEFAULT_FILENAME = 'cellbase.xlsx'

    def __init__(self):
        self._path = ''
        self._filename = Cellbase.DEFAULT_FILENAME
        self.workbook = None
        self.schemas = {}
        self.celltables = {}

    @property
    def path(self):
        return self._path

    @property
    def filename(self):
        return self._filename

    @property
    def dir(self):
        return os.path.join(self.path, self.filename)

    def load(self, path, filename, raise_err=False):
        """
        Load workbook from given dir

        Notice how this named as "load" instead of "open" as it does not open a connection or stream with the workbook.
        Instead, it simply load the data into memory and any changes will only be saved unless save or save_as
        is called.

        :param path: Path of workbook to load
        :type path: str
        :param filename: Filename of workbook to load
        :type filename: str
        :param raise_err: Whether to raise FileNotFoundError or create new workbook
        :return: self
        :rtype: Cellbase
        """
        self._path = path
        self._filename = filename
        self._on_load(raise_err)
        return self

    def create_if_none(self, worksheet_name):
        """
        Create worksheet and add to cell_tables if there's no such worksheet. It is first called in every data
        accessing methods like query, insert, update, etc.

        :param worksheet_name: Name of worksheet to inspect or create if required
        :type worksheet_name: str
        """
        if worksheet_name not in self.celltables:
            if worksheet_name not in self.schemas:
                raise ValueError(
                    "Trying to create Celltable '%s' without schema " % worksheet_name)
            self.celltables[worksheet_name] = self._on_create(worksheet_name)

    def drop(self, worksheet_name):
        """
        Delete specified worksheet.

        :param worksheet_name: Name of worksheet to delete
        :type worksheet_name: str
        """
        self._on_drop(worksheet_name)
        self.celltables.pop(worksheet_name)

    def save_as(self, path, filename, overwrite=False):
        """
        Save workbook to dir. FileExistsError will be raised if file exists and overwrite is False.

        :param path: Path of workbook to load
        :type path: str
        :param filename: Path to save the workbook
        :type filename: str
        :param overwrite: Whether to overwrite if file exists
        :type overwrite: bool
        :raises FileExitsError: File exists and overwrite is False
        """
        if os.path.exists(filename) and not overwrite:
            raise FileExistsError("%s already exists, set overwrite=True if this is expected.")
        self._on_save(path, filename)

    def save(self):
        """
        Save workbook to the filename specified in open, overwrite if file exist.
        """
        self.save_as(self.path, self.filename, overwrite=True)

    @abstractmethod
    def _on_load(self, raise_err):
        pass

    @abstractmethod
    def _on_create(self, worksheet_name):
        pass

    @abstractmethod
    def _on_drop(self, worksheet):
        pass

    @abstractmethod
    def _on_save(self, path, filename):
        pass

    def register(self, on_create):
        """
        Register format of worksheet to deal with, only required for newly created worksheet
        Example::

        {'TABLE_NAME': ['COL_NAME_1', 'COL_NAME_2']}
        :param on_create: Format of Celltable to deal with
        :type on_create: dict
        :return:
        """
        self.schemas.update(on_create)

    def query(self, worksheet_name, where=None):
        """
        Return data from Celltable with specified worksheet_name, that match the conditions.
        Return all data if no condition given.

        :param worksheet_name: Name of worksheet to query from
        :type worksheet_name: str
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return:
            List of dict that store value corresponding to the column id.
            * row_idx is the default value to return, where it specifies the row index of returned data.
            row_idx is corresponding to the actual row index in spreadsheet, so the minimum index is 2 where 1st row
            is taken by the column ids(header)
            For example, [{"row_idx": 2, "id": 1, "name": "jp1"}, {"row_idx": 3, "id": 2, "name": "jp2"}]
        :rtype: list
        """
        self.create_if_none(worksheet_name)
        return self.celltables[worksheet_name].query(where=where)

    def insert(self, worksheet_name, value_in_dict):
        """
        Insert new row to the worksheet

        :param worksheet_name: Name of worksheet to insert to
        :type worksheet_name: str
        :param value_in_dict:
            Dict that describe the row to insert, where row_idx is not required.
            For example, {"id": 1, "name": "jp1"}
        :type value_in_dict: dict
        :return: row_idx of new row
        :rtype: int
        """
        self.create_if_none(worksheet_name)
        return self.celltables[worksheet_name].insert(value_in_dict)

    def update(self, worksheet_name, value_in_dict, where=None):
        """
        Update row(s) that match the condition.
        If row_idx is given in value_in_dict, wheres & conds will be ignored and only the exact row will be updated.

        :param worksheet_name: Name of the worksheet to update
        :type worksheet_name: str
        :param value_in_dict: Dict that describe the row, where row_idx is optional.
        :type value_in_dict: dict
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Number of rows updated
        :rtype: int
        """
        self.create_if_none(worksheet_name)
        return self.celltables[worksheet_name].update(value_in_dict, where=where)

    def delete(self, worksheet_name, where=None):
        """
        Delete row(s) that match conditions.

        :param worksheet_name: Name of worksheet to delete row(s)
        :type worksheet_name: str
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Number of rows deleted
        :rtype: int
        """
        self.create_if_none(worksheet_name)
        return self.celltables[worksheet_name].delete(where=where)

    def traverse(self, worksheet_name, fn, where=None, select=None):
        """
        Access cells directly from rows where conditions matched.

        :param worksheet_name: Name of worksheet to traverse
        :type worksheet_name: str
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
        self.create_if_none(worksheet_name)
        return self.celltables[worksheet_name].traverse(fn, where=where, select=select)

    def format(self, worksheet_name, formatter, where=None, select=None):
        """
        Convenience method that built on top of traverse to format cell(s).
        If formatter is given, all other formats will be ignored.

        :param worksheet_name: Name of worksheet to format
        :type worksheet_name: str
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :param select:
            The columns of the row to update.
            For example, ["id"], where only column under "id" will be formatted
        :type select: list
        :param formatter:
            CellFormatter or dict that hold all formats.
        :type formatter: CellFormatter or dict
        :return: Number of rows formatted
        :rtype: int
        """
        self.create_if_none(worksheet_name)
        return self.celltables[worksheet_name].format(formatter, where=where, select=select)

    def __len__(self):
        """
        Return numbers of worksheet

        :return: Numbers of worksheet
        """
        return len(self.celltables)

    def __getitem__(self, worksheet_name):
        """
        Get Celltable with worksheet_name

        :param worksheet_name: worksheet_name to find
        :return: Celltable to find
        :rtype: Celltable
        """
        self.create_if_none(worksheet_name)
        return self.celltables[worksheet_name]

    def __setitem__(self, worksheet_name, celltable):
        """
        Celltable must be created by Cellbase

        :raise AssertionError: When attempt to assign
        """
        raise AssertionError("Celltable must be created by Cellbase")

    def __delitem__(self, worksheet_name):
        """
        Drop worksheet

        :param worksheet_name: Name of worksheet to drop
        :type worksheet_name: str
        """
        self.drop(worksheet_name)

    def __contains__(self, worksheet_name):
        """
        Check if worksheet exists

        :param worksheet_name: Name of worksheet to check
        :type worksheet_name: str
        :return: If worksheet exists
        :rtype: bool
        """
        return worksheet_name in self.celltables


class LocalCellbase(Cellbase):

    def __init__(self):
        super().__init__()
        self.workbook = Workbook()

    def _on_load(self, raise_err):
        if os.path.exists(self.dir):
            self.workbook = load_workbook(self.dir)
        elif raise_err:
            raise FileNotFoundError("No workbook found at %s" % self.dir)
        else:
            self.workbook = Workbook()
        for worksheet in self.workbook.worksheets:
            self.celltables[worksheet.title] = LocalCelltable(worksheet)

    def _on_create(self, worksheet_name):
        worksheet = self.workbook.create_sheet(title=worksheet_name, index=0)
        worksheet.append(self.schemas[worksheet_name])
        return LocalCelltable(worksheet)

    def _on_drop(self, worksheet_name):
        worksheet_to_drop = self.celltables[worksheet_name].worksheet
        # Workbook must contain at least 1 visible sheet
        visible_sheets = [worksheet for worksheet in self.workbook.worksheets
                          if worksheet.sheet_state == 'visible']
        if len(visible_sheets) == 1 and visible_sheets[0] is worksheet_to_drop:
            self.workbook.create_sheet()
        self.workbook.remove(worksheet_to_drop)

    def _on_save(self, path, filename):
        self.workbook.save(os.path.join(path, filename))


class GoogleCellbase(Cellbase):
    ATTRIBUTES = ('client_secret', 'service_account_file', 'credentials_directory')

    def __init__(self, export_format=ExportType.CSV, **kwargs):
        super().__init__()
        unexpected_attrs = [attr for attr in kwargs if attr not in GoogleCellbase.ATTRIBUTES]
        print(unexpected_attrs)
        if unexpected_attrs:
            raise AttributeError("Unexpected attribute%s, expecting%s only" %
                                 (unexpected_attrs, GoogleCellbase.ATTRIBUTES))
        else:
            self.__dict__.update(kwargs)
        self.export_format = export_format

    def _on_load(self, raise_err):
        client = pygsheets.authorize(self.client_secret, self.service_account_file, self.credentials_directory)
        try:
            self.workbook = client.open(self._filename)
        except SpreadsheetNotFound:
            if raise_err:
                raise FileNotFoundError("No workbook found at %s" % self.filename)
            else:
                self.workbook = client.create(self._filename, folder=self.path or 'root')
        for worksheet in self.workbook.worksheets():
            self.celltables[worksheet.title] = GoogleCelltable(worksheet, worksheet.title in self.schemas)

    def _on_create(self, worksheet_name):
        worksheet = self.workbook.add_worksheet(title=worksheet_name, index=0)
        worksheet.update_row(1, self.schemas[worksheet_name])
        return GoogleCelltable(worksheet)

    def _on_drop(self, worksheet_name):
        # TODO: Should make sure there is at least 1 visible worksheet else create 1 before delete
        worksheet_to_drop = self.workbook.worksheet_by_title(worksheet_name)
        self.workbook.del_worksheet(worksheet_to_drop)

    def _on_save(self, path, filename):
        self.workbook.export(self.export_format, path, filename)

    def __getattr__(self, attr):
        try:
            return self.__dict__[attr]
        except KeyError:
            if attr not in GoogleCellbase.ATTRIBUTES:
                raise AttributeError("Unexpected attribute %s, expecting%s only" % (attr, GoogleCellbase.ATTRIBUTES))
            else:
                default = None
                if attr == 'client_secret':
                    default = 'client_secret.json'
                    self.__dict__[attr] = default
                else:
                    self.__dict__[attr] = default
                return default
