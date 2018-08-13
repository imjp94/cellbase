from abc import abstractmethod, ABC


class DAO:
    """
    Data-Access-Object acts as an abstraction layer to interact with :class:`Cellbase`
    """
    COL_ROW_IDX = "row_idx"

    def __init__(self, cellbase):
        self.cellbase = cellbase

    @abstractmethod
    def worksheet_name(self):
        """
        Name of worksheet that data stored with
        :return: Name of worksheet
        :rtype: str
        """
        pass

    @abstractmethod
    def new_entity(self):
        """
        Return new instance of Entity
        :return: New instance of Entity
        :rtype: Entity
        """
        pass

    def query(self, where=None):
        """
        Return data from Cellbase that match conditions, return all if no condition given.

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
        return [self.new_entity().from_dict(value) for value in
                self.cellbase.query(self.worksheet_name(), where=where)]

    def insert(self, entity):
        """
        Insert new row of data with Entity object, after insertion, entity.row_idx will be updated as well.

        :param entity: Entity object to insert
        :type entity: Entity
        :return: Given Entity object
        """
        entity.row_idx = self.cellbase.insert(self.worksheet_name(), entity.to_dict())
        return entity

    def update(self, entity, where=None):
        """
        Update row(s) of data where conditions match with Entity object

        :param entity: Entity object to update to
        :type entity: Entity
        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Number of rows updated
        :rtype: int
        """
        return self.cellbase.update(self.worksheet_name(), entity.to_dict(), where=where)

    def delete(self, where=None):
        """
        Delete row(s) of data where conditions match

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :return: Number of rows deleted
        :rtype: int
        """
        return self.cellbase.delete(self.worksheet_name(), where=where)

    def traverse(self, fn, where=None, select=None):
        """
        Access cells directly from rows where conditions matched.

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
        return self.cellbase.traverse(self.worksheet_name(), fn, where=where, select=select)

    def format(self, where=None, select=None,
               formatter=None,
               font=None, fill=None, border=None,
               number_format=None, protection=None,
               alignment=None, style=None
               ):
        """
        Convenience method that built on top of traverse to format cell(s).
        If formatter is given, all other formats will be ignored.

        :param where: dict of columns id to inspect. For example, {'id': 1, 'name': 'jp'}.
        :type where: dict
        :param select:
            The columns of the row to update.
            For example, ["id"], where only column under "id" will be formatted
        :type select: list
        :param formatter:
            CellFormatter that hold all formats. When this is not None other formats will be ignored.
        :type formatter: CellFormatter
        :param font: Font of cell
        :type font: openpyxl.styles.Font
        :param fill: Fill cell with color
        :type fill: openpyxl.styles.Fill
        :param border: Border of cell
        :type border: openpyxl.styles.Border
        :param number_format: Number format of cell
        :type number_format: str
        :param protection: Protection of cell
        :type protection: openpyxl.styles.Protection
        :param alignment: Alignment of cell
        :type alignment: openpyxl.styles.Alignment
        :param style: Named style
        :type style: str
        :return: Number of rows formatted
        :rtype: int
        """
        return self.cellbase.format(
            self.worksheet_name(), where=where, select=select,
            formatter=formatter,
            font=font, fill=fill, border=border,
            number_format=number_format, protection=protection,
            alignment=alignment, style=style
        )

    def drop(self):
        """
        Delete worksheet that specified in worksheet_name
        """
        return self.cellbase.drop(self.worksheet_name())

    def get_celltable(self):
        return self.cellbase[self.worksheet_name()]

    celltable = property(fget=get_celltable)

    def __len__(self):
        """
        :return: Length of rows doesn't include header
        """
        return len(self.celltable)

    def __getitem__(self, row_idx):
        """
        Automatically return list of entities when row_idx is callable, else
        return single entity object or None

        :param row_idx: Can be row index or callable
        :return:
            List of entities when row_idx is callable,
            else single entity object or None
        """
        result = [self.new_entity().from_dict(value) for value in self.celltable[row_idx]]
        return result if callable(row_idx) else result[0] if result else None

    def __setitem__(self, row_idx, entity):
        """
        Update if contains row_idx else insert.
        Insert will raise UserWarning when row_idx is callable

        :param row_idx: Row index or callable
        :param entity: Entity to set
        :raise UserWarning: When row_idx is callable and row_idx is not exists
        :type entity: Entity
        """
        self.celltable[row_idx] = entity.to_dict()

    def __delitem__(self, row_idx):
        """
        Delete with row index

        :param row_idx: Row index or callable
        """
        del self.celltable[row_idx]

    def __contains__(self, row_idx):
        """
        Check if row index exists in Celltable

        :param row_idx: Row index or callable
        :return: If row exists
        :rtype: bool
        """
        return self.celltable[row_idx]


class Entity(ABC):
    """
    Associate with :class:`DAO` to convert data to desired type
    """
    def __init__(self):
        """
        Call super() to declare row_idx
        """
        self.row_idx = None

    @abstractmethod
    def from_dict(self, values):
        """
        Parse data from the dict returned by :class:`Cellbase`.

        .. note:: Call super() to handle row_idx

        :param values: Dict returned by :class:`Cellbase`
        :type values: dict
        """
        self.row_idx = values[DAO.COL_ROW_IDX]
        return self

    @abstractmethod
    def to_dict(self):
        """
        Convert this entity to dict that :class:`Cellbase` asked for

        .. note:: Call super() to handle row_idx

        :return: Dict representation of this entity
        :rtype: dict
        """
        values = {}
        if self.row_idx is not None:
            values[DAO.COL_ROW_IDX] = self.row_idx
        return values


class CellFormatter:
    """
    Helper class that store all the formats for cell
    """
    def __init__(
            self,
            font=None, fill=None, border=None,
            number_format=None, protection=None,
            alignment=None, style=None
    ):
        self.font = font
        self.fill = fill
        self.border = border
        self.number_format = number_format
        self.protection = protection
        self.alignment = alignment
        self.style = style

    def format(self, cell):
        """
        Format cell is any formats is not None

        :param cell: Cell to format
        """
        if self.is_empty():
            return
        if self.font is not None:
            cell.font = self.font
        if self.fill is not None:
            cell.fill = self.fill
        if self.border is not None:
            cell.border = self.border
        if self.number_format is not None:
            cell.number_format = self.number_format
        if self.protection is not None:
            cell.protection = self.protection
        if self.alignment is not None:
            cell.alignment = self.alignment
        if self.style is not None:
            cell.style = self.style

    def is_empty(self):
        return all([self.font, self.fill, self.border, self.number_format, self.protection, self.alignment, self.style])
