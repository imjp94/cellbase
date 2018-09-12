from abc import abstractmethod, ABC


class DAO:
    """
    Data-Access-Object acts as an abstraction layer to interact with :class:`Cellbase`
    """
    COL_ROW_IDX = "row_idx"

    def __init__(self, cellbase):
        self.cellbase = cellbase

    @property
    def celltable(self):
        return self.cellbase[self.worksheet_name()]

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

    def format(self, formatter, where=None, select=None):
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
            CellFormatter or dict that hold all formats.
        :type formatter: CellFormatter or dict
        :return: Number of rows formatted
        :rtype: int
        """
        return self.cellbase.format(self.worksheet_name(), formatter, where, select)

    def drop(self):
        """ Delete worksheet that specified in worksheet_name """
        return self.cellbase.drop(self.worksheet_name())

    def __len__(self):
        """ Length of rows doesn't include header """
        return len(self.celltable)

    def __getitem__(self, row_idx):
        """ Return list of entities when row_idx is callable, else return single entity object or None """
        result = [self.new_entity().from_dict(value) for value in self.celltable[row_idx]]
        return result if callable(row_idx) else result[0] if result else None

    def __setitem__(self, row_idx, entity):
        """ Update if contains row_idx else insert. Insert will raise UserWarning when row_idx is callable """
        self.celltable[row_idx] = entity.to_dict()

    def __delitem__(self, row_idx_or_callable):
        """ Delete with row index or callable"""
        del self.celltable[row_idx_or_callable]

    def __contains__(self, row_idx):
        """ Check if row index exists in Celltable """
        return self.celltable[row_idx]


class Entity(ABC):
    """
    Associate with :class:`DAO` to convert data to desired type
    """
    def __init__(self):
        """ Call super() to declare row_idx """
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
