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
        Return data from Celltable with specified worksheet, that match the conditions.

        Return all data if where omitted.

        Example::

            where = {'id': 1, 'name': 'jp'}

        :return:
            List of dict that store value corresponding to the row_idx & column id.
            For example, [{"row_idx": 2, "id": 1, "name": "jp1"}, {"row_idx": 3, "id": 2, "name": "jp2"}]
        """
        return [self.new_entity().from_dict(value) for value in
                self.cellbase.query(self.worksheet_name(), where=where)]

    def insert(self, entity):
        """
        Insert new row to the worksheet

        Example::

            value_in_dict = {"id": 1, "name": "jp1"}
        :return: row_idx of new row
        """
        entity.row_idx = self.cellbase.insert(self.worksheet_name(), entity.to_dict())
        return entity

    def update(self, entity, where=None):
        """
        Update row(s) that match the condition.

        If row_idx is given in value_in_dict, where will be ignored and only the exact row will be updated.

        Example::

            where = {'id': 1, 'name': 'jp'}.
        :return: Number of updated rows
        """
        return self.cellbase.update(self.worksheet_name(), entity.to_dict(), where=where)

    def delete(self, where=None):
        """
        Delete row(s) that match conditions.

        Delete all if where omitted.

        Example::

            where = {'id': 1, 'name': 'jp'}.
        :return: Number of deleted rows
        """
        return self.cellbase.delete(self.worksheet_name(), where=where)

    def traverse(self, fn, where=None, select=None):
        """
        Access cells directly from rows where conditions matched.

        Select all column if select omitted

        Example::

            fn = lambda cell: cell.fill = PatternFill(fill_type="solid", fgColor="00FFFF00").

            select = ['id']  # only column under "id" will be accessed
        :return: Number of traversed rows
        """
        return self.cellbase.traverse(self.worksheet_name(), fn, where=where, select=select)

    def format(self, formatter, where=None, select=None):
        """
        Convenience method that built on top of traverse to format cell(s).

        formatter can be :class:`cellbase.formatter.CellFormatter` or dict
        :return: Number of formatted rows
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
