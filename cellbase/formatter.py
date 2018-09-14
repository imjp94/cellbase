from abc import ABC, abstractmethod


class CellFormatter(ABC):
    """ Helper class to store all the cell formatting and helps to apply them to cell object """

    def __init__(self, **kwargs):
        """ Keyword arguments are the attributes of targeted cell object """
        unexpected_attrs = [attr for attr in kwargs if attr not in self.attrs()]
        if unexpected_attrs:
            raise AttributeError("Unexpected attribute%s, expecting%s only" % (unexpected_attrs, self.attrs()))
        else:
            self.__dict__.update(kwargs)

    @abstractmethod
    def attrs(self):
        pass

    def format(self, cell):
        """ Format cell object with all stored formatting """
        attrs = self.non_nones()
        if len(attrs) == 0:
            return
        for attr in attrs:
            self.on_format(cell, attr)

    def on_format(self, cell, attr):
        """ When internal formatting performed, call format() instead to format cell """
        setattr(cell, attr, self.__getattr__(attr))

    def non_nones(self):
        """ Return list of formatting that are not None """
        return [attr for attr in self.attrs() if self.__getattr__(attr)]

    def __len__(self):
        """ Length of not None formatting """
        return len(self.non_nones())

    def __setattr__(self, attr, value):
        if attr not in self.attrs():
            raise AttributeError("Unexpected attribute %s, expecting%s only" % (attr, self.attrs()))
        else:
            self.__dict__[attr] = value

    def __getattr__(self, attr):
        if attr not in self.attrs():
            raise AttributeError("Unexpected attribute %s, expecting%s only" % (attr, self.attrs()))
        else:
            try:
                return self.__dict__[attr]
            except KeyError:
                self.__dict__[attr] = None
                return None


class LocalCellFormatter(CellFormatter):
    """ CellFormatter that format :class:`openpyxl.cell.Cell` """
    ATTRIBUTES = ('font', 'fill', 'border', 'number_format', 'protection', 'alignment', 'style')

    def attrs(self):
        return LocalCellFormatter.ATTRIBUTES


class GoogleCellFormatter(CellFormatter):
    """ CellFormatter that format :class:`pygsheets.Cell` """
    ATTRIBUTES = ('color', 'horizontal_alignment', 'vertical_alignment', 'wrap_strategy', 'note', 'set_text_format',
                  'set_number_format', 'set_text_rotation')
    METHODS = ('set_text_format', 'set_number_format', 'set_text_rotation')

    def on_format(self, cell, attr):
        if attr in GoogleCellFormatter.METHODS:
            self.format_method(cell, attr)
        else:
            super().on_format(cell, attr)

    def format_method(self, cell, attr):
        """ Handle formatting of cell that must call method """
        attr_pair = self.__getattr__(attr)
        getattr(cell, attr)(attr_pair[0], attr_pair[1])

    def attrs(self):
        return GoogleCellFormatter.ATTRIBUTES
