[![Build Status](https://travis-ci.com/imjp94/cellbase.svg?branch=dev)](https://travis-ci.com/imjp94/cellbase)
[![Build Status](https://travis-ci.com/imjp94/cellbase.svg?branch=master)](https://travis-ci.com/imjp94/cellbase)

# Cellbase v0.1.2

Abstraction layer for accessing spreadsheet as database, built on top
of [openpyxl](https://openpyxl.readthedocs.io/en/stable/).

## Usage

Read, write or edit spreadsheet in database like environment, for example:

```python
cellbase = Cellbase().load('simple.xlsx')
dao = SimpleDAO(cellbase)  # Object inherits from DAO
entity = Simple(id=1, name='jp')  # Object inherits from Entity
# Basic database operations
dao.insert(entity)
dao.query({'row_idx': entity.row_idx})
entity.name = 'imjp'
dao.update(entity)
dao.delete({'row_idx': entity.row_idx})
# Format cells' font, fill, border, etc...
dao.format({'row_idx': entity.row_idx},
    fill=PatternFill(fill_type="solid", fgColor="00FFFF00"))
# Access openpyxl.cell.Cell directly
dao.traverse(lambda cell: do_something(cell),
    {'row_idx': entity.row_idx}, select=['id'])
cellbase.save()
```

## Installing

Install from pypi:

```console
pip install cellbase
```

## For Your Information

There are some rules/concepts being followed by Cellbase,
not necessary to know but it is nice to be awared of them.

- Cellbase = Workbook = Database
- Celltable = Worksheet = Table
- DAO is the helper to access data from Cellbase
- Entity is resposible to convert data to/from dict
- Cellbase named 'load' for reading file instead of 'open' as currently it
    does not open connection/stream to file, which means any changes made
    are not saved or updated until save/save_as is called
- Implemetation of DAO & Entity are optional
- 'where' argument in most methods expect dict in format as below:
    ```python
    where = {'col_name_1', value_1, 'col_name_2': value_2}
    ```
- 'select' argument in traverse & format expect list in format as below:
    ```python
    select = ['col_name_1', 'col_name_2']
    ```
- 'row_idx' is the actual row index in spreadsheet
- 'row_idx' starts from 2 as 1st row is taken by header, which means:
    ```python
    dao.query({'row_idx', 1})  # Will raise KeyError
    ```
- Cellbase doesn't expect input values(dict) consist of 'row_idx' but values
    returned by query() will definitely consist 'row_idx'
- Cellbase expect variable names declared in first row.

    Empty variable will caused whole column to be ignored(column 3).

    It doesn't really matter for rows, empty row as row 3 is still a valid
    row.

    | var_1   | var_2   | (empty) | var_3   |
    |---------|---------|---------|---------|
    | data    | data    | data    | data    |
    | (empty) | (empty) | (empty) | (empty) |
    | data    | data    | (empty) |   data  |

## Getting Started

Cellbase is made to be easily picked up, you may start right away in python
console or implement DAO & Entity to simplify the codes in your scripts.

```python
from cellbase import Cellbase

# Without specifying filename, it will save as 'cellbase.xlsx' by default
cellbase = CellBase()  
# Register the format of worksheet to deal with(only for new worksheet)
# 'Simple' is the worksheet name, while 'id' and 'name' are column names
cellbase.register({'Simple': ['id', 'name']})
```

- Without DAO & Entity:
    ```python
    row_idx = cellbase.insert('Simple', {'id': 1, 'name': 'jp'})
    values = cellbase.query('Simple', {'row_idx': row_idx})
    cellbase.update('Simple', {'row_idx': row_idx, 'id': 1, 'name': 'imjp'})
    cellbase.delete('Simple', {'row_idx': row_idx})
    ```

- With DAO & Entity:

    First create DAO,
    ```python
    dao = SimpleDAO(cellbase)
    ```
    then do what the last example did,

    except saving declaration of table name &
    access data from object inherits Entity
    ```python
    entity = Simple(id=1, name='jp')
    dao.insert(entity)
    dao.query({'row_idx': entity.row_idx})
    entity.name = 'imjp'
    dao.update(entity)
    dao.delete({'row_idx': entity.row_idx})
    ```

Finally, save it to file

```python
cellbase.save()
```

## More

### Cellbase load, save, save_as, drop, register

Load from file

```python
cellbase.load('filename.xlsx')
```

Save to filename used in load, otherwise,
current working directory as 'cellbase.xlsx'

```python
cellbase.save()
```

Save as another file, will raise FileExistsError if overwrite is False

```python
cellbase.save_as('another_filename.xlsx', overwrite=True)
```

Drop worksheet

```python
cellbase.drop('worksheet_name')
# or drop with DAO
dao.drop()
```

Register structure of worksheet to deal with(only required for new worksheet),
otherwise, ValueError will be raised when creating worksheet as Cellbase
doesn't know what are the title of worksheet and column names to create.

```python
cellbase.register({'TABLE_NAME_1': ['COL_NAME_1', 'COL_NAME_2']})
```

### Example of DAO & Entity

DAO

```python
from cellbase import DAO


class SimpleDAO(DAO):
    # Optional, just to make life easier
    TABLE_NAME = 'Simple'
    COL_ID = 'id'
    COL_NAME = 'name'

    def worksheet_name(self):
        return SimpleDAO.TABLE_NAME

    def new_entity(self):
        return Simple()  # New instance of entity for query to return result
```

Entity

```python
from cellbase import Entity


class Simple(Entity):
    def __init__(self, id=0, name=""):
        super().__init__()  # Declare row_idx
        self.id = id
        self.name = name

    def from_dict(self, values):
        super().from_dict(values)  # Inherits to handle row_idx
        self.id = values[SimpleDAO.COL_ID]
        self.name = values[SimpleDAO.COL_NAME]
        return self

    def to_dict(self):
        values = super().to_dict()  # Inherits to handle row_idx
        values[SimpleDAO.COL_ID] = self.id
        values[SimpleDAO.COL_NAME] =  self.name
        return values
```

### Lambda

After getting used with Cellbase you might find that simple
equality search like this is not enough:

```python
dao.query({'id': 1, 'name': 'imjp'})
```

For example, if you need to access all records where name contains 'jp',
you might find lambda useful:

```python
dao.query({'name': lambda value: 'jp' in value})
dao.update(entity, {'name': lambda value: 'jp' in value})

cellbase.query(worksheet_name, {'name': lambda value: 'jp' in value})
cellbase.update(worksheet_name, data, {'name': lambda value: 'jp' in value})
# So as traverse & format...
```

or find with row_idx

```python
dao.query({'row_idx': lambda row_idx: 3 <= row_idx <= 9})
dao.update(entity, {'row_idx': lambda row_idx: 3 <= row_idx <= 9})

cellbase.query(worksheet_name, {'row_idx': lambda row_idx: 3 <= row_idx <= 9})
cellbase.update(worksheet_name, data, {'row_idx': lambda row_idx: 3 <= row_idx <= 9})
# So as traverse & format...
```

### Magic method(Must implement DAO & Entity)

```python
# Magic method only works with row_idx
total_row_number = len(dao)  # __len__
entity = dao[row_idx]  # __getitem__
dao[row_idx] = entity  # __setitem__
contains = row_idx in dao  # __contains__
del dao[row_idx]  # __delitem

# Of course it works with lambda/callable too
entity = dao[lambda row_idx: 3 <= row_idx <= 9]  # __getitem__
contains = lambda row_idx: 3 <= row_idx <= 9 in dao  # __contains__
del dao[lambda row_idx: 3 <= row_idx <= 9]  # __delitem
# Exception
# __setitem__ only support update, insertion will raise warning
if lambda row_idx: 3 <= row_idx <= 9 in dao:
    dao[lambda row_idx: 3 <= row_idx <= 9] = entity  # update
else:
    dao[lambda row_idx: 3 <= row_idx <= 9] = entity  # no effect at all
```

### Formatting

Other than setting value, you may format cells as well:

```python
dao.format(where, select, fill, font, border...)
# or wrap all formats in CellFormatter
dao.format(where, select, cell_formatter)
```
See [CellFormatter](cellbase/helper/helper.py), for more information.

### Low Level Access

Low level might be a strong word, but you can have direct access to
cells(openpyxl.cell.Cell) through traverse:

```python
dao.traverse(lambda cell: do_something(cell), where, select)
```

### For more example, checkout [Tests](tests/cellbase_test.py)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
