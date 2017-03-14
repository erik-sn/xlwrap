## xlwrap
[![Build Status](https://travis-ci.org/erik-sn/xlwrap.svg?branch=master)](https://travis-ci.org/erik-sn/xlwrap)
[![codecov](https://codecov.io/gh/erik-sn/xlwrap/branch/master/graph/badge.svg)](https://codecov.io/gh/erik-sn/xlwrap)

### Purpose
This project is meant to be a simple wrapper around xlrd and 
openypxl. This allows for manipulating of .xls, .xlsx, and .xlsm
files with the same library. The API and feature-set are purposefully small
in order to better facilitate new programmers or people using python as
a tool for productivity.


#### Installation
```commandline
pip install xlwrap
```

#### Running Tests
```commandline
git clone https://github.com/erik-sn/xlwrap.git
cd xlwrap
python -m unittest discover tests
```

#### Usage
This library is not designed to generate new workbooks, mainly to open, read, and manipulate existing ones. It is also more
data oriented - styling and aesthetics are mostly excluded.
    
    Notes:
    - The API is identical for all supported file types: .xls, .xlsx, .xlsm
    - All functions use a 1-based indexing in order to make the translation
    from excel to programming more intuitive
    - .xls does not currently support writing or saving
    
 - [Initializing](#initializing) 
 - [Changing Sheet](#changing-sheet) 
 - [Retrieving a Value](#retrieving-a-value) 
 - [Retrieving a Row](#retrieving-a-row) 
 - [Retrieving a Column](#retrieving-a-column) 
 - [Inserting a Value](#inserting-a-value) 
 - [Searching](#searching) 
 - [Array Representation](#array) 
 - [Manager Information](#manager-information) 
 - [Saving](#saving)
 - [Closing](#closing) 
 - [Raw Objects](#raw-objects) 


#### Initializing
Initializing a manager object opens the file you have specified
with the library corresponding to its file extension (**xls**, **xlsx**
or **xlsm**). Then the first sheet in the workbook is opened and set
to be operated on.
```python
from xlwrap import ExcelManager
manager = ExcelManager('path/to/file.xlsx')
```

#### Changing Sheet
`.change_sheet()`

By default the first sheet is opened, change it with this:
```python
manager.change_sheet(2)  # 1 based index!
# or
manager.change_sheet('Sheet2')
```

#### Retrieving a Value: 
`.read(row, column)`, `.read(cell_name)`

row, column cell references or string based excel references are supported
```python
value = manager.read(1, 1)
# or
value = manager.read('A1')
```
#### Retrieving a Row
`.row(row_index)`

Retrieve a list of all cell values at this row index:
```python
row = manager.row(1)
```

#### Retrieving a Column
`.column(column_index)`, `.column(name)`

Retrieve a list of all cell values at this column index:
```python
column = manager.row(1)
# or
column = manager.row('A')
```

#### Inserting a Value
`.write(row, column, value=value)`, `.write(cell_name, value=value)`

row, column cell references or string based excel references are supported
```python
manager.write(1, 1, value='test this!')
# or
manager.write('A1', value='test this!')
```

#### Searching
`.search(value)`

Returns a tuple in the format `(row, column)` if the value exists, `(None, None)` otherwise. By default this
search is **case insensitive**. 
```python
# A1 = 'Find This Value'
row, column = manager.search('find this value')  # (1, 1)
row, column = manager.search('find', contains=True) # match any partial matches (1, 1)
row, column = manager.search('find this value', case_insensitive=False) # search case sensitive, (None, None)
```
By default the first match is returned. You can change this to second, third, etc. with `match`:
```python
# C3 = 'find this value'
row, column = manager.search('find this value', match=2)  # (3, 3)
```
You can also retrieve a list of all matches with `many`:
```python
indexes = manager.search('find this, value', many=True)  # [(1, 1), (3, 3)]
indexes = manager.search('does not exist', many=True) # []
```
#### Array
`.array()`

This returns a 2D list of lists containing all values in the current sheet
```python
array = manager.array()
```
#### Manager Information
`.info()`

Return some basic information about the manager in a dictionary:
```python
info = manager.info()
"""
{
    'file': *file path*,
    'sheet': *current sheet name*,
    'reads': *how many reads have taken place*,
    'writes': *how many writes have taken place*,
    
}
"""
info = manager.info(string=True) # string representation of dictionary
```

#### Saving
`.save()`

save the currently opened workbook:

```python
manager.save()  # save in the same location you read from
manager.save('new/filepath/here.xlsx')  # save in a new location
```

#### Raw Objects
At any time you can pull the raw `workbook` and `worksheet` objects from the manager
and use the corresponding openpyxl/xlrd api to operate on them:
```python
workbook = manager.workbook
sheet = manager.sheet
```
