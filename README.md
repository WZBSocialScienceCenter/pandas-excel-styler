# Excel Styler for pandas

Styling individual cells in Excel output files created with [Python's Data Analysis Library pandas](http://pandas.pydata.org/).

## Description

This small addition to *pandas* allows to specify styles for individual cells in Excel exports from pandas DataFrames,
which is not possible with the existing API methods. With this you can for example highlight certain cells by
specifying font styles, font colors or background patterns and colors.

Of course it is also possible to add things like conditional formatting and other advanced functions to Excel files
with [XlsxWriter](http://xlsxwriter.readthedocs.io/working_with_pandas.html) (see also
[Improving Pandas’ Excel Output](http://pbpython.com/improve-pandas-excel-output.html)). However, sometimes it is
necessary set styles like font or background colors on individual cells on the "Python side". In this scenario,
XlsxWriter won't work, since
[*"XlsxWriter and Pandas provide very little support for formatting the output data from a dataframe apart from default formatting such as the header and index cells and any cells that contain dates of datetimes."*](http://xlsxwriter.readthedocs.io/working_with_pandas.html#formatting-of-the-dataframe-output)
In this case, you're better off with this tool, for example when you are running complicated data validation
routines (which you probably don't want to implement in VBA) and want to highlight the validation results by coloring
individual cells in the output Excel sheets.

## Example

```python
import numpy as np
import pandas as pd

from pandas_excel_styler import DataFrameExcelStyler

# create a data frame and work with it
df = ...

# wrap it as DataFrameExcelStyler
# this will not copy the data from the original data frame (unless you set copy=True)!
df = DataFrameExcelStyler(df)

# define a style

red_font_style = {"font": {"color": "red"}}

# create a cell_styles matrix
# it must have the same number of rows and columns as your DataFrame
cell_styles = np.empty(df.shape, dtype='object')
cell_styles.fill(None)  # filling it with None means that no styling is applied to any cell

# set a style to the top left cell
cell_styles[0, 0] = red_font_style

# write output, use the specified cell style matrix
df.to_excel('test.xls', cell_styles=cell_styles)

```

See `examples.py` for more examples and usage cases.

## Styling the cells

The styling definitions of individiual cells can be done by creating a nested dictionary of style options for
[*xlwt*](https://github.com/python-excel/xlwt). The dictionary will be converted to "easyxf" directives by pandas so
that `{"font": {"color": "red"}}` will become `"font: color red;"`. See the
[Python Excel Tutorial](https://github.com/python-excel/tutorial/raw/master/python-excel.pdf), page 28 for available
options.

### Colors in Excel

For an overview of possible colors, see the
[Python Excel Tutorial](https://github.com/python-excel/tutorial/raw/master/python-excel.pdf), page 33. But be aware
that *"Colours in Excel files are a confusing mess"* (ibid.).

## Requirements

* tested with pandas 0.18.1 and 0.19.0
  * **only works with "xls" files via xlwt engine ("xlsx" support via OpenPyXL is not yet available)**
* NumPy is required
