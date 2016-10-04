# -*- coding: utf-8 -*-
"""
Created on Fri Sep 30 15:37:48 2016

@author: mkonrad
"""

import numpy as np
import pandas as pd

from pandas_excel_styler import DataFrameExcelStyler, create_style_for_validations


#%% Create some random data

np.random.seed()

col1 = np.random.random(20)
col2 = np.random.randint(0, 11, 20)
col3 = np.random.choice(list('abcf'), 20)

# create a data frame
df = pd.DataFrame.from_items([('one', col1), ('two', col2), ('three', col3)])

# wrap it as DataFrameExcelStyler
# this will not copy the data from the original data frame (unless you set copy=True)!
df = DataFrameExcelStyler(df)

# of course you could also fill a DataFrameExcelStyler directly
#df = DataFrameExcelStyler.from_items([('one', col1), ('two', col2), ('three', col3)])

### define some styles

# for the colors, see https://github.com/python-excel/tutorial/raw/master/python-excel.pdf page 33
# "Colours in Excel files are a confusing mess"

bold_style = {"font": {"bold": True}}
red_font_style = {"font": {"color": "red"}}
red_bg_style = {"pattern": {"pattern": "solid_fill", "fore_color": "red"}}
orange_bg_style = {"pattern": {"pattern": "solid_fill", "fore_color": "orange"}}


### Example 1 ###

# create a cell_styles matrix
# it must have the same number of rows and columns as your DataFrame
cell_styles = np.empty((df.shape[0], df.shape[1]), dtype='object')
cell_styles.fill(None)  # filling it with None means that no styling is applied to all cells

# set some styles
cell_styles[0, 0] = bold_style
cell_styles[1, 1] = red_font_style
cell_styles[2, 2] = red_bg_style

print(cell_styles)

# output as example output
f_out = 'example_output/example1.xls'
print("writing output to file", f_out)
df.to_excel(f_out, cell_styles=cell_styles)    # uses xlwt, works
#df.to_excel('example_output/test.xlsx', cell_styles=cell_styles)  # uses openpyxl, doesn't work

#%%

### Example 2 ###

# Conditional coloring

cell_styles = np.empty((df.shape[0], df.shape[1]), dtype='object')
cell_styles.fill(None)

# all values in column "one" below 0.25 will get a red background
# values between 0.25 and 0.5 will get an orange background
col_idx = 0
cell_styles[df.one.values < 0.25, col_idx] = red_bg_style
cell_styles[(df.one.values >= 0.25) & (df.one.values < 0.5), col_idx] = orange_bg_style

print(cell_styles)

f_out = 'example_output/example2.xls'
print("writing output to file", f_out)
df.to_excel(f_out, cell_styles=cell_styles)

#%%

### Test 3 ###

# style from validation columns

# set some boolean validation columns
df['one_valid'] = df['one'] >= 0.5
df['three_valid'] = df['three'].isin(list('abc'))

# create the styles
cell_styles = create_style_for_validations(df, remove_validation_cols=True)

print(cell_styles)

f_out = 'example_output/example3.xls'
print("writing output to file", f_out)
df.to_excel(f_out, cell_styles=cell_styles)
