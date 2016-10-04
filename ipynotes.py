# -*- coding: utf-8 -*-
"""
Created on Fri Sep 30 15:37:48 2016

@author: mkonrad
"""

# see also https://xlsxwriter.readthedocs.io/working_with_pandas.html

import numpy as np
import pandas as pd

from pandas import compat
import pandas.formats.format as fmt


#%%

class ExcelFormatterStyler(fmt.ExcelFormatter):
    def __init__(self, *args, **kwargs):
        self.cell_styles = kwargs.pop('cell_styles')
        
        super(ExcelFormatterStyler, self).__init__(*args, **kwargs)
        
        if self.cell_styles is not None:
            assert self.cell_styles.shape[0] == self.df.shape[0]
            assert self.cell_styles.shape[1] == self.df.shape[1] + 1

    def _format_regular_rows(self):
        for cell in super(ExcelFormatterStyler, self)._format_regular_rows():
            #print(cell.row, cell.col, cell.val)
            if self.cell_styles is not None:
                st = self.cell_styles[cell.row - 1, cell.col]
                if st is not None:
                    cell.style = st
            yield cell

class DataFrameExcelStyler(pd.DataFrame):
    def to_excel(self, excel_writer, sheet_name='Sheet1', na_rep='',
                 cell_styles=None,   # new argument
                 float_format=None, columns=None, header=True, index=True,
                 index_label=None, startrow=0, startcol=0, engine=None,
                 merge_cells=True, encoding=None, inf_rep='inf', verbose=True):
        from pandas.io.excel import ExcelWriter
        need_save = False
        if encoding is None:
            encoding = 'ascii'

        if isinstance(excel_writer, compat.string_types):
            excel_writer = ExcelWriter(excel_writer, engine=engine)
            need_save = True

        formatter = ExcelFormatterStyler(self, na_rep=na_rep, cols=columns,
                                         header=header,
                                         cell_styles=cell_styles,   # new argument
                                         float_format=float_format, index=index,
                                         index_label=index_label,
                                         merge_cells=merge_cells,
                                         inf_rep=inf_rep)
        formatted_cells = formatter.get_formatted_cells()
        excel_writer.write_cells(formatted_cells, sheet_name,
                                 startrow=startrow, startcol=startcol)
        if need_save:
            excel_writer.save()



np.random.seed()

col1 = np.random.random(20)
col2 = np.random.randint(0, 11, 20)
col3 = np.random.choice(list('abcf'), 20)

df = DataFrameExcelStyler.from_items([('one', col1), ('two', col2), ('three', col3)])

bold_style = {"font": {"bold": True}}
red_font_style = {"font": {"color": "red"}}
red_bg_style = {"pattern": {"pattern": "solid_fill", "fore_color": "red"}}
orange_bg_style = {"pattern": {"pattern": "solid_fill", "fore_color": "orange"}}

cell_styles = np.empty((df.shape[0], df.shape[1] + 1), dtype='object')
cell_styles.fill(None)

### Test 1 ###

cell_styles[1, 1] = red_font_style
cell_styles[2, 2] = bold_style
cell_styles[3, 3] = red_bg_style

print(cell_styles)

df.to_excel('tmp/test.xls', cell_styles=cell_styles)    # uses xlwt, works
#df.to_excel('tmp/test.xlsx', cell_styles=cell_styles)  # uses openpyxl, doesn't work

### Test 2 ###

cell_styles = np.empty((df.shape[0], df.shape[1] + 1), dtype='object')
cell_styles.fill(None)

cell_styles[df.one.values < 0.25, 1] = orange_bg_style
cell_styles[(df.one.values >= 0.25) & (df.one.values < 0.5), 1] = red_bg_style
print(cell_styles)

df.to_excel('tmp/test2.xls', cell_styles=cell_styles)    # uses xlwt, works

#%%

### Test 3 ###

df['one_valid'] = df['one'] >= 0.5
df['three_valid'] = df['three'].isin(list('abc'))

def create_style_for_validations(df, suffix='_valid', error_style='red', remove_validation_cols=False):
    cell_styles = np.empty((df.shape[0], df.shape[1] + 1), dtype='object')
    cell_styles.fill(None)
    
    if type(error_style) == str:
        error_style = {"pattern": {"pattern": "solid_fill", "fore_color": error_style}}
    elif type(error_style) != dict:
        raise ValueError("Parameter 'error_style' must be either of type 'str' (background color) or a style 'dict'")
        
    for col_idx, colname in enumerate(df.columns.values):
        if colname.endswith(suffix):
            continue
        validation_colname = colname + suffix
        
        if validation_colname in df.columns.values:
            cell_styles[~df[validation_colname].values, col_idx + 1] = error_style
            if remove_validation_cols:
                validation_col_idx = np.nonzero(df.columns == validation_colname)[0][0] + 1
                cell_styles = np.delete(cell_styles, validation_col_idx, axis=1)
                del df[validation_colname]
    
    return cell_styles

cell_styles = create_style_for_validations(df, remove_validation_cols=True)

print(cell_styles)

df.to_excel('tmp/test3.xls', cell_styles=cell_styles)    # uses xlwt, works