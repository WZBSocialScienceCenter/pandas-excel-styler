# -*- coding: utf-8 -*-
"""
Extended classes for styling individual cells of pandas Excel exports.
Currently, only export to "xls" files via xlwt is supported. "xlsx" export via OpenPyXL is not available.

Created on Tue Oct  4 11:32:10 2016

@author: Markus Konrad <markus.konrad@wzb.eu>
"""


import numpy as np
import pandas as pd

from pandas import compat
import pandas.formats.format as fmt

#%% Class extensions to pandas

class ExcelFormatterStyler(fmt.ExcelFormatter):
    """
    Extended ExcelFormatter class that accepts additional "cell_styles" matrix
    """
    
    def __init__(self, *args, **kwargs):
        """
        Extended constructor that accepts additional "cell_styles" matrix
        """
        self.cell_styles = kwargs.pop('cell_styles')
        
        super(ExcelFormatterStyler, self).__init__(*args, **kwargs)
        
        if self.cell_styles is not None:
            expected_n_rows = self.df.shape[0]
            expected_n_cols = self.df.shape[1]
            if self.cell_styles.shape[0] != expected_n_rows or self.cell_styles.shape[1] == expected_n_cols:
                ValueError("Argument 'cell_styles' must have the same shape like the data frame: %dx%d"
                           % (expected_n_rows, expected_n_cols))

    def _format_regular_rows(self):
        """
        Extended function that handles formatting of regular cells also considering cell_styles matrix
        """
        # get cell formatting from parent method
        col_offset = -1 if self.index else 0
        for cell in super(ExcelFormatterStyler, self)._format_regular_rows():
            if self.cell_styles is not None and cell.row > 0 and cell.col > 0:
                # consider cell style for this regular cell
                st = self.cell_styles[cell.row - 1, cell.col + col_offset]
                if st is not None:
                    if type(cell.style) != dict:
                        cell.style = st
                    else:
                        cell.style.update(st)
            yield cell


class DataFrameExcelStyler(pd.DataFrame):
    """
    Extended DataFrame class that accepts additional "cell_styles" matrix and uses the extended
    ExcelFormatterStyler as formatter class.
    """
    
    def to_excel(self, excel_writer, sheet_name='Sheet1', na_rep='',
                 cell_styles=None,   # new argument
                 float_format=None, columns=None, header=True, index=True,
                 index_label=None, startrow=0, startcol=0, engine=None,
                 merge_cells=True, encoding=None, inf_rep='inf', verbose=True):
        """
        Extended function that adds support for "cell_styles" argument
        """
        from pandas.io.excel import ExcelWriter
        need_save = False
        if encoding is None:
            encoding = 'ascii'

        if isinstance(excel_writer, compat.string_types):
            excel_writer = ExcelWriter(excel_writer, engine=engine)
            need_save = True
        
        # use the extended formatter class and pass the cell_styles argument
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


#%% Utility functions

def create_style_for_validations(df, suffix='_valid', error_style='red', remove_validation_cols=False):
    """
    Create a "cell_styles" matrix for a data frame with boolean validation result columns suffixed
    with <suffix>.
    """
    # check arguments
    if type(error_style) == str:
        error_style = {"pattern": {"pattern": "solid_fill", "fore_color": error_style}}
    elif type(error_style) != dict:
        raise ValueError("Argument 'error_style' must be either of type 'str' (background color) or a style 'dict'")

    
    # create empty cell style matrix
    cell_styles = np.empty((df.shape[0], df.shape[1]), dtype='object')
    cell_styles.fill(None)
        
    # iterate through the columns
    for col_idx, colname in enumerate(df.columns.values):
        if colname.endswith(suffix):
            continue
        validation_colname = colname + suffix
        
        if validation_colname in df.columns.values:   # found a validation result column
            # set the style for all "invalid" cells
            cell_styles[~df[validation_colname].values, col_idx] = error_style
            
            if remove_validation_cols:  # optionally remove the validation result column
                # remove from cell_styles
                validation_col_idx = np.nonzero(df.columns == validation_colname)[0][0]
                cell_styles = np.delete(cell_styles, validation_col_idx, axis=1)
                
                # remove from the original data frame
                del df[validation_colname]
    
    return cell_styles

