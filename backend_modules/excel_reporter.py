#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Sep 26 11:56:23 2024

@author: ericlysenko
"""

import math
import pandas as pd
import sys
import os
sys.path.append(os.getcwd())
try:
    from backend_modules.xlFuncs import xlFuncs 
    from backend_modules.xlIntrfc import xlIntrfc 
except:
    from .xlFuncs import xlFuncs 
    from .xlIntrfc import xlIntrfc 
from openpyxl import Workbook

class excel_reporter(xlFuncs):
    def __init__(self):
        pass
    
    def create_report(self,base_df,**kwargs):
        """Creates and formats workbook
        **kwargs:
            wb=None,
            header=None,
            sheet_name='Sheet1'"""
        
        df = base_df.copy()
        ## ********************* unpack kwargs *******************************
        if 'hide_cols' in kwargs.keys():
            hide_cols = kwargs['hide_cols']
        else:
            hide_cols = []
        
        if 'mult_idcs' in kwargs.keys():
            mult_idcs = kwargs['mult_idcs']
        else:
            mult_idcs = df.index.tolist()
        
        # get the sheet name
        if 'sheet_name' in kwargs.keys():
            sheet_name = kwargs['sheet_name']
        else:
            sheet_name = 'Sheet1'
        
        # get the wb
        if 'wb' in kwargs.keys():
            wb = kwargs['wb']
            ws = wb.create_sheet(sheet_name)
        else:
            wb = Workbook()
            ws = wb.worksheets[0]
            ws.title = sheet_name
        
        # get fontcolors
        # format:{<color>:[col1,col2,...]}
        if 'fontcolor_cols' in kwargs.keys():
            fontcolor_cols = kwargs['fontcolor_cols']
        else:
            fontcolor_cols = {}
        # ********************************************************************
        
        # initialize xlInterface
        xli = xlIntrfc()
        
        # move the columns names into the dataframe
        df = self.hdrToLine(df)
        mult_idcs = [idx + 1 for idx in mult_idcs]
        
        # include top level header
        fontsize_dic = {}
        if 'header' in kwargs:
            # find index to put header in
            start_col = 0 + len(hide_cols)
            header = kwargs['header']
            
            # create new top line
            top_data = [''] * len(df.columns)
            top_data[start_col] = header
            df.loc[-1] =  top_data # adding a row
            df.index = df.index + 1  # shifting index
            df.sort_index(inplace=True) 
            
            # create new dataframe reference
            df_ref = self.wsToCellRef(ws,df,vert=-1)
            
            # realign internal refs in the df
            df = self.shiftWsRef(ws,df,vert=1)
            mult_idcs = [idx + 1 for idx in mult_idcs]
            
            # set the header font size
            cell_col = df_ref.iat[0,start_col]
            fontsize_dic[cell_col] = 24
        else:
            df_ref = self.wsToCellRef(ws,df,vert=-1)
        
        # create the mapping dictionary
        df_ref = self.wsToCellRef(ws,df,vert=-1)
        
        # hide columns
        for col in hide_cols:
            col_idx = df.columns.tolist().index(col)
            col_letter = self.numToCol(col_idx + 1)
            ws.column_dimensions[col_letter].hidden= True
        
        # format data as a table
        df_sub = df[[col for col in df.columns if not col in hide_cols]]
        # get indices of all nas
        df_no_na = df_sub.dropna(how='all',axis=0)
        na_idcs = list(set(df_sub.index) - set(df_no_na.index))
        # partition the df into sub_df
        sections = []
        for i in range(len(na_idcs)):
            top_idx = na_idcs[i]
            remaining = len(na_idcs) - i
            if remaining > 1:
                next_idx = na_idcs[i + 1]
                section = df_ref.iloc[top_idx+1:next_idx]
            elif remaining == 1:
                section = df_ref.iloc[top_idx:]
            else:
                pass
            
            if len(section) > 0:
                sections.append(section)
        
        # fontcolor
        fontcolor_dic = {}
        for color in fontcolor_cols.keys():
            cols = fontcolor_cols[color]
            for col_hdr in cols:
                col = df_ref.loc[mult_idcs,col_hdr]
                for cell_addr in col:
                    fontcolor_dic[cell_addr] = color
        
        # dollar cells
        if 'dllr_cols' in kwargs.keys():
            dllr_cols = kwargs['dllr_cols']
        if 'prc_cols' in kwargs.keys():
            prc_cols = kwargs['prc_cols']
        
        dllr_cells = []
        prc_cells = []
        for col in dllr_cols:
            dllr_cells = dllr_cells + df_ref[col].tolist()
        for col in prc_cols:
            prc_cells = prc_cells + df_ref[col].tolist()
        
        # paste and format the ws
        map_dic = self.mapDic(df, df_ref)
        xli.insrIdx(ws, map_dic)
        xli.frmtWs(ws,data_dfs=sections,fontsize_dic=fontsize_dic,fontcolor_dic=fontcolor_dic,prc_cells=prc_cells,dllr_cells=dllr_cells)
        
        return wb
        
        
        