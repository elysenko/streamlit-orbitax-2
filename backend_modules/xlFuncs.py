# -*- coding: utf-8 -*-
"""
Created on Tue Oct  3 11:31:31 2023

@author: 1148055
"""
import string
import numpy as np
import pandas as pd
import math

class xlFuncs:
    def __init__(self):
        self.rngs = {} # use to hold the addresses of all the important variables on each worksheet
    
    def colToNum(self, col):
        """ converts an excel column to a number """
        
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num
    
    def numToCol(self,column_int):
        """ converts a number to an excel column """
        
        start_index = 1   #  it can start either at 0 or at 1
        letter = ''
        while column_int > 25 + start_index:
            letter += chr(65 + int((column_int-start_index)/26) - 1)
            column_int = column_int - (int((column_int-start_index)/26))*26
        letter += chr(65 - start_index + (int(column_int)))
        
        return letter
    
    def incrCol(self, col_letter, num_to_add):
        """inputs a letter, adds num_to_add to it and then outputs the letter"""
        
        col_num = self.colToNum(col_letter)
        col_num = col_num + num_to_add
        new_letter = self.numToCol(col_num)
        
        return new_letter
    
    def getCellVal(self,addr, ws=None, df=None):
        """returns the value on a sheet of a specified address"""
        
        if ws == None:
            ws = self.ws
        
        if df == None:
            df = self.getWsDataOpxl(ws)
        
        # get the df_ref
        df_ref = self.wsToCellRef(ws,df)
        
        # return the column and index of the address
        idx, col = (df_ref == addr).stack().idxmax()
        
        # use the col and idx to get the actual value from the dataframe
        val = df.loc[idx,col]
        
        return val
        
    
    def getCellAddr(self,ws,df,val,num=0,bas=1,offs=0):
        """ searches a df for a value and returns the Excel location for it """
        
        # i is the row, c is the column
        i, c = np.where(df == val)
        idx = i[num] + 2 + offs # adjust for different bases and the header row and provides another offset if needed
        col = c[0] + bas # adjust for the basis
        col = self.numToCol(col)
        addr = f"'{ws.title}'!{col}{idx}"
        
        return addr
    
    def getCellRow(self,ws,df,val,offs=0):
        """ Takes away the letter in front of the address """
        
        addr = self.getCellAddr(ws,df,val,offs=offs)
        row_idx, col_idx = self.addrToIdcs(addr)
        
        return row_idx
    
    def getCellCol(self,ws,df,val,offs=0):
        """ Takes away the letter in front of the address """
        
        addr = self.getCellAddr(ws,df,val,offs=offs)
        addr = addr.split("!")[-1]
        row_idx, col_idx = self.addrToIdcs(addr)
        
        col = self.numToCol(col_idx)
        
        return col
    
    def getColLtr(self,df,hdr,horiz=0):
        """ returns the letter of a column in Excel """
        
        col_num = df.columns.get_loc(hdr) + 1 + horiz # offset for basis change
        col_ltr = self.numToCol(col_num)
        
        return col_ltr
    
    def getColIdx(self,cell_addr,cell_idx=-1):
        """Returns the column index of a cell""" 
    
        addr = cell_addr.split("!")[-1]
        row_idx, col_idx = self.addrToIdcs(addr,cell_idx)
        
        return col_idx
    
    def getRowIdx(self,cell_addr,cell_idx=-1):
        """Returns the column index of a cell""" 
    
        addr = cell_addr.split("!")[-1]
        row_idx, col_idx = self.addrToIdcs(addr,cell_idx)
        
        return row_idx
    
    def swapIdcs(self,df,first,sec):
        """Exchanges the first and second indices in the dataframe"""
        
        temp = df.copy()
        df.loc[first] = temp.loc[sec]
        df.loc[sec] = temp.loc[first]
        
        return df
    
    def rngToIdcs(self,rng):
        """Converts a range to start and stop row/col indices"""
        
        rng_dic = {}
        
        start_row_idx, start_col_idx = self.addrToIdcs(rng,cell_idx=0)
        rng_dic['min_row'] = start_row_idx
        rng_dic['min_col'] = start_col_idx
        
        end_row_idx, end_col_idx = self.addrToIdcs(rng,cell_idx=-1)
        rng_dic['max_row'] = end_row_idx
        rng_dic['max_col'] = end_col_idx
        
        return rng_dic
    
    def addrToIdcs(self,addr,cell_idx=-1):
        """Converts a cell address to the column and row indices"""
        
        addr = addr.split("!")[-1]
        addr = addr.split(':')[cell_idx]
        addr_list = [*addr]
        row_list = []
        col_list = []
        for elem in addr_list:
            try:
                elem = int(elem)
                row_list.append(str(elem))
            except:
                col_list.append(elem)
        
        col =  "".join(col_list)
        
        col_idx = self.colToNum(col)
        row_idx = int("".join(row_list))
    
        return row_idx, col_idx
    
    def getWsDataOpxl(self,ws=None,col_hdrs=True):
        """
        gets the data in a worksheet as a df from Openpyxl
        col_hdrs=True -> the top row of the worksheet becomes the column headers
        """
        
        if ws == None:
            ws = self.ws
            
        # extract rows as tuple values
        tuple_vals = list(ws.values)
        
        # convert to a dataframe
        df = pd.DataFrame(tuple_vals)
        
        # format cells
        df = df.fillna("")
        df = df.astype(str)
        
        # cast the first row as the column headers
        if col_hdrs:
            df = self.lineToHdr(df).reset_index(drop=True)
            
        return df
    
    def getWsDataXw(self,ws=None):
        """ gets the data in a worksheet as a df from Xlwings"""
        
        if ws == None:
            ws = self.ws
        df = ws.used_range.options(pd.DataFrame, 
                             header=1,
                             index=False, 
                             ).value            
        
        return df
    
    def mapDic(self,df,df_ref,purge=False):
        """zips a df and df_ref together into a dictionary of address:value pairs"""
        
        # Example DataFrames
        df_copy = df.copy()
        if purge:
            df_copy = self.purgeWsRef(self.ws, df)
        
        # Get the index and column labels of df1
        df_index_labels = df_copy.index
        df_column_labels = df_copy.columns
        # Get the index and column labels of df1
        df_ref_index_labels = df_ref.index
        df_ref_column_labels = df_ref.columns
        
        # Create a MultiIndex from the index and column labels
        df_multi_index = pd.MultiIndex.from_product([df_index_labels, df_column_labels], names=['row', 'column'])
        
        # Create a DataFrame with the values from df1 and the corresponding row and column indices
        df_combined = pd.DataFrame({'value_df': df_copy.values.flatten(),
                                    'row': df_multi_index.get_level_values('row'),
                                    'column': df_multi_index.get_level_values('column')})
        
        # Create a MultiIndex from the index and column labels
        df_ref_multi_index = pd.MultiIndex.from_product([df_ref_index_labels, df_ref_column_labels], names=['row', 'column'])
        
        # Create a DataFrame with the values from df1 and the corresponding row and column indices
        df_ref_combined = pd.DataFrame({'value_df_ref': df_ref.values.flatten(),
                                    'row': df_ref_multi_index.get_level_values('row'),
                                    'column': df_ref_multi_index.get_level_values('column')})
        
        
        # Create the dictionary from the merged DataFrame
        mapping_dict = dict(zip(df_ref_combined['value_df_ref'], df_combined['value_df']))
        
        return mapping_dict

    
    def wsToCellRef(self,ws = None, df = None, horiz=0, vert=0):
        """Takes a df and returns an identically formatted df with cell numbers instead of data"""
        
        if ws == None:
            ws = self.ws
        
        if not isinstance(df,pd.DataFrame):
            df = self.getWsDataOpxl(ws)
        
        # get the original index
        all_idcs = df.index.to_list()
        idx_name = df.index.name
        
        # reset the index for ease of pasting
        df.reset_index(drop=True,inplace=True)
        
        df_ref = df.copy()
        name = ws.title
        
        prfx = f"'{name}'!"
        
        idx = pd.Series([i + 2 for i in range(len(df_ref))]).astype(str) # accounts for the change in header and the header as the first row
        # Assume the header is the first row of the df
        df_ref = df_ref.astype(str)
        
        # create a df from the cols
        excel_col_nums = list(range(1,len(df.columns)+1))
        excel_cols = pd.Series([self.numToCol(col) for col in excel_col_nums])
        excel_row = excel_cols.to_frame().transpose()
        excel_cols_df = pd.DataFrame(np.repeat(excel_row.values, len(idx), axis=0))
        
        # create a df from the idx
        idx_col = idx.to_frame()
        idx_df = pd.concat([idx_col] * len(excel_col_nums), axis=1)
        idx_df.columns = excel_cols_df.columns
        
        # create df ref by concacting the prefix, excel cols and idx cols
        df_ref = prfx + excel_cols_df + idx_df
        df_ref.columns = df.columns
        
        # reset the indices to their original
        df_ref.index = all_idcs
        df_ref.index.names = [idx_name]
        # maintain coherence between original df and new df_ref indices
        df.index = all_idcs
        df.index.names = [idx_name]
        
        df_ref = self.shiftRef(df_ref,vert=vert,horiz=horiz)
        
        return df_ref
    
    def purgeWsRef(self,ws,df):
        """Removes the reference to the wbs in the formulas"""
        
        prfx = self.wsPrfx(ws)
        def rmPrfx(value,prfx):
            if isinstance(value,str):
                return value.replace(prfx,"")
            return value
        
        df = df.applymap(lambda x: rmPrfx(x, prfx))
        
        return df
        
    def cellsToCompRng(self,cells):
        """Edits a range to be a rolling addition range"""
        
        start_rng = self.anchrCells(self.cellsToRange(cells))
        start_cell = start_rng.split(":")[0]
        cell_addrs = cells.str.split("!").str[-1]
        
        new_cells = start_cell + ':' + cell_addrs
        
        return new_cells
    
    def getCellRefByOtherCol(self,ws,df,trgt_col,val_col,val):
        """ gives the cell reference in a col where another col cell equals X"""
        
        # get an identically sized dataframe with addresses instead of valuse
        df_cells = self.wsToCellRef(ws,df)
        
        # get the index of the cell with the value
        idx = df[df[val_col] == val].index[0]
        
        # get the column number of trgt_col
        col_letter = self.getColLtr(df,trgt_col)
        col_num = self.colToNum(col_letter) - 1 # - 1 for change in base
        
        # get the cell in the target column with the same index
        cell_ref = df_cells.iloc[idx,col_num]
        
        # return the cell reference
        return cell_ref
        
    def reorder(self,wb,first):
        """ reorders the sheets by putting the input as the first sheet"""
        
        ws1 = wb.worksheets[0]
        ws = wb.sheets(first)
        
        ws.api.Move(Before=ws1.api, After=None)
    
    def shiftRef(self,df_ref,horiz=0,vert=0):
        """Takes in a df_ref and shifts all of the cells horizontally and vertically depending on inputs"""
        
        is_str = False
        is_srs = False
        is_list = False
        
        def allSame(s):
            """Determines if all elements in a series are the same"""
            a = s.to_numpy() # s.values (pandas<0.24)
            return (a[0] == a).all()
        
        def isVert(srs):
            """Determines if a series is horizontal or vertically aligned"""
            prfx = srs.iloc[0].split('!')[0] + '!'
            
            srs = srs.str.replace(prfx,"", regex=True)
            
            letters = srs.str.replace('\d+', '', regex=True)
                                       
            return allSame(letters)
        
        # set logical condition for continuing
        if horiz != 0 or vert != 0:
            # convert to a dataframe if it is a series
            if isinstance(df_ref,list):
                is_list=True
                df_ref = pd.Series(df_ref)
                    
            if isinstance(df_ref,pd.Series):
                is_srs = True
                vert_flag = isVert(df_ref)
                horiz_flag = not vert_flag
                nw_srs = df_ref.copy()
                if vert_flag:
                    new_ref = nw_srs.to_frame()
                elif horiz_flag:
                    new_ref = nw_srs.to_frame().transpose()
                    # new_ref = pd.concat([new_ref]*len(new_ref),axis=1)
                
            elif isinstance(df_ref,str):
                is_str = True
                new_ref = pd.DataFrame([df_ref],index=[0])

            else:
                # create a copy to avoid errors
                new_ref = df_ref.copy()
            
            # identify the sheet prefix
            prfx = new_ref.iloc[0,0].split("!")[0] + '!'
            
            # get a horizontal slice
            horiz_row = new_ref.iloc[0,:]
            
            # remove the prefix
            horiz_row = horiz_row.str.replace(prfx,"", regex=True)
            
            # remove the numbers from the row
            letters = horiz_row.str.replace('\d+', '', regex=True)
            
            # Shift the letters over the specified horiz number
            new_letters = []
            for letter in letters:
                new_letter = self.shiftLetter(letter, horiz)
                new_letters.append(new_letter)
            
            # get a vertical slice
            vert_row = new_ref.iloc[:,0]
            
            # remove the prefix
            vert_row = vert_row.str.replace(prfx,"", regex=True)
            
            # remove the letters from the row
            nums = vert_row.str.replace('\D', '', regex=True)
            
            #  Shift the numbers over by the specified vert number
            new_nums = (nums.astype(int) + int(vert)).astype(str)
    
            # populate the dataframe with the resulting numbers
            for col_num in range(len(new_ref.columns)):
                col = prfx + new_letters[col_num] + new_nums
                new_ref.iloc[:,col_num] = col
                
            if is_srs:
                # new_ref = new_ref.iloc[:,0]
                if vert_flag:
                    new_ref = new_ref.iloc[:,0]
                elif horiz_flag:
                    new_ref = new_ref.iloc[0,:]
            
            if is_list:
                new_ref = new_ref.to_list()
            
            elif is_str:
                new_ref = new_ref.iloc[0,0]
                
            
            return new_ref
        
        # if neither offset condition is met, return the original
        return df_ref
    
    def shiftWsRef(self,ws,df,horiz=0,vert=0,anchr=True):
        """Shifts all references to the ws in df formulas by horiz and vert"""
        
        prfx = self.wsPrfx(ws)
        
        is_series=False
        if isinstance(df,pd.Series):
            # convert the series to a dataframe if it is a series
            df = df.to_frame()
            is_series=True
            
        def shiftRef(cell):
            """Shifts the references in the cell based on the ws name, vert and horiz offsets"""
            
            if isinstance(cell,str) and len(cell) > 0:
                str_list = cell.split(prfx)
                new_cell = str_list[0]
                if len(str_list) > 1:
                    for i in range(1,len(str_list)):
                        part = str_list[i]
                        
                        # get the associated address of the cell
                        addr_parts = []
                        while len(part)>0 and (part[0].isalpha() or part[0].isnumeric() or part[0] == ':' or part[0] == '$'):
                            addr_parts.append(part[0])
                            part = part[1:]
                        addr = ''.join(addr_parts)
                        
                        # convert the address into a new address
                        addr_list = addr.split(':')
                        new_addr_list = []
                        # iterate through all parts of the cell address in case it is in the form 'Name'!A1:A2 
                        for old_addr in addr_list:
                            # retrieve the column indices of the cell address
                            old_row_idx, old_col_idx = self.addrToIdcs(old_addr)
                            new_col_idx = old_col_idx + horiz
                            new_row_idx = old_row_idx + vert
                            new_col = self.numToCol(new_col_idx)
    
                            # update the address lists
                            # the anchr formula is to not touch anchored cells. This is helpful if the frame is referencing a cell in another frame on the same sheet
                            if anchr and '$' in old_addr:
                                new_addr_list.append(old_addr)
                            else:
                                new_addr_list.append(new_col + str(new_row_idx))

                        # Create the new address and add the the existing formula
                        new_addr = prfx + ':'.join(new_addr_list)
                        new_cell = new_cell + new_addr + part
                
                return new_cell
            
            return cell
        
        try:
            df = df.map(lambda cell: shiftRef(cell))
        except:
            df = df.applymap(lambda cell: shiftRef(cell))
            
        if is_series:
            df = df.iloc[:,0]
        
        return df
    
    def shiftLetter(self,letter,horiz):
        """Shifts a letter over a specified number"""
        
        letter_num = self.colToNum(letter)
        new_letter_num = letter_num + int(horiz)
        new_letter = self.numToCol(new_letter_num)
        
        return new_letter
    
    def dfToRange(self,ws, df_ref,horiz=0,vert=0):
        """Intakes a df and returns the range for the max and minimum cells"""
        
        # get the start and end cells
        prfx = self.wsPrfx(ws)
        start_cell = df_ref.iloc[0,0]
        end_cell  = df_ref.iloc[-1,-1]
        
        # find the offset values
        new_start_cell = self.shiftRef(start_cell,horiz=horiz,vert=vert)
        new_end_cell = self.shiftRef(end_cell,horiz=horiz,vert=vert)
        
        # concatenate the two together
        rng = new_start_cell + ':' + new_end_cell.replace(prfx,'')
        
        return rng
        
    
    def cellsToRange(self,cells,horiz=0,vert=0):
        """Takes in a list of cells and returns the range using the form A1:Z10"""
        
        # convert from series to list if is a series
        if isinstance(cells,pd.Series):
            cells = cells.tolist()
            
        elif isinstance(cells,pd.DataFrame):
            cells = cells.iloc[0].tolist()
        
        # see if there is a prefix for the sheet name
        if "!" in cells[0]:
            prfx = cells[0].split('!')[0] + '!'
        else:
            prfx = ''
        
        # put the list into a series
        if not isinstance(cells,pd.Series):
            sr = pd.Series(cells)
        else:
            sr = cells
        
        # remove the prefix
        refs = sr.str.replace(prfx, '',regex=True)
        
        # get a list of all the numbers
        nums = pd.to_numeric(refs.str.replace('\D', '', regex=True))
        
        # get a list of all the letters
        letters = sr.str.replace(prfx,'',regex=True).str.replace('\d+', '',regex=True)
        
        # find the lowest number
        min_num = min(nums)
        
        # find the highest number
        max_num = max(nums)
        
        # find the lowest letter
        letter_nums = [self.colToNum(letter) for letter in letters]
        min_letter = self.numToCol(min(letter_nums))
        
        # find the highest letter
        max_letter = self.numToCol(max(letter_nums))
        
        # make the range
        first = f"{min_letter}{min_num}"
        last = f"{max_letter}{max_num}"
        
        rng = f'{prfx}{first}:{last}'
        
        return rng
    
    def cellsToRangeMulti(self,ws,cell_list,horiz=1,vert=0):
        """Takes in a list of cells and returns the range using the form A1:Z10
        Returns a list of the ranges that can be grouped together"""
        
        # concatenate strings
        rngs_long = []
        for cells in cell_list:
            rng = self.cellsToRange(cells)
            rngs_long.append(rng)
        
        # concatenate ranges that are stacked on each other
        rngs = pd.Series(rngs_long)

        # perhaps future fuctionality to combine vertically
        # rngs = self.combineAdjacentRanges(ws,rngs)
        
        return rngs
    
    def combineAdjacentRanges(self,ws,series):
        """combine adjacent ranges into a single range"""
        ranges = [ws.range(range_str) for range_str in series]
        ranges.sort(key=lambda r: (r.row, r.column))
    
        combined_ranges = []
        current_start, current_end = ranges[0].column, ranges[0].last_cell.column
    
        for cell_range in ranges[1:]:
            if cell_range.row == ranges[0].row and cell_range.column == current_end + 1:
                current_end = cell_range.last_cell.column
            else:
                combined_ranges.append((current_start, current_end, ranges[0].row))
                current_start, current_end = cell_range.column, cell_range.last_cell.column
    
            combined_ranges.append((current_start, current_end, ranges[0].row))
        
        # Convert combined ranges back to Excel range format
        combined_ranges_str = [f"'{ws.title}'!{self.numToCol(start)}{row}:{self.numToCol(end)}{row}" for (start, end, row) in combined_ranges]
        
        # drop duplicates
        combined_ranges_series = pd.Series(combined_ranges_str)
        combined_ranges_series = combined_ranges_series.drop_duplicates()
        
        return combined_ranges_series
    
    def subTblRef(self,ws,df,sub_cols):
        """returns the references to a sub table within a dataframe. 
        The subtable headers are defined by sub_cols"""
        
        df_ref = self.wsToCellRef(ws,df)
        
        # find the row with the column headers
        while not all(item in df.columns for item in sub_cols):
            df = self.lineToHdr(df)
            df_ref = self.lineToHdr(df_ref)
        
        # get the proper columns
        col_idcs = [df.columns.get_loc(col) for col in sub_cols]
        df = df.iloc[:,col_idcs]
        df_ref = df_ref.iloc[:,col_idcs]
        df_ref.columns = sub_cols
        
        return df_ref
    
    def subTbl(self,df,sub_cols):
        """returns a sub table within a dataframe. 
        The subtable headers are defined by sub_cols"""
        
        temp = df.copy()
        while not all(item in temp.columns for item in sub_cols):
            temp = self.lineToHdr(df)
            
        col_idcs = [temp.columns.get_loc(col) for col in sub_cols]
        temp = temp.iloc[:,col_idcs]
        
        return temp
        
    def lineToHdr(self,df,idx=0,rebase=False):
        """Replaces the header of a df with the first row of data"""
        
        # set first row to header by default
        new_header = df.iloc[idx] #grab the first row for the header
        temp = df.copy()
        temp = temp[idx+1:] #take the data less the header row
        temp.columns = new_header #set the header row as the df header
        
        if rebase:
            temp = temp.reset_index(drop=True)
            
        return temp
    
    def hdrToLine(self,df,new_idx=0,cols='',idx=1,drop=False,num=1):
        """Move the header down to the first row"""
        
        # create a new dataframe with the given offset from the header
        df_copy = df.copy()
        
        for i in range(num):
            # create a new dataframe with the given offset from the header
            new_df = pd.DataFrame(columns=df_copy.columns)
            
            # preserve the index name if there is one
            if df_copy.index.name is None:
                new_df.loc[new_idx] = df_copy.columns
            else:
                new_df.loc[df_copy.index.name] = df_copy.columns
            
            # keep index by default in case the indices are strings, not numbers
            new_df = pd.concat([new_df,df_copy], ignore_index=False)
            
            # replace old columns with new ones specified
            repl_cols = isinstance(cols,list) and len(cols) == len(new_df.columns)
            if repl_cols:
                new_df.columns = cols
            
            # reset integer indices
            is_integer_idx = new_df.index.astype(str).str.isdigit().all()
            if is_integer_idx:
                new_df = new_df.reset_index(drop=True)
            
            # remove the old columns if specified
            if drop:
                new_df.columns=["" for col in df.columns]
            
            # create a new dataframe with the given offset from the header
            df_copy = new_df.copy()
            
        return new_df
    
    def setColAsIdx(self,df,col_idx=0):
        """Sets the first column as the index column"""
        
        col1 = df.columns.tolist()[col_idx]
        
        df = df[:,col_idx:]
        
        df.set_index(col1,inplace=True)
        
        return df
    
    def idxToCol(self,df,new_col_hdr='index',keep=True):
        """brings in the column as an index while prserving the index"""
        
        # bring in the index to the column
        df = df.reset_index()
        
        # replace the new column header with the specified new_col_hdr
        df = df.replace({'index':new_col_hdr})
        
        # set the index to the original value
        if keep:
            df.index = df.iloc[:,0]
        
        return df
    
    def keyByVal(self,dic,val):
        """Returns the Key in a dictionary based on the value"""
        key = dic.keys()[dic.values().index(val)] 
        
        return key
    
    def anchrCells(self,cells):
        """converts a cell to anchor form"""
        
        # put in correct format
        if isinstance(cells, str):
            cells = [cells]
        if isinstance(cells,list):
            cells = pd.Series(cells)
        
        # isolate prefix (ie sheet name)
        prfx = cells.iloc[0].split("!")[0] + '!'
        
        # remove the prefix
        naked = cells.str.replace(prfx,"", regex=True)
        
        # split into list
        split = naked.str.split(":")
        
        # split into cell cols for ranges
        cell_prfx = "cells"
        col1_hdr = f"{cell_prfx}1"
        cell1 = split.str.get(0).fillna('')
        cell1.name = col1_hdr
        col2_hdr = f"{cell_prfx}2"
        cell2 = split.str.get(1).fillna('')
        cell2.name=col2_hdr
        df = pd.concat([cell1,cell2], axis=1)
        
        lttr_prfx = "lttr"
        num_prfx  = "num"
        new_df =pd.DataFrame(columns=[f"{lttr_prfx}1",f"{num_prfx}1",f"{lttr_prfx}2",f"{num_prfx}2"],index=df.index)
        for col in df.columns:
            num = col[-1]
            # remove the numbers from the row
            new_df.loc[:,f"{lttr_prfx}{num}"] = df.loc[:,f"{cell_prfx}{num}"].str.replace('\d+', '', regex=True)
            
            # remove the letters from the row
            new_df.loc[:,f"{num_prfx}{num}"] = df.loc[:,f"{cell_prfx}{num}"].str.replace('\D', '', regex=True)
        
        # combine into range format
        anchr = prfx + "$" + new_df.loc[:,f"{lttr_prfx}1"] + "$" + new_df.loc[:,f"{num_prfx}1"] + ":$" + new_df.loc[:,f"{lttr_prfx}2"] + "$" + new_df.loc[:,f"{num_prfx}2"] 
        
        # convert areas that are not a range to single cell format
        anchr = anchr.str.replace(':$$','',regex=False)
        
        if len(anchr) == 1:
            anchr = anchr.iloc[0]
        elif isinstance(cells,list):
            anchr = anchr.tolist()
        
        return anchr
        
    
    def rtEq(self,rt_dic,fp,lot,sec):
        """Create the equation to calculate the rate"""
        
        # start with the df
        fp_df = fp.df
        
        # retrieve ranges
        rt_eq = '=SUMPRODUCT('
        
        # OLD - weighting against SEPM
        years = fp_df[fp.cal_yr_hdr].drop_duplicates().tolist()
        fp_df['eq'] = np.nan
        fp_df['totl'] = np.nan
        # match the year to the rate dictionary 
        rates = []
        for year in years:
            lot_cells = fp_df[fp_df[fp.cal_yr_hdr] == year].loc[:,lot].tolist()
            lot_rng  = self.cellsToRange(lot_cells)
            range_rate = "( " + rt_dic[year] + " * SUM(" + lot_rng + ") )" 
            rates.append(range_rate)
        # combine the equations into a list and add them to the overall equation
        rt_eq = rt_eq + " + ".join(rates) + ")" 
        sum_cells = fp_df.loc[:,lot].tolist()
        sum_rng = self.cellsToRange(sum_cells)
        totl = f"SUM({sum_rng})"
        
        # divide by the total to get the average rate to apply
        rt_eq = rt_eq + f"/{totl}"
        
        return rt_eq
    
    def assgnCell(self,addr,val,ws=None):
        """Assigns a value to a cell in openpyxl"""
        
        if ws == None:
            ws = self.ws
            
        ws.cell(row=self.opxlRow(addr),column=self.opxlCol(addr)).value = val
        
        return
    
    def opxlRow(self,addr):
        """Returns the row of an address"""
        chars = [*addr]
        
        row = "".join([char if char.isdigit() else "" for char in chars]) 
        
        return int(row)
    
    def opxlCol(self,addr):
        """Returns the row of an address"""
        chars = [*addr]
        
        col = "".join([char if not char.isdigit() else "" for char in chars]) 
        col = self.colToNum(col)
        
        return col
    
    def wsPrfx(self,ws):
        """returns the cell prefix based on the ws name"""
        
        name = ws.title
        
        prfx = f"'{name}'!"
        
        return prfx
    
    def opxlUsedRange(self,ws):
        """Returns the used range in the openpyxl worksheet"""
        # get the minimum row
        min_row = ws.min_row
        while min_row < ws.max_row:
           cells = ws[min_row]
           if all([cell.value is None for cell in cells]):
               min_row += 1
           else:
               break
        
        # get the maximum row
        max_row = ws.max_row
        while max_row > 0:
            cells = ws[max_row]
            if all([cell.value is None for cell in cells]):
                max_row -= 1
            else:
                break
    
        # get the maximum column
        min_col = ws.min_column
        while min_col < ws.max_column:
            cells = next(ws.iter_cols(min_col=min_col, max_col=min_col, max_row=max_row))
            if all([cell.value is None for cell in cells]):
                min_col += 1
            else:
                break
        
        # get the maximum column
        max_col = ws.max_column
        while max_col > 0:
            cells = next(ws.iter_cols(min_col=max_col, max_col=max_col, max_row=max_row))
            if all([cell.value is None for cell in cells]):
                max_col -= 1
            else:
                break
        
        # concatenate into range string
        start_addr = self.numToCol(ws.min_column) + str(ws.min_row)
        end_addr = self.numToCol(max_col) + str(max_row)
        addr_str = f"'{ws.title}'!{start_addr}:{end_addr}"
        
        return addr_str
    
