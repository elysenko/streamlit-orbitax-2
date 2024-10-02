# -*- coding: utf-8 -*-
"""
Created on Wed Dec  8 16:34:12 2021

@author: Eric
"""

import pandas as pd
from datetime import datetime
import numpy as np
import sys
import os
sys.path.append(os.getcwd())
try:
    from backend_modules.excel_reporter import excel_reporter
    from backend_modules.xlFuncs import xlFuncs 
    import backend_modules.package_creator as pc 
except:
    from excel_reporter import excel_reporter
    from xlFuncs import xlFuncs 
    import package_creator as pc 
import datetime as dt
import os
import math
from openpyxl import Workbook

class rprtGenerator(xlFuncs):
    """
    This class handles all report generation
    """
    def __init__(self,codes,dataPckg):
        self.dataPckg = dataPckg
        self.codes = self.cleanCodes(codes.copy())
        self.moDict = {1:"q1",2:"q1",3:"q1",
                        4:"q2",5:"q2",6:"q2",
                        7:"q3",8:"q3",9:"q3",
                        10:"q4",11:"q4",12:"q4"}
        self.qtrStrt = {'q1':1,'q2':4,'q3':7,'q4':10}
        self.mo2daydict = {1:31,2:28,3:31,4:30,5:31,6:30,7:31,8:31,9:30,10:31,11:30,12:31}
        self.flrChar = '-'
        self.inRngHdr = 'In Rng'
        self.bllngHdr = 'Contract Qtrs Billed'
        self.contrRngHdr = 'Contract Range Quarters'
        self.frstYrHdr = 'Is First Yr'
        self.payDueHdr = 'Payment Due'
        self.rprtHdr = 'Qtrly Report Headers'
        self.isEndingHdr = "Is Ending"
        self.clLstQtrs = 3
        self.hdrDic = {"itc": "International Tax Calculator", "pa": "Packages", "beps": "BEPS", "dac6": "DAC6","gmt": "GMT",'icw':"ICW"}
        self.australia_flag = dataPckg['australia']
        
    def cleanCodes(self,codes):
        """
        casts the string types as ints for the Material Code, % ACV Subject to Royalty,
        MAX ROYALTY, and PCKG columns
        """
        codes['Material Code'] = codes['Material Code'].astype(int)
        codes['% ACV Subject to Royalty'] = codes['% ACV Subject to Royalty'].astype(str).str.rstrip('%').astype('float') / 100.0
        codes['MAX ROYALTY'] = codes['MAX ROYALTY'].replace('[\$,]', '', regex=True)
        codes['MAX ROYALTY'] = pd.to_numeric(codes['MAX ROYALTY'], errors="coerce")
        codes['PCKG'] = codes['PCKG'].str.strip()
        codes = codes.replace({'% ACV Subject to Royalty':1},'na')
        codes = codes.fillna("na")
        return codes

    def appendSection(self, report, df, hdr):
        if len(df) > 0:
            # REINDEX THE DATAFRAME
            report.reset_index(inplace=True,drop=True)
            # ADD AN AEMPTY ROW
            report.loc[report.shape[0]] = np.nan # Adds an empty row
            # ADD THE HEADER ROW
            hdr_row = pd.DataFrame({'Client':hdr},index=[0])
            report = pd.concat([report,hdr_row],axis=0)
            # ADD THE DATA
            report = pd.concat([report,df],axis=0)

        return report

    def gen_rep(self,acv_df):
        """
        Called from the flask backend to create the report. pckg includes all the data from
        the web GUI, and df is the data from the uploaded ACV
        """
        # *********************************************
        # CREATE THE CLIENT LIST FOR THIS QUARTER
        
        cl_lst = self.getClientList(acv_df,fltr_col=False)
        # GET ITEMS THAT ARE BILLED THIS QUARTER OR ARE LATE NEW/LATE RENEWAL
        cols = ['Billed This Quarter','Late New','Late Renewal']
        cl_lst = cl_lst[cl_lst[cols].any(axis=1)]
        # DIVIDE THE REPORT BY PACKAGE TYPES
        pckgTypes = ['itc','pa','beps','dac6','gmt','icw']
        rep = pd.DataFrame(columns=cl_lst.columns)
        new_cols = ['Is First Billing','Late New']
        rnwl_cols = ['Late Renewal']
        for tp in pckgTypes:
            # ADD THE SECTIONS
            df = cl_lst[cl_lst['PCKG']==tp].copy()

            new = df[df[new_cols].any(axis=1)]
            new_hdr = self.hdrDic[tp] + " - New:"
            rep = self.appendSection(rep, new, new_hdr)
            rnwl = pd.concat([df[df['Is First Billing'] == False],df[df[rnwl_cols].any(axis=1)]]).sort_values('Client')
            rnwl_hdr = self.hdrDic[tp] + " - Renewals:"
            rep = self.appendSection(rep, rnwl, rnwl_hdr)
        
        # FILTER OUT ROWS AFTER THIS QUARTER
        
        
        # CALCULATE THE NUMBER OF MONTHS BILLED IF LESS THAN 12
        rep = rep.reset_index(drop=True)
        rep['Months Billed (<12)'] = rep[rep['SP Customer Number'].notnull()].apply(lambda row: self.getNumMoBlld(row), axis=1)

        # PUT THE % ROYALTY BASE DUE
        perc_roy = {'itc': {"New":0.35, 'Renewal': 0.3},
                    'pa': {"New":0.35, 'Renewal': 0.3},
                    'beps': {"New":0.35, 'Renewal': 0.3},
                    'dac6': {"New":0.5, 'Renewal': 0.5},
                    'gmt': {"New":0.5, 'Renewal': 0.5},
                    'icw': {"New":0.5, 'Renewal': 0.5},
                    }

        rep['% Royalty Base Due'] = np.nan
        for tp in pckgTypes:
            pckg = rep[rep['PCKG'] == tp]
            new_idx = pckg.loc[pckg[new_cols].any(axis=1),'% Royalty Base Due'].index
            rnwl_idx_1 = pckg.loc[pckg['Is First Billing'] == False,'% Royalty Base Due'].index
            rnwl_idx_2 = pckg.loc[pckg[rnwl_cols].any(axis=1),'% Royalty Base Due'].index
            rep.loc[new_idx,'% Royalty Base Due'] = perc_roy[tp]['New']
            rep.loc[rnwl_idx_1,'% Royalty Base Due'] = perc_roy[tp]['Renewal'] 
            rep.loc[rnwl_idx_2,'% Royalty Base Due'] = perc_roy[tp]['Renewal'] 
        # DROP EXTRA COLUMNS
        extra_cols = ['Ending Within 1 Month','Ending Within 2 Months','Ending Within 3 Months','Billed This Quarter','Is First Billing']
        rep = rep.drop(columns=extra_cols)

        # REMOVE FALSES TO MAKE IT EASIER TO READ
        rep['Late New'] = rep['Late New'].replace(False, np.nan)
        rep['Late Renewal'] = rep['Late Renewal'].replace(False, np.nan)

        # CHANGE THE TRUE TO SAY LATE FOR THE LATE COLS
        rnwl_map = {True:'late'}
        rep['Late New'] = rep['Late New'].map(rnwl_map)
        rep['Late Renewal'] = rep['Late Renewal'].map(rnwl_map)

        # CHANGE % ACV Subject to Royalty FROM DECIMAL TO FULL NUMBER
        # rep['% ACV Subject to Royalty'] = rep['% ACV Subject to Royalty'] * 100

        # ADD THE Payment Due COLUMN AND TOTAL ROW
        rep['Payment Due'] = np.nan
        ttl_row = pd.DataFrame({'Client' : "TOTAL"},index=[0])
        rep.loc[rep.shape[0]] = np.nan # Adds an empty row
        rep = pd.concat([rep,ttl_row],axis=0)

        # ADD XS TO THE FIRST COLUMN FOR FILTERING
        rep = addXCol(rep)

        # REORDER THE DF
        end_cols = ["Late New","Late Renewal"]
        cols = rep.columns
        cols = cols.drop(end_cols).tolist()
        cols = cols + end_cols
        rep = rep.reindex(columns=cols)
        
        # **************** Calculate the Amount Due **************************
        # create df_ref for cell references
        wb = Workbook()
        ws = wb.worksheets[0]
        yr = self.dataPckg['yr']
        qtr = self.dataPckg['qtr']
        sheet_name = 'Royalty Report'
        ws.title = sheet_name
        df_ref = self.wsToCellRef(ws,rep) 
        
        # identify indices for multiplication
        payment_hdr = 'Payment Due'
        str_idcs = rep[rep[payment_hdr].apply(lambda x: isinstance(x, str))].index
        no_na_idcs = rep['% ACV Subject to Royalty'].dropna().index
        total_idx = rep[rep['Client']=='TOTAL'].index
        mult_idcs = list(set(no_na_idcs) - set(str_idcs) - set(total_idx))
        
        # identify the columns that need to be used
        mo_col = df_ref.loc[mult_idcs,'Months Billed (<12)']
        act_col = df_ref.loc[mult_idcs,'$ Actual Royalty Base']
        prc_col = df_ref.loc[mult_idcs,'% Royalty Base Due']
        
        # equation
        rep.loc[mult_idcs,payment_hdr] = ""
        rep.loc[mult_idcs,payment_hdr] = '=IF(' + mo_col + '>0,' + mo_col + '/12*' + act_col + '*' + prc_col + ',' + act_col + '*' + prc_col + ')'
        
        # Calculate Total
        rep.loc[total_idx,payment_hdr] = '=SUM(' + self.cellsToRange(df_ref.loc[mult_idcs,payment_hdr]) + ')'
        # ********************************************************************
        
        # create excel workbook
        er = excel_reporter()
        
        payload = {'sheet_name' : sheet_name, 
                   'header' : f'Royalty Report {qtr}, {yr}',
                   'hide_cols' : ['Filter'],
                   'mult_cols' : {payment_hdr:['% Royalty Base Due','$ Actual Royalty Base','Months Billed (<12)']},
                   'prc_cols' : [col for col in rep.columns if '%' in col],
                   'dllr_cols' : [col for col in rep.columns if '$' in col or payment_hdr == col],
                   'fontcolor_cols' : {(220,20,60):['Late New','Late Renewal']},
                   'mult_idcs' : mult_idcs,
                   }
        wb = er.create_report(rep,**payload)
        
        return wb
    
    def multiply(hdr='',cols=[],op = '*'):
        """Performs the operation on the cells"""
    
    def getNumMoBlld(self, row):
        """
        determines how many months a company should be billed for if their contract does not span a full year
        """
        
        contrStrtDt = row['Contract Start Date']
        contrEndDt = row['Contract End Date']
        contrStrtMo = contrStrtDt[0:2]
        contrEndMo = contrEndDt[0:2]
        contrStrtQtr = self.getStartQuarter(contrStrtMo)
        contrEndQtr = self.getStartQuarter(contrEndMo)
        contrStrtYr = contrStrtDt[6:]
        contrEndYr = contrEndDt[6:]
        contrRng = self.getQtrList(contrStrtQtr,contrStrtYr,contrEndQtr,contrEndYr)
        contrRng = self.detProRata(contrRng,row)
        if len(contrRng) < 4:
            # SEE IF THE END QUARTER IS 1 YEAR AFTER THE
            strtDt = row['Contract Start Date']
            endDt = row['Contract End Date']
            strtDt = datetime.strptime(strtDt, "%m/%d/%Y")
            endDt = datetime.strptime(endDt, "%m/%d/%Y")
            # This is how we round extra months that don't comprise a full year:
            # Take the difference in months.
            month_diff = endDt.year*12 + endDt.month - strtDt.year*12 - strtDt.month
            day_diff = endDt.day - strtDt.day
            # if there's more than a +1 day offset between the end date and the start date, it counts as an extra month
            if day_diff > 1:
                month_diff = month_diff + 1
            return int(month_diff)

        else:
            return np.nan

    def updMatCds(self, acv_df):
        """
        Use to store and update material numbers as they are changed
        """
        #    - import the table from `mat_code_changes.csv` as old_mat_cds
        #    - unpackage the acv table from dataPckg['data']
        #    - in the acv dataframe, filter the column 'Old Material Codes'
        #       to just have the last word (do this in the function 'rfrmCols')
        #       (reference chatGPT or google for 'pandas column get last word' for help)
        #    - filter the acv table for rows where the numbers in 'Old Material Codes' and 'Material Codes' are not the same
        #       if the numbers are not the same, then add the pair to the old_mat_cds dataframe
        #    - save old_mat_cds into 'mat_code_changes.csv'
        #    - add old_mat_cds to the dataPckg dictionary' to use later
        #       (ex. dataPckg['old_mat_cds'] = old_mat_cds)
        
        try:
            filepath = r'previous_mat_codes/mat_code_changes.csv'
            filepath = self.getFldrPath(filepath)
        except:
            filepath = r'./../previous_mat_codes/mat_code_changes.csv'
            filepath = self.getFldrPath(filepath)
        chng = pd.read_csv(filepath)
        # Find instances where the codes do not match. If they don't the code is changing
        try:
            cds = acv_df[["Old Material Code","Material Code"]].copy()
            cds['Old Material Code'] = cds['Old Material Code'].str.split().str[-1]
            cds['Old Material Code'] = cds['Old Material Code'].astype(int)
            cds = cds[cds['Old Material Code'] != cds['Material Code']]
            # Rename "Material Code" as "New Material Code" for easier table interpretation
            cds = cds.rename(columns={'Material Code': 'New Material Code'})
        except:
            cds = pd.DataFrame(columns=["Old Material Code","Material Code"])
        chng = pd.concat([chng,cds])
        chng.drop_duplicates(inplace= True)
        chng.to_csv(filepath,index=False)
        # Keep a copy for later use
        return chng

    def reformatDf(self,df):
        """
        This returns a df of the ACV data. The data initially is not formatted correctly
        (i.e. not left justified and the column headers are not in the top row)
        """
        def rplcHdrs(df):
            """
            isolate where header data is on the excel sheet by deleting headers until "Contract Number" is found
            """

            while not "Contract Number" in df.columns and not 'SP ': #Arbitrary Header to know when the data begins
                new_header = df.iloc[0] #grab the fdfst row for the header
                df = df[1:] #take the data less the header row
                df.columns = new_header
                hdrs = list(df.columns.values)
            return df
        def reformHdrs(df):
            """
            Renames certain headers for clarity
            """
            hdrs2Chnge = {
                    "Sub Material Num (Numeric)" : "Material Code",
                    "Royalty Material codes" : "Old Material Code"
                    }
            df = df.rename(columns=hdrs2Chnge)
            return df
        def rfrmCols(df):
            """
            Get rid of column white space, null columns, or columns with headers
            that are not strings
            """
            df = df.loc[:, df.columns.notnull()]
            df = df.rename(columns=lambda colName: str(colName).strip())
            col = df.columns[0]
            while not isinstance(col,str):
                df = df.drop(columns=col)
                col = df.columns[0]
            return df
        def drpCols(df):
            """
            Removed Columns that are duplicate information and cause problems
            """
            cols2Drp = [
                    ]
            # Older files do not have these columns

            for col in cols2Drp:
                try:
                    df = df.drop(columns=[col])
                except:
                    pass
            return df
        df = drpCols(df)
        df = rplcHdrs(df)
        df = reformHdrs(df)
        df = rfrmCols(df)

        # Some of the data has no material code. We are not interested in these
        df = df[df['Material Code'].notna()]

        # drop duplicate values if they are like the header
        col_hdr = df.columns.tolist()[0]
        df = df[df[col_hdr] != col_hdr]

        # cast as an integer to make sure we got only codes
        df['Material Code'] = df['Material Code'].astype(int)
        return df

    def mrgCds(self,df):
        """
        After cleaning data, we want to merge the data with self.codes to know what
        each data package each entry is based on the material code
        """
        df = self.reformatDf(df)
        to_add = []
        pckgs = self.dataPckg['packages']
        for pckg in pckgs:
            if pckgs[pckg] == 1:
                to_add.append(pckg)
        codes = self.codes[self.codes['PCKG'].isin(to_add)]
        australia_flag = self.dataPckg['australia']
        if australia_flag:
            df.rename(columns={"Sub Material Num (Numeric)":'Material Code'})
        data = pd.merge(df,
                    codes,
                    on ='Material Code',
                    how ='left')
        # print(data)
        data = data.dropna(subset=['PCKG'])
        return data

    def isNaN(self,num):
        """
        returns whether a number is NaN or not
        """
        return num!= num

    def formatDates(self,df):
        """
        Takes dataof the form YYYMMDD and converts it to MM/DD/YYYY. It checks
        two columns - 'MYR Contract Start Date', 'SAP Contract Start Date',
        'MYR Contract End Date', and 'Contract End Date' to see which has the
        Correct dates, and then formats them
        """
        australia_flag = self.dataPckg['australia']
        def formatDateRow(row):
            """
            Isolates the correct date to use based on whether a date is in the
            multi year col or not
            """
            def getDate(row,period):
                """
                Gets the date from either the multi year column or the SAP Date column
                LOGIC:
                If there is a multi year date, use that, otherwise use SAP
                start and end dates
                """
                use_multi = False
                if australia_flag:
                    mlti_yr_strt = np.nan
                    mlti_yr_end = np.nan
                    strt_hdr  = 'Contract Start Date'
                    strt_dt = row[strt_hdr]
                    end_hdr  = 'Contract End Date'
                    end_dt = row[end_hdr]
                else:
                    mlti_yr_strt = np.nan
                    mlti_yr_end = np.nan
                    strt_dt = row['Contract Start Date']
                    end_dt = row['Contract End Date']
                if period == 'start':
                    if self.isNaN(mlti_yr_strt) or use_multi == False:
                        nw_strt_dt = verifDateForm(strt_dt,australia_flag)
                        return nw_strt_dt
                    else:
                        date = mlti_yr_strt
                elif period == 'end':
                    if self.isNaN(mlti_yr_end) or use_multi == False:
                        nw_end_dt = verifDateForm(end_dt,australia_flag)
                        return nw_end_dt
                    else:
                        date = mlti_yr_end
                if not isinstance(date,str):
                    date = str(date)
                yr = date[:4]
                mo = date[4:6]
                day = date[6:8]
                date = mo+'/'+day+'/'+yr
                return date
            row['start date'] = getDate(row,period='start')
            row['end date'] = getDate(row,period='end')
            return row

        def verifDateForm(date,australia=False):
            """
            Makes sure dates are of the form MM/DD/YYYY
            """
            def appElem(lst,elem):
                """
                takes in a list comprising the date, and the next element of the
                date. If the length of the element is 1, add a 0 to the front of
                it so the month and day always have 2 characters.
                """
                if len(elem) == 1:
                    elem = '0' + elem
                lst.append(elem)
                return lst
            if isinstance(date,datetime):
                date = datetime.strftime(date,"%m/%d/%Y")
            if "/" in date:
                elems = date.split("/")
            elif "-" in date:
                elems = date.split("-")
            else:
                print("cannot identify date break in start and end dates")
            nw_elems = []

            # reorder the elements of the date
            if len(elems[0]) <= 2:
                # Appends elements in a certain order depending on the report type
                if australia:
                    nw_elems = appElem(nw_elems,elems[1])
                    nw_elems = appElem(nw_elems,elems[0])
                    nw_elems = appElem(nw_elems,elems[2])
                else:
                    nw_elems = appElem(nw_elems,elems[0])
                    nw_elems = appElem(nw_elems,elems[1])
                    nw_elems = appElem(nw_elems,elems[2])
            else:
                nw_elems = appElem(nw_elems,elems[1])
                nw_elems = appElem(nw_elems,elems[2])
                nw_elems = appElem(nw_elems,elems[0])
            nw_date = ''
            for elem in nw_elems:
                nw_date = nw_date + elem + "/"

            # remove the last '/' on the end
            nw_date = nw_date[:-1]
            return nw_date

        # Gets the correct date and creates Start Date and End Date columns for it
        df = df.apply(lambda row:formatDateRow(row),axis = 1)
        # drop the columns that are not needed
        # df = df.drop(columns = ['SAP Contract Start Date','Contract End Date','Multi Year Date','MYR Contract End Date'])
        return df

    def getRoyBs(self,data):
        """
        adds a column to 'data' that is the 'ACV Subject to Royalty'
        This column tells you how much of the total amount is subject to royalty
        """
        data['ACV Subject to Royalty'] = data.apply(lambda row:self.detAcv(row),axis = 1)
        return data

    def sepInRngData(self,dataPckg):
        """
        Gets data that is billable for the quarter specified by the user
        """
        inRngHdr = dataPckg['inRngHdr']
        data = dataPckg['data']
        dataPckg['masterdata'] = data
        data = data[data[inRngHdr] == True]
        dataPckg['data'] = data

        return dataPckg

    def getClLsts(self):
        """
        gets a list of all clients from the last three quarters
        """
        m = dt.date.today().month
        qtr_now = (m-1)//3 + 1
        yr_now = dt.date.today().year
        files = []
        for i in range(0,4):
            qtr_now, yr_now = self.subtrQtr(qtr_now, yr_now)
            files.append(f"q{qtr_now}_{yr_now}.csv")
        orb_path = r'/home/orbitax/mysite'
        athn_path = r'/home/athenaconsulting/app'
        if os.path.exists(orb_path):
            path = orb_path
        elif os.path.exists(athn_path):
            path = athn_path
        else:
            path = os.getcwd()
        prv_cl_lst = pd.DataFrame()
        for f in files:
            temp = pd.read_csv(f"{path}{f}", low_memory=False, encoding = "ISO-8859-1")
            if prv_cl_lst.empty:
                prv_cl_lst = temp
            else:
                prv_cl_lst = pd.concat([prv_cl_lst,temp], axis=0)
        prv_cl_lst = prv_cl_lst.dropna(subset=['Material Code'])
        prv_cl_lst = prv_cl_lst.drop_duplicates(ignore_index=True)

    def subtrQtr(self,qtr,yr):
        """
        gives the previous quarter. If it is the first, it gives the fourth
        """
        qtr = qtr - 1
        if qtr < 1:
            qtr = 4
            yr = yr - 1
        return qtr,yr

    def detInRng(self,dataPckg):
        """
        detects if a contract's billing quarter is in the specified billing quarter
        """
        data = dataPckg['data']
        data = data.apply(lambda row:self.verifInRngMulti(row=row,dataPckg=dataPckg), axis = 1)
        dataPckg['data'] = data
        return dataPckg

    def num2mon(self,num,commas = False):
        """
        Takes a number and converts it to a string of money with a $
        """
        if commas:
            num = "{:,}".format(num)
        else:
            num=str(num)
        mon="$"+num
        return mon

    def mon2num(self,mon):
        """
        Takes a string of money and returns a float
        """
        if isinstance(mon,str):
            mon = mon.replace("$","")
            mon = mon.replace(",","")
        return float(mon)

    def detAcv(self,row):
        """
        Determines the Royalty Base Amount
        """
        mr = row['MAX ROYALTY']
        rr = row['% ACV Subject to Royalty']
        if isinstance(rr,str):
            if row['% ACV Subject to Royalty'] == "na":
                australia_flag = self.dataPckg['australia']
                self.acv_hdr = 'ACV'
                roy = row[self.acv_hdr]
                return roy
            else:
                rr = float(row['% ACV Subject to Royalty'])
        if isinstance(mr,str):
            mr = float(mr)
        acv = self.mon2num(row['ACV'])
        roy = round(acv * rr,0)
        if roy > mr:
            roy = mr
        return roy
    #
    def getStartQuarter(self,mo):
        """
        Takes in a month and returns the quarter associated with it
        """
        quarters = {1:"q1",2:"q1",3:"q1",
                    4:"q2",5:"q2",6:"q2",
                    7:"q3",8:"q3",9:"q3",
                    10:"q4",11:"q4",12:"q4"}
        strtQrtr = quarters[int(mo)]
        return strtQrtr

    def getBilQrt(self,row,qrtrSel):
        """
        Checks the row to see if it should be billed in the qrtrSel
        """
        strtDt = str(row['start date'])
        strtMo = int(strtDt[0:2])
        blngQrtr = self.getStartQuarter(strtMo)
        return blngQrtr

    def date2Num(self,date):
        """
        Takes a date of the form YYYY/MM/DD and returns and int of YYYYMMDD
        """
        if isinstance(date,str):
            if '/' in date:
                date = date.split('/')
                yr = date[2]
                mo = date[0]
                day = date[1]
                dateNum = int(yr+mo+day)
            date = int(dateNum)
        return date

    def add02dt(self,dt):
        """
        in the dictionary of going between months and quarters, this ensures that the month is two digits instead of one
        """
        if len(str(dt)) == 1:
            dt = "0" + str(dt)
        else:
            dt = str(dt)
        return dt

    def verifInRngMulti(self,row,dataPckg):
        """
        Row contains the data. strtYr/Qtr and endYr/Qtr are the billing times we are looking through
        """
        contrStrtDt = row['start date']
        contrEndDt = row['end date']
        contrStrtMo = contrStrtDt[0:2]
        contrEndMo = contrEndDt[0:2]
        contrStrtQtr = self.getStartQuarter(contrStrtMo)
        contrEndQtr = self.getStartQuarter(contrEndMo)
        contrStrtYr = contrStrtDt[6:]
        contrEndYr = contrEndDt[6:]
        bllngHdr = dataPckg['bllngHdr']
        contrRngHdr = dataPckg['contrRngHdr']
        strtQtr = dataPckg['strtQtr']
        endQtr = dataPckg['endQtr']
        strtYr = dataPckg['strtYr']
        endYr = dataPckg['endYr']
        inRngHdr = dataPckg['inRngHdr']
        isEndingHdr = dataPckg['isEndingHdr']
        contrRng = self.getQtrList(contrStrtQtr,contrStrtYr,contrEndQtr,contrEndYr)
        contrRng = self.detProRata(contrRng,row)
        bllngRng = [i for i in contrRng if contrStrtQtr in i]
        row[bllngHdr] = bllngRng
        row[contrRngHdr] = contrRng
        repRng = self.getQtrList(strtQtr,strtYr,endQtr,endYr)
        inRng = [i for i in bllngRng if i in repRng]
        if len(inRng) >= 1:
            inRngBool = True
        else:
            inRngBool = False
        row[inRngHdr] = inRngBool
        if bllngRng[-1] in repRng:
            isEnding = True
        else:
            isEnding = False
        row[isEndingHdr] = isEnding
        return row

    def list2dt(self,lst):
        """
        takes a list of three numbers[mo,day,yr] and converts it to a string date
        """
        date = ""
        for elem in lst:
            date = date + str(elem) + "/"
        date = date[:-1]
        return date

    def subtrDay(self,date):
        """
        subtracts a day from a date and returns the new date
        """
        yr = date[6:]
        mo = date[0:2]
        day = date[3:5]
        yr = int(yr)
        mo = int(mo)
        day = int(day)
        day = day - 1
        if day < 1:
            mo = mo - 1
            if mo < 1:
                yr = yr - 1
                mo = 12
            day = self.mo2daydict[mo]
        mo = self.add02dt(mo)
        day = self.add02dt(day)
        return str(mo) + "/" + str(day) + "/" + str(yr)

    def add3mo(self,date):
        """
        Takes in a list of quarters and removes the last element until the list comprises yearly increments from the start quarter
        """
        yr = date[6:]
        mo = date[0:2]
        yr = int(yr)
        mo = int(mo)
        mo = mo + 3
        if mo > 12:
            yr = yr + 1
            mo = mo - 12
        day = self.mo2daydict[mo]
        mo = self.add02dt(mo)
        day = self.add02dt(day)
        newDate = mo + "/" + day + "/" + str(yr)
        return  newDate

    def detProRata(self,qtrLst,row):
        """
        determines the quarters that a company should be billed for
        """
        def is3MoPast(row):
            """
            determines if fewer than three months will elapse before a contract ends
            """
            strtDt = row['Contract Start Date']
            projEndDt = row['Contract End Date']
            endYr = row['Contract End Date'][6:]
            projEndDt = self.subtrDay(strtDt)
            projEndDt = projEndDt.split("/")
            projEndDt[2] = endYr
            projEndDt = self.list2dt(projEndDt)
            projEndDtPls3 = self.add3mo(projEndDt)
            projEndDtPls3 = self.date2Num(projEndDtPls3)
            rEndDt = self.date2Num(row['Contract End Date'])
            if projEndDtPls3 < rEndDt:
                # More than 3 months elapsed and the company is pro rated
                return True
            else:
                return False
        strtQtrNum = qtrLst[0][1:2]
        endQtrNum = int(strtQtrNum) - 1
        if endQtrNum == 0:
            endQtrNum = 4
        endQtr = "q" + str(endQtrNum)
        stop = False
        # If it's more than 3 months after the starting of the contract, pro rate the billing. Else, remove that billing segment
        # print(row)
        if (not is3MoPast(row)) and len(qtrLst)>3:
            while stop != True:
                if qtrLst[-1][:2] != endQtr:
                    del qtrLst[-1]
                else:
                    stop = True
        return qtrLst

    def getQtrList(self,strtQtr,strtYr,endQtr,endYr):
        """
        Takes a start year, end year, start quarter, and end quarter, and constructs a list of all quarters that span the period
        """
        def getStrtIter(qtr):
            """
            returns a quarter number using base 0
            """
            strtQtrDict = {"q1":0,"q2":1,"q3":2,"q4":3}
            return strtQtrDict[qtr]
        def getEndIter(qtr):
            """
            converts a quarter into a number
            """
            endQtrDict = {"q1":1,"q2":2,"q3":3,"q4":4}
            return endQtrDict[qtr]
        def frmtOut(qtr,yr):
            """
            returns a string of the quarter and year due
            """
            return qtr + " " + yr + " Due"
        qtrDic = {0:"q1",1:"q2",2:"q3",3:"q4"}
        qtrLst = ["q1","q2","q3","q4"]
        qtrRngLst = []
        yrLst = list(range(int(strtYr),int(endYr)+1))
        for yr in yrLst:
            yr = str(yr)
            if yr == strtYr or yr == endYr:
                if yr == strtYr:
                    # strt and end are arbitrary iterators to go through all of the quarters
                    strt = getStrtIter(strtQtr)
                    if len(yrLst) == 1:
                        end = getEndIter(endQtr)
                    else:
                        end = getEndIter("q4")
                elif yr == endYr:
                    strt = getStrtIter("q1")
                    end = getEndIter(endQtr)
                for i in range(strt,end):
                    qtr = qtrDic[i]
                    qtrRngLst.append(frmtOut(qtr,yr))
            else:
                for qtr in qtrLst:
                    qtrRngLst.append(frmtOut(qtr,yr))
        return qtrRngLst

    def getHdrsFromPerd(self,perd,sngl):
        """
        returns headers for the report depending on if it is a single quarter or multi quarter report
        """
        qtr = str(perd[:2]) + " "
        if sngl:
            qtr = ''
            perDue = '% Royalty Base Due'
        else:
            perDue = f'{qtr}% Base Due'
        numMoBlld = f'{qtr} Months Billed (<12)'
        return perDue, numMoBlld

    def fillQtrCols(self,dataPckg):
        """
        Fills a df with new columns that are in self.qtrCols
        """
        # unpackage the dataPckg. Get the start and end data
        strtQtr = dataPckg['strtQtr']
        endQtr = dataPckg['endQtr']
        strtYr = dataPckg['strtYr']
        endYr = dataPckg['endYr']
        data = dataPckg['data']

        # data = data.apply(lambda row: self.fillDispCols(row),axis = 1)
        self.qtrCols = self.getQtrList(strtQtr, strtYr, endQtr, endYr)
        for col in self.qtrCols:
            data[col] = self.flrChar
        dataPckg['data'] = data
        return dataPckg

    def getBlnkCls(self):
        blnk_cl_lst = pd.DataFrame(columns=['Client','SP Customer Number','Contract Number','Material Code','PCKG','Is First Billing'])
        return blnk_cl_lst

    def getAllCls(self,directory,curr_cl_lst_name):
        """
        given a path to the client lists directory, it extracts all client lists and
        concatenates them into a single dataframe
        """
        dfs = []
        for filename in os.listdir(directory):
            if self.clLstBfr(filename,curr_cl_lst_name):
                f = os.path.join(directory, filename)
                df = pd.read_csv(f)
                qtr = " ".join(filename.split(".")[0].split("_")[1:3])
                df['Qtr'] = qtr
                dfs.append(df)
        if len(dfs) > 0:
            new_df = pd.concat(dfs, axis=0, ignore_index=True).reset_index(drop=True)
            # print(new_df.columns)
            # print(new_df)
            new_df = new_df[new_df['SP Customer Number'].notnull()]
            new_df = new_df.sort_values('Client')
        else:
            new_df = self.getBlnkCls()
        return new_df

    def get4ClAgo(self,directory,curr_cl_lst_name):
        curr_qtr = int(curr_cl_lst_name.split("_")[2][1])
        curr_yr = int(curr_cl_lst_name.split("_")[3].split(".csv")[0])
        # get the list 4 quarters ago
        prev_yr = curr_yr - 1
        filename = self.mkClFlnm(curr_qtr, prev_yr)
        # print('client list directory')
        # print(directory)
        if filename in os.listdir(directory):
            f = os.path.join(directory, filename)
            df = pd.read_csv(f)
        else:
            df = pd.DataFrame(columns=['Client','SP Customer Number','Contract Number','Material Code','PCKG','Is First Billing'])
        return df

    def clLstBfr(self,filename,cl_lst_name):
        """
        checks two client list files and makes sure the one in the folder (filename)
        is in the correct range of client lists needed
        """
        # take apart the names to determine the year and quarter for the
        # folder client list (filename) and the current client list name
        fldr_qtr = int(filename.split("_")[2][1])
        fldr_yr = int(filename.split("_")[3].split(".csv")[0])
        rprt_qtr_num = int(cl_lst_name.split("_")[2][1])
        rprt_yr = int(cl_lst_name.split("_")[3].split(".csv")[0])

        # create month counts for each
        rprt_mo_ct = rprt_yr*12 + 3*rprt_qtr_num
        fldr_mo_ct = fldr_yr*12 + 3*fldr_qtr

        # get the mo_ct for three quarters ago
        num_qtrs_ago = self.clLstQtrs
        cl_strt_qtr = rprt_qtr_num - num_qtrs_ago

        # if the previous quarter is in the prior year, account for that
        if cl_strt_qtr < 1:
            cl_strt_yr = rprt_yr - 1
            cl_strt_qtr = cl_strt_qtr + 4
        else:
            cl_strt_yr = rprt_yr
        cl_strt_mo_ct = cl_strt_yr*12 + 3*cl_strt_qtr

        # gives the range of quarters for the client lists
        cl_window = range(cl_strt_mo_ct, rprt_mo_ct) # range does not include last, so no need to subtract. Does not use client list if it is same date as rprt to generate

        # compare
        return fldr_mo_ct in cl_window

    def getFldrPath(self,fldr=None):
        orb_pth = r'/home/orbitax/mysite'
        athn_pth = r'/home/athenaconsulting/app'

        if os.path.exists(orb_pth):
            path = os.path.join(orb_pth,fldr)
        elif os.path.exists(athn_pth):
            path = os.path.join(athn_pth,fldr)
        else: # Save locally
            path = os.path.join(os.getcwd(),fldr)
        # if the folder doesn't exist, then initialize it
        if not os.path.exists(path):
            os.mkdir(path)

        return path

    def srchClLsts(self,row,cl_lst):
        """
        searches the client lists in the local folder and concatenates them into a single client list to compare against.
        Only uses client lists previous to the one passed
        """
        def allEntriesDifferent(df, col_name):
            num_unique = df[col_name].nunique()
            num_entries = df[col_name].count()
            return num_unique == num_entries

        cl_num = row['SP Customer Number']
        mat_cd = row['Material Code']
        df = cl_lst[cl_lst['SP Customer Number'] == cl_num]
        df = df[df['Material Code'].isin([mat_cd])]

        if len(df) > 0: # means that no client record was found in the last three quarters and it is the client's first year
            if df['Is First Billing'].iloc[0]:
                isFrstYr = True
                endTag = "FrstPerc"
            else:
                isFrstYr = False
                endTag = "ScndPerc"
        else:
            isFrstYr = True
            endTag = "FrstPerc"

        return isFrstYr, endTag

    def withinLastYr(self,qtr,yr):
        """
        determines if a report is for data within the last year
        """
        m = dt.date.today().month
        qtr_now = (m-1)//3 + 1
        yr_now = dt.date.today().year
        qtrs = []
        for i in range(0,4):
            qtr_now, yr_now = self.subtrQtr(qtr_now, yr_now)
            qtrs.append(f"q{qtr_now}_{yr_now}")
        if f"{qtr.lower()}_{yr}" in qtrs:
            return True
        return False

    def rvrseDictSrch(self,dic,val):
        """
        Searches for a key backwards in a dictionary based on the value. Helper Function
        """
        key = list(dic.keys())[list(dic.values()).index(val)]
        return key

    def isFirstYr(self,row,yr,qtr):
        """
        This determines whether or not this is the company's first year contracted with
        Orbitax. Returns a boolean value of True or False
        """
        strtDt = str(row['start date'])
        strtYr = int(strtDt[6:])
        strtMo = int(strtDt[0:2])
        qtrMo = int(self.rvrseDictSrch(self.moDict,qtr))
        frstYr = False
        if not isinstance(yr,int):
            yr=int(yr)
        strtQtr = self.moDict[strtMo]
        strtMo = self.qtrStrt[strtQtr] #Change start date to reflect first billing quarter
        contract_start = 12*strtYr + strtMo
        selected_quarter = 12*yr + qtrMo
        mo_since_contr_strt = selected_quarter - contract_start
        if mo_since_contr_strt < 12:
            frstYr = True
        return frstYr

    def renameHeaders(self,row):
        """
        We ultimately use the df as the foundation for the excel sheet. Here we
        rename columns that are important and set them in the correct order
        """
        row['Contract Start Date'] = row['start date']
        row['Contract End Date'] = row['end date']
        row['$ Max Royalty Base'] = row['MAX ROYALTY']
        return row

    def fillDispCols(self,row):
        """ This function is part calculation part renaming columns of the data.
        It calculates the true royalty base, based on whether there is a maximum
        royalty base. if there is and the calculated royalty base is above that, use the
        max royalty base.
        """

        def detActRoy(baseRoy,max_bs):
            """
            detects the actual royalty owed based on the calculated royalty and max royalty
            """
            if baseRoy > max_bs:
                actRoy = max_bs
            else:
                actRoy = baseRoy
            return actRoy
        # Gets what percentage of the ACV is subject to royalty. If no value displayed, all of it is
        acvBsPerc = row['% ACV Subject to Royalty']
        if acvBsPerc == 'na':
            acvBsPerc = 1
        row['% ACV Subject to Royalty'] = acvBsPerc
        # TR keeps changing the name of this header "SP Customer Name"/"SP Name". This ensures it is found
        row = self.renameHeaders(row)
        acv = self.mon2num(row[self.acv_hdr]) # convert the money string into a number
        row['$ ACV'] = acv
        # calculate the base royalty owed
        baseRoy = acv * acvBsPerc
        row['$ Calc Royalty Base'] = baseRoy
        max_bs = row['MAX ROYALTY']
        # The n character signifies there is no max royalty, and it is whatever the calculated base royalty is
        if isinstance(max_bs,str) and "n" in max_bs.lower():
            actRoy = baseRoy
        else:
            if isinstance(max_bs,str):
               max_bs = int(max_bs)
            # Otherwise the actual royalty must be determined based on the base royalty calculated and the max royalty
            actRoy = detActRoy(baseRoy, max_bs)
        row['$ Actual Royalty Base'] = actRoy
        return row

    def getXlHdrs(self,dataPckg):
        """
        gets the headers that will fill the Excel report
        """
        data = dataPckg['data']
        sngl = dataPckg['sngl']
        payDueHdr = dataPckg['payDueHdr']
        rprtHdrs = dataPckg['rprtHdr']

        def get_rprt_hdrs(data,sngl):
            rprtHdrs = []
            sngl = False
            if len(self.qtrCols) == 1:
                sngl = True
            for perd in self.qtrCols:
                perDue,numMoBlld = self.getHdrsFromPerd(perd,sngl)
                if numMoBlld in data.columns:
                    rprtHdrs.append(numMoBlld)
                if perDue in data.columns:
                    rprtHdrs.append(perDue)
                if perd in data.columns and not sngl:
                    rprtHdrs.append(perd)
            return rprtHdrs

        def getNextXlCol(ltrs):
            """
            gets the next Excel header column
            """
            # If "Z" is the only letter in ltrs, add a letter and make everything 'A'
            retVal = ""
            onlz = True
            for ltr in ltrs:
                if ltr.lower() != 'z':
                    onlz = False
            if onlz:
                for i in range(0,len(ltrs)+1):
                    retVal = "A" + retVal
            else:
                modPrev = True
                for ltr in reversed(ltrs):
                    if modPrev:
                        nxtLtr = chr((ord(ltr.upper())+1 - 65) % 26 + 65)
                        if nxtLtr == "A":
                            modPrev = True
                        else:
                            modPrev = False
                    else:
                        nxtLtr = ltr
                    retVal = nxtLtr + retVal
            return retVal
        hdrs = [
                'Client',
                'SP Customer Number',
                'PCKG',
                'Material Code',
                'Contract Start Date',
                'Contract End Date',
                '$ ACV',
                '% ACV Subject to Royalty',
                '$ Calc Royalty Base',
                '$ Max Royalty Base',
                '$ Actual Royalty Base',
                ]
        rprtHdrs = get_rprt_hdrs(data,sngl)
        for rprtHdr in rprtHdrs:
            hdrs.append(rprtHdr)
        hdrs.append(payDueHdr)
        totalCol=data['Payment Due'].map(lambda x: x.lstrip('*').lstrip('$')).astype(float).round(0)
        total = totalCol.sum()
        total = self.num2mon(total)
        dataPckg['hdrs'] = hdrs
        dataPckg['total'] = total
        return dataPckg

    def groupByPckg(self,dataPckg):
        """
        takes the rows and groups them according to the package they are in
        """
        data = dataPckg['data']
        cols = dataPckg['hdrs']
        df = pd.DataFrame(columns=cols)
        sngl = dataPckg['sngl']
        usdPck = self.dataPckg['packages']
        cds = self.codes
        frstYrHdr = self.frstYrHdr
        for key in usdPck.keys():
            if usdPck[key] == 1: # 1 or 0 are passed from the html checkbox to indicate checked or unchecked
                tempCds = cds[cds['PCKG']==key]['Material Code'].tolist()
                cdData = data[data['Material Code'].isin(tempCds)]
                if len(cdData) > 0:
                    if sngl:
                        frstYr = cdData[cdData[frstYrHdr] == True]
                        aftrFrst =  cdData[cdData[frstYrHdr] == False]
                        frstHdr = key + " - New:"
                        scndHdr = key + " - Renewals:"
                        dataLst = [[frstHdr,frstYr],[scndHdr,aftrFrst]]
                    else:
                        hdr = key + ":"
                        dataLst = [[hdr,cdData]]
                    for stuff in dataLst:
                        hdr = stuff[0]
                        data_slice = stuff[1]
                        data_slice = data_slice[cols]
                        if data_slice.shape[0] > 0:
                            df.loc[df.shape[0]] = "" # Adds an empty row
                            # add the header
                            pckgHdrRow = pd.DataFrame({cols[0]:hdr},index=[0])
                            df = pd.concat([df,pckgHdrRow], axis=0)
                            alphbtzd = data_slice.sort_values("Client")
                            df = pd.concat([df,alphbtzd], ignore_index=True, axis=0)
        dataPckg['data'] = df

        return dataPckg

    def addTotal(self,df,cols):
        """
        Adds the total to the resulting dataframe
        """
        # isolate the payment due column and only include non empty rows
        ttl = df[df[cols[-1]]!=""].dropna()[cols[-1]]
        # use regex to remove the '$' and ',' symbols from the columns
        ttl = ttl.str.replace('[\$,*]', '', regex=True).astype(float)
        ttl = self.num2mon(ttl.sum())
        df.loc[df.shape[0]] = "" # Adds an empty row
        ttlRow = pd.DataFrame({cols[0]:"TOTAL",cols[-1]:ttl},index=[0])
        df = pd.concat([df,ttlRow], ignore_index=True)
        return df

    def formatDataCol(self,dataPckg):
        """
        Prepends a '$' to data in cols with a '$' in the header, a '%' if there is
        a '%' in the header, and rounds to two decimal places if there is a '$' in the header
        """
        data = dataPckg['data']
        data = data.fillna('')
        cols = dataPckg['hdrs']
        newData = pd.DataFrame()
        for hdr in cols:
            if "$" in hdr:
                if not hdr == "$ Max Royalty Base":
                    data[hdr].loc[(data[hdr]!='na')&(data[hdr]!=self.flrChar)&(data[hdr]!="")] = data[hdr].loc[(data[hdr]!='na')&(data[hdr]!=self.flrChar)&(data[hdr]!="")].astype(float).round(0)
                data[hdr].loc[(data[hdr]!='na')&(data[hdr]!=self.flrChar)&(data[hdr]!='')] = "$" + data[hdr].loc[(data[hdr]!='na')&(data[hdr]!=self.flrChar)&(data[hdr]!="")].astype(str)
                data[hdr].loc[data[hdr]=='na'] = 'N/A'
            elif "%" in hdr:
                data[hdr].loc[(data[hdr]=='na')] = self.flrChar
                data[hdr].loc[(data[hdr]!=self.flrChar)&(data[hdr]!="")] = (data[hdr].loc[(data[hdr]!=self.flrChar)&(data[hdr]!="")]* 100).astype(int).astype(str) + " %"
            newData[hdr] = data[hdr]
        dataPckg['data'] = newData
        return dataPckg

    def monToEnd(self,endDate):
        """
        returns the number of months until a given endDate
        """
        from datetime import date
        try:
            endDate = datetime.strptime(endDate, "%m/%d/%Y")
        except:
            endDate = datetime.strptime(endDate, "%d/%m/%Y")
        currDate = date.today()
        moDiff = (endDate.year - currDate.year) * 12 + endDate.month - currDate.month
        if (endDate.day > currDate.day):
            moDiff = moDiff + 1
        return moDiff

    def getEndingMonths(self,row,filename):
        """
        determines if a row's end date is withing 1,2, or 3 months of today's date
        """
        end_date_hdrs = ['Contract End Date','Sub End Date']
        end_date_hdr = self.dateHdr(row,end_date_hdrs)
        endDate = row[end_date_hdr]
        monToEnd = self.monToEnd(endDate)
        
        if monToEnd == 1:
            row["Ending Within 1 Month"] = True
        elif monToEnd == 2:
            row["Ending Within 2 Months"] = True
        elif monToEnd == 3:
            row["Ending Within 3 Months"] = True
        return row
    
    def dateHdr(self,row,hdrs):
        """returns the header of the date someone shoud use"""
        
        for hdr in hdrs:
            date = row[hdr]
            try:
                # makesure it is a date, otherwise pass
                self.monToEnd(date)
                return hdr
            except:
                pass
        
        return False
    
    def normalizeDt(self,date):
        """
        Takes a date of the form MM/DD/YYYY and returns the numeric date
        corresponding to the first day of the quarter it is in
        """
        yr = date.split("/")[2]
        mo = date.split("/")[0]
        qtr = self.moDict[int(mo)]
        mo =  str(self.qtrStrt[qtr.lower()])
        if len(mo) == 1:
            mo = "0" + mo
        day = "01"
        date = int(yr + mo + day)
        return date

    def qtrToDate(self,qtr,yr):
        """
        Takes a quarter and year and returns a date of the form MM/DD/YYYY
        corresponding to the first day of the quarter
        """
        mo =  str(self.qtrStrt[qtr.lower()])
        day = "01"
        yr = str(yr)
        date = mo + "/" + day + "/" + yr
        return date

    def getClInLst(self, cl_lst, cust_num, cd_chngs, mat_cd, contr_num, cl_name = ""):
        """
        Gets the clients in the client list with the corresponding
        customer number, contract number, and material code
        """
        mat_cds = [mat_cd]
        if mat_cd in cd_chngs['New Material Code'].to_list():
            old_mat_dic = cd_chngs.set_index('New Material Code')['Old Material Code'].to_dict()
            mat_cd = old_mat_dic[mat_cd]
            mat_cds.append(mat_cd)

        dfs = []
        for mat_cd in mat_cds:
            df = cl_lst[cl_lst['Material Code'] == mat_cd]
            dfs.append(df)

        df = pd.concat(dfs)
        # df = cl_lst[cl_lst['Material Code'] == mat_cd]

        # Filter on Client Name
        temp = df[df['Client'] == cl_name]
        if len(temp) == 0:
            #  If Client Name not found, filter on Customer Number
            temp = df[df['SP Customer Number'] == cust_num]
        if len(temp) == 0:
            #  If Customer Number not found, filter on Contract Number
            temp = df[df['Contract Number'] == contr_num]

        return temp

    def inClLst(self, cl_lst, cust_num, cd_chngs, mat_cd, contr_num, cust_name):
        """
        Determines if a customer is in the previous client lists based on
        SP Customer Number and Material Code
        """
        df = self.getClInLst(cl_lst,cust_num, cd_chngs, mat_cd, contr_num, cust_name)
        if len(df) > 0:
            return True
        else:
            return False

    def detIfFrstBllng(self,row, cl_lst, cd_chngs, filename):
        """
        determines if a contracts has started or if TR is preemptively
        putting it on the ACV.
        Checks to see if the company is on the three previous client lists.
        Also checks if the start date of the contract is after the current
        billing quarter.
        """
        # FIND THE NEXT QUARTER BASED ON THE FILENAME
        australia = self.dataPckg['australia']
        if australia:
            date = filename.split(".")[0].split(" ")
            yr = date[-2]
            qtr = date[-1]
        else:
            date = filename.split(".")[0].split("_")
            yr = date[1]
            qtr = date[0]


        # VERIFY THE CLIENT IS NOT IN ANY OF THE PREVIOUS RELEVANT ACVs
        cust_num = row['SP Customer Number']
        mat_cd = row['Material Code']
        cust_name = row['Client']
        contr_num = row['Contract Number']
        if str(cust_num) == '1003849882.0':
            poo = 1
        if cust_name == "HUMANA INC.":
            poo= 1
        if self.inClLst(cl_lst, cust_num, cd_chngs, mat_cd, contr_num, cust_name):
            """
            VERIFIES IF THE CLIENT HAS ALREADY BEEN BILLED AT SOME POINT
            CHECK IF ANY PREVIOUS CONTRACT END DATE IS BEFORE THE CURRENT START DATE
            """
            if cust_name == "HUMANA INC.":
                poo = 1
            df = self.getClInLst(cl_lst, cust_num, cd_chngs, mat_cd, contr_num, cust_name)
            end_dates = set(df['Contract End Date'].to_list())
            # IF ANY END DATES ARE BEFORE THE CURRENT START DATE, IT'S A RENEWAL
            alrdy_end = False
            curr_strt = self.normalizeDt(self.qtrToDate(qtr,yr))
            for end_date in end_dates:
               end_date =  self.normalizeDt(end_date)
               if end_date <= curr_strt:
                   alrdy_end = True
            return not alrdy_end
        else:
            return True

    def detIfBlldThsQtr(self,row,filename):
        """
        Determines if a company should be billed this quarter or not.
        If the quarter and year are a match and if the ACV value is greater than
        $0
        """

        australia = self.australia_flag
        # Unpack the year and quarter from the filename
        if australia:
            date = filename.split(".")[0].split(" ")
            yr = date[-2]
            qtr = date[-1]
        else:
            date = filename.split(".")[0].split("_")
            yr = date[1]
            qtr = date[0]

        start = row['start date']
        strt_mo = start[:2]
        strt_yr = start[6:]
        strt_qtr = self.moDict[int(strt_mo)]
        acv = row['$ ACV']
        cust_name = row['Client']
        if cust_name == "SYLVAMO":
            poo = 1
        if strt_yr == yr and strt_qtr.lower() == qtr.lower() and acv > 0:
            return True
        else:
            return False

    def detLateNew(self,row,cl_lst,cl_lst_4, cd_chngs, cl_dir,filename):
        """
        If the billing quarter for the entry is before the report and the
        entry has not been found in the prior four reports, then it is a missed
        NEW payment
        """

        cust_num = row['SP Customer Number']
        mat_cd = row['Material Code']
        contr_num = row['Contract Number']
        cust_name = row['Client']
        if "SYLVAMO" in cust_name:
            poo = 1

        if self.isLate(row, cd_chngs, cl_dir, filename) and not self.inClLst(cl_lst, cust_num, cd_chngs, mat_cd, contr_num, cust_name) and not self.inClLst(cl_lst_4,  cust_num, cd_chngs, mat_cd, contr_num, cust_name):
            return True
        else:
            return False

    def detLate(self, row, cl_lst, cl_lst_4, cd_chngs, cl_dir, filename,df):
        """
        Determines if a company billed outside of this quarter should be considered late
        """
        cust_num = row['SP Customer Number']
        cust_name = row['Client']
        
        if row.name == 379:
            poo = 1
        if row.name == 378:
            poo = 1
        if 'HLB' in cust_name:
            poo =1
        if str(cust_num) == "1005739546.0":
            poo = 1
        if self.isLate(row, cd_chngs, cl_dir, filename):
            return True
        else:
            return False

    def detLateRenewal(self, row, cl_lst, cl_lst_4, cd_chngs, cl_dir, filename):
        """
        If the billing quarter for the entry is before the report and the
        entry has not been found in the prior three reports, but is
        found is the prior fourth quarter, then it is a missed
        RENEWAL payment

        If the customer is not found in the quarter they are supposed to be billed
        nor any quarters after that, then it is a late payment. If the customer
        is found within the last 4 quarters, then it is a late renewal. If they
        are not, they are a late NEW.
        """

        cust_num = row['SP Customer Number']
        mat_cd = row['Material Code']
        contr_num = row['Contract Number']
        cust_name = row['Client']
        if "HLB MANN JUDD" in cust_name:
            poo = 1

        if self.isLate(row, cd_chngs, cl_dir, filename) and (self.inClLst(cl_lst,  cust_num, cd_chngs, mat_cd, contr_num, cust_name) or self.inClLst(cl_lst_4,  cust_num, cd_chngs, mat_cd, contr_num, cust_name)):
            return True
        else:
            return False

    def date2Qtr(self,date):
        """
        Takes a date of the form MM/DD/YYYY and returns qtr and yr
        """
        mo = date[:2]
        yr = date[6:]
        qtr = self.moDict[int(mo)]
        return qtr, yr

    def getNxtQtr(self, qtr,yr):
        """
        Returns the next quarter after the given one
        """
        qtr_num = str(qtr)[-1]
        nxt_qtr_num = int(qtr_num) + 1
        if nxt_qtr_num > 4:
            nxt_qtr_num = 1
            yr = int(yr) + 1
        qtr = "q" + str(nxt_qtr_num)
        yr = str(yr)
        return qtr, yr

    def detClLsts(self,qtr,yr, curr_filename):
        """
        Gets the filenames of all client lists since the given qtr and yr
        """
        filenames = []
        curr_qtr, curr_yr = self.getQtrYrFrmAcv(curr_filename)
        curr_date = self.normalizeDt(self.qtrToDate(curr_qtr,curr_yr))
        filename = self.mkClFlnm(qtr.upper(), yr)
        filenames.append(filename)
        nxt_qtr, nxt_yr = self.getNxtQtr(qtr,yr)
        past_date = self.normalizeDt(self.qtrToDate(nxt_qtr,nxt_yr))
        while past_date < curr_date:
            filename =  self.mkClFlnm(nxt_qtr.upper(), nxt_yr)
            filenames.append(filename)
            nxt_qtr, nxt_yr = self.getNxtQtr(nxt_qtr, nxt_yr)
            past_date = self.normalizeDt(self.qtrToDate(nxt_qtr,nxt_yr))
        return filenames

    def isLate(self, row, cd_chngs, cl_dir, curr_filename):
        """
        If the customer is not found in the quarter they are supposed to be billed
        nor any quarters after that, then it is a late payment.
        """
        strt_date = row['start date']
        qtr, yr = self.date2Qtr(strt_date)
        filenames = self.detClLsts(qtr,yr,curr_filename)
        dfs = []

        cust_name = row['Client']
        if cust_name == "HUMANA INC.":
            poo = 1
        if cust_name == "BRANDTECH INC":
            poo = 1
        if cust_name == "AMERICAN EXPRESS TRAVEL RELATED SER":
            poo = 1

        for filename in filenames:
            cl_path = os.path.join(cl_dir, filename)
            if os.path.exists(cl_path):
                dfs.append(pd.read_csv(cl_path, low_memory=False, encoding = "ISO-8859-1"))
        if len(dfs) > 0:
            cl_lst = pd.concat(dfs)
        else:
            cl_lst = self.getBlnkCls()
        #
        cust_num = row['SP Customer Number']
        mat_cd = row['Material Code']
        cust_name = row['Client']
        contr_num = row['Contract Number']
        if cust_name == "HUMANA INC.":
            poo = 1
        if self.inClLst(cl_lst, cust_num, cd_chngs, mat_cd, contr_num, cust_name):
            return False
        else:
            return True

    def getNextQtr(self,qtr,yr):
        """
        Returns the next quarter and year after the given ones
        """
        qtr_num = int(qtr[-1])
        qtr_num = qtr_num + 1
        if qtr_num > 4:
            yr = int(yr) + 1
            qtr_num = 1
        qtr = "q" + str(qtr_num)
        yr = str(yr)
        return qtr,yr

    def concatRows(self,df):
        # Concatenate rows with the same contract numbers, material code, start date, end date, customer name, and customer number
        contr_nums = df['Contract Number'].to_list()
        duplicates = set([x for x in contr_nums if contr_nums.count(x) > 1])

        for contr_num in duplicates:
            temp = df[df['Contract Number'] == contr_num]
            subset = ['Contract Number', 'Material Code', 'start date', 'end date', 'Client', 'SP Customer Number']
            temp2 = temp.drop_duplicates(subset=subset, keep='first', inplace=True)
            if len(temp2) < len(temp):
                poo = 1


    def getClientList(self,acv_df,wb=None,fltr_col=True):
        """
        This function saves and returns the formatted client list
        """
        australia = self.australia_flag
        print("creating client list")
        # print("acv_df")
        # print(acv_df)
        # scrub data of empty rows
        acv_df = acv_df.dropna(how='all')
        
        # reformat
        acv_df = self.reformatDf(acv_df)
        packages = self.dataPckg['packages']
        filename = self.dataPckg['filename']

        # create the filename for the new ACV based on the old filename
        curr_cl_lst_name = self.mkClFnmFrmAcv(filename)
        # GET THE CODE CHANGES FROM THE ACV
        cd_chngs = self.updMatCds(acv_df)
        # GET THE PATH TO THE CLIENT LISTS BASED ON WHAT MACHINE YOU'RE USING
        cl_dir = self.getFldrPath("client_lists")
        # print('cl_dir: (next)')
        # print(cl_dir)
        # GET THE LAST THREE CLIENT LISTS COMBINED INTO ONE
        cl_lst = self.getAllCls(cl_dir, curr_cl_lst_name)
        # GET THE CLIENT LIST FROM A YEAR AGO
        cl_lst_4 = self.get4ClAgo(cl_dir, curr_cl_lst_name)
        dataPckg = {}
        dataPckg['data'] = acv_df
        dataPckg['hdrs'] = ["Client","SP Customer Number","Contract Number",'Material Code','PCKG','Contract Start Date',
                 'Contract End Date',"$ ACV", "% ACV Subject to Royalty",
                 "$ Calc Royalty Base", "$ Max Royalty Base", "$ Actual Royalty Base",
                 'Ending Within 1 Month','Ending Within 2 Months','Ending Within 3 Months',
                  'Billed This Quarter', 'Is First Billing', 'Late New', 'Late Renewal']
        contact_hdrs = ['SP AM Rep Name','SP OWM AM Rep Name', 'SP CSM Rep Name']
        for hdr in contact_hdrs:
            if hdr in acv_df.columns:
                dataPckg['hdrs'] = dataPckg['hdrs'] + [hdr]
        dataPckg['sngl'] = False
        # GET THE ENTRIES THAT MATCH TO THE MATERIAL CODES SAVED/SPECIFIED
        df = self.mrgCds(acv_df)
        # FORMAT THE DATES CORRECTLY IN THE DATASET
        df = self.formatDates(df)
        # GET THE ROYALTY BASELINE FOR THE ENTRIES
        df = self.getRoyBs(df)
        df = df.apply(lambda row: self.fillDispCols(row),axis = 1)
        df = df.apply(lambda row:self.renameHeaders(row),axis = 1)

        df.loc[:,'Ending Within 1 Month'] = ""
        df.loc[:,'Ending Within 2 Months'] = ""
        df.loc[:,'Ending Within 3 Months'] = ""
        df = df.replace({'$ Max Royalty Base':"na"},"N/A")
        df = df.apply(lambda row:self.getEndingMonths(row,filename), axis = 1)

        # IF THE MO AND YR FOR THE SAP START CORRESPOND TO THE ACV QTR AND YR, MARK TRUE
        df['Billed This Quarter'] = df.apply(lambda row:self.detIfBlldThsQtr(row,filename), axis = 1)
        
        # IF THE ENTRY IS FOUND WITHIN THE LAST THREE QUARTERS AND THE START DATE IS BEFORE THE REPORT DATE, MARK TRUE
        print(df)
        df['Is First Billing'] = df[df['Billed This Quarter'] == True].apply(lambda row:self.detIfFrstBllng(row,cl_lst,cd_chngs,filename), axis = 1)
        
        # SEE IF THE CUSTOMERS ARE BILLED IN THE FUTURE
        df['future'] = df.apply(lambda row: self.isFuturePayment(row),axis=1)
        # DO NOT INCLUDE CLIENTS THAT ARE TO BE CHARGED IN THE FUTURE
        df = df[df['future'] == False]
        
        # SEE IF THE CUSTOMER IS LATE
        if len(df[df['Billed This Quarter'] == False]) >0:
            df['Late'] = df.apply(lambda row:self.detLate(row,cl_lst,cl_lst_4, cd_chngs, cl_dir, filename,df), axis = 1)
        else:
            df['Late'] = False
        if df['Late'].any():
            df['Late Renewal'] = df[df['Late'] == True].apply(lambda row:self.detLateRenewal(row,cl_lst,cl_lst_4, cd_chngs, cl_dir, filename), axis = 1)
            df['Late New'] = df[df['Late'] == True].apply(lambda row:self.detLateNew(row,cl_lst,cl_lst_4, cd_chngs, cl_dir,filename), axis = 1)
        else:
            df['Late Renewal'] = False
            df['Late New'] = False
        dataPckg['data'] = df

        print("grouping by packaging")
        dataPckg = self.groupByPckg(dataPckg)
        
        # remove repeat rows (not needed here)
        df = dataPckg['data']
        df = df.drop_duplicates()
        
        # remove first row, it's empty
        df = df.iloc[1: , :]
        # remove the header rows
        df = df[df['SP Customer Number'].notnull()]
        
        if wb is None:
            self.svClLst(df,filename)
        # add the filter column if needed
        if fltr_col:
            df = addXCol(df)
        
        if not wb is None:
            er = excel_reporter()
            yr = self.dataPckg['yr']
            qtr = self.dataPckg['qtr']
            df = self.hdrToLine(df)
            df.loc[0] = np.nan
            payload = {
                'sheet_name':'Client Report',
                'header': f'Client Report {qtr}, {yr}',
                'hide_cols': ['Filter'],
                'wb':wb,
                'prc_cols': [col for col in df.columns if '%' in col],
                'dllr_cols': [col for col in df.columns if '$' in col ]

                }
            wb = er.create_report(df,**payload)
            return wb
        return df
    
    def isFuturePayment(self,row):
        """Determines whether the customer is billed in the future"""
        
        qtr = self.dataPckg['qtr']
        mo  = int(self.qtrStrt[qtr.lower()])
        yr  = int(self.dataPckg['yr'])
        
        custmr_strt = row['start date']
        custmr_mo   = int(custmr_strt[:2])
        custmr_qtr  = self.moDict[custmr_mo]
        custmr_mo   = int(self.qtrStrt[custmr_qtr.lower()])
        custmr_yr   = int(custmr_strt[6:])
        
        if custmr_mo + 12*custmr_yr > mo + 12*yr:
            return True
        else:
            return False
        
    
    def svClLst(self,df,acv):
        """
        save the client list with the name being the quarter and year
        specified by the user
        """
        orb_pth = r'/home/orbitax/mysite'
        athn_pth = r'/home/athenaconsulting/app'
        fldr = "client_lists"
        qtr, yr = self.acvQtrYr(acv)
        filename = self.mkClFnmFrmAcv(acv)
        if os.path.exists(orb_pth):
            curr_dir = orb_pth
        elif os.path.exists(athn_pth):
            curr_dir = athn_pth
        else: # Save locally
            curr_dir = os.getcwd()
        # if self.withinLastYr(qtr, yr):
        path = os.path.join(curr_dir,fldr)
        # create the folder if it is not created already
        if not os.path.exists(path):
            os.mkdir(path)
        # create the full path with filename and save
        path = os.path.join(path,filename)
        df.to_csv(path, index=False)
        return  True

    def getQtrYrFrmAcv(self,filename):
        """
        Returns the quarter and year of a client list based on its filename
        """
        australia = self.australia_flag
        filename = filename.split(".")[0]
        if australia:
            cl_arr = filename.split(" ")
            qtr = cl_arr[-1]
            yr = cl_arr[-2]
        else:
            cl_arr = filename.split("_")
            qtr = cl_arr[0]
            yr = cl_arr[1]

        return qtr, yr

    def mkClFlnm(self,qtr,yr):
        """
        creates the filename to save for the client list based on the  qtr and yr
        """
        if self.australia_flag:
            filename = f'australia_clients_{qtr}_{yr}.csv'
        else:
            filename = f'domestic_clients_{qtr}_{yr}.csv'
        return filename

    def mkClFnmFrmAcv(self,acv):
        """
        creates the filename to save for the client list based on the ACV filename
        """
        if self.australia_flag:
            fname = acv.split(".")[0]
            qtr = fname.split(" ")[-1]
            yr = fname.split(" ")[-2]

        else:
            qtr = acv.split("_")[0]
            yr = acv.split("_")[1]

        return self.mkClFlnm(qtr,yr)

    def acvQtrYr(self,acv):

        australia = self.dataPckg['australia']
        acv = acv.split(".")[0]
        if australia:
            qtr = acv.split(" ")[-1]
            yr = acv.split(" ")[2]
        else:
            qtr = acv.split("_")[0]
            yr = acv.split("_")[1]

        return qtr, yr
    
    def save_wb(self,wb,filename):
        """Saves the Workbook and returns the path"""
        folder = 'reports'
        path = self.getFldrPath(folder)
        full_path = os.path.join(path,filename)
        wb.save(full_path)
        
        return path
        
def resetClLsts():
    name_dic = [{"qtr":"Q4","yr":"2021"},
                {"qtr":"Q1","yr":"2022"},{"qtr":"Q2","yr":"2022"},
                {"qtr":"Q3","yr":"2022"},{"qtr":"Q4","yr":"2022"},
                ]

    for dic in name_dic:
        qtr = dic['qtr']
        yr = dic['yr']
        # GET THE ACV DF
        filename = f"{qtr}_{yr}_ACV.csv"
        acv_path = rf'C:\Users\Eric\Documents\Python Scripts\ORBITAX\ACVs\{filename}'
        acv_df = pd.read_csv(acv_path, low_memory=False, encoding = "ISO-8859-1")

        # GET THE LIST OF CODES AND THEIR CORRESPONDING ROYALTY AMOUNTS
        codes = pd.read_csv(r'./mat_codes.csv')
        rp = rprtGenerator(codes)
        packages = {'itc':1,'pa':1,'beps':1,'dac6':1,'gmt':1,'icw':1}

        # CREATE THE CLIENT LIST
        rp.getClientList(acv_df,fltr_col=True)

def roundUp(n, decimals=0):
    multiplier = 10**decimals
    rounded_val = math.ceil(n * multiplier) / multiplier
    return rounded_val

def addXCol(df):
    """Adds a column of X's to the first column so a user can filter"""

    df.reset_index(inplace=True,drop=True)
    x_col_data = ['x' for idx in df.index]
    x_col = pd.DataFrame(data=x_col_data,columns=['Filter'])

    new_df = pd.concat([x_col,df],axis=1)

    return new_df



if __name__=='__main__':

    # DO NOT DELETE: Global Vars for reporting
    australia = False

    qtr = 'Q3'
    yr = '2023'
    base_folder = r"C:\Users\Eric\Documents\Python Scripts\ORBITAX\ACVs"
    

    # Australia ACV's
    if australia:
        # base_folder = base_folder +  r"\..\AUSTRALIA ACV - Last 5 Quarters"
        filename = f"Australia ACV {yr} {qtr}.csv"
    else:
        # normal ACV's
        filename = rf"{qtr}_{yr}_ACV.csv"
    xl_filename = filename.split('.')[0] + '.xlsx'
    xl_filepath = rf"./../reports/{xl_filename}"
    acv_path = r"/Users/ericlysenko/Documents/2024/streamlit-orbitax-2/session_state/current_df.csv"
    acv_df = pd.read_csv(acv_path, low_memory=False, encoding = "ISO-8859-1")

    # USE THIS IF YOU WANT TO RUN THE REPORT ON CERTAIN CUSTOMER NUMBERS
    # cstmr_nums = [
    #     1003850981,
    #     1003850266,
    #     1003851080,
    #     1003851420,
    #     1003851659,
    #     1004302229,
    #     1004953147,
    #     1005594815,
    #     1005672406,
    #     1005728746
    #             ]

    # DO NOT DELETE: Testing the main report
    # itc = 1
    # pa = 1
    # beps = 1
    # dac6 = 1
    # gmt=1
    # icw = 1
    # pckgDic = {'itc':itc,'pa':pa,'beps':beps,'dac6':dac6,'gmt':gmt,'icw':icw}
    roy_perc_df = pd.read_csv(r"/Users/ericlysenko/Documents/2024/streamlit-orbitax-2/session_state/royalty_percents.csv")
    pckgDic = pc.get_used_packages(roy_perc_df)
    frstYrDic = pc.get_first_year_perc(roy_perc_df)
    scndYrDic = pc.get_second_year_perc(roy_perc_df)
    dataPckg = {
                'qtr':qtr,'yr':yr,
                'packages': pckgDic,
                'sngl': True, 'test': False,
                "filename": filename,
                'australia':australia,
                }
    for key in frstYrDic.keys():
        dataPckg[key] = frstYrDic[key]
    for key in scndYrDic.keys():
        dataPckg[key] = scndYrDic[key]
    
    try:
        codes = pd.read_csv(r'./mat_codes.csv')
    except:
        codes = pd.read_csv(r'./../session_state/mat_codes.csv')
    rprtr = rprtGenerator(codes,dataPckg)
    wb = rprtr.gen_rep(acv_df)
    moe = 1

    # DO NOT DELETE: Testing the client list
    codes = pd.read_csv(r'./../session_state/mat_codes.csv')
    rp = rprtGenerator(codes,dataPckg)
    packages = {'itc':1,'pa':1,'beps':1,'dac6':1}
    wb = rp.getClientList(acv_df,wb,fltr_col=True)
    wb.save(xl_filepath)
    # resetClLsts()

    """

    [X] partial months
    [X] late to end
    [ ] veritas (23,Q2) - Showed up as new but should have been renewal
    [ ] Gilead - New in Q1 and show up again new in Q2
    [ ] Achimi - New in Q1 and show up again new in Q2

    """