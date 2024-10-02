#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Oct  1 12:32:39 2024

@author: ericlysenko
"""

import pandas as pd

def perc_col_val(df_perc,pckg,col):
    """Regtrieves the value for a given column in the royalty percent df"""
    
    df = df_perc.copy()
    
    df = df[df['Package Select']==pckg]
    
    val = df[col].iloc[0]
    
    return val

def get_used_packages(roy_perc_df):
    """Returns a dictionary specifying if each package will be used or not"""
    
    moe = 1
    
    itc  = 1 if perc_col_val(roy_perc_df,'ITC','Include?') else 0
    pa   = 1 if perc_col_val(roy_perc_df,'Packages','Include?') else 0
    beps = 1 if perc_col_val(roy_perc_df,'BEPS','Include?') else 0
    dac6 = 1 if perc_col_val(roy_perc_df,'DAC6','Include?') else 0
    gmt  = 1 if perc_col_val(roy_perc_df,'GMT','Include?') else 0
    icw  = 1 if perc_col_val(roy_perc_df,'ICW','Include?') else 0
    
    packages = {'itc':itc,'pa':pa,'beps':beps,'dac6':dac6,'gmt':gmt,'icw':icw}
    
    return packages

def get_first_year_perc(roy_perc_df):
    """Returns the first year percent for each package"""
    
    itc = perc_col_val(roy_perc_df,'ITC','First Year Royalty (%)')
    pa = perc_col_val(roy_perc_df,'Packages','First Year Royalty (%)')
    beps = perc_col_val(roy_perc_df,'BEPS','First Year Royalty (%)')
    dac6 = perc_col_val(roy_perc_df,'DAC6','First Year Royalty (%)')
    gmt = perc_col_val(roy_perc_df,'GMT','First Year Royalty (%)')
    icw = perc_col_val(roy_perc_df,'ICW','First Year Royalty (%)')
    
    packages = {'itcFrstPerc':itc,'paFrstPerc':pa,'bepsFrstPerc':beps,'dac6FrstPerc':dac6,'gmtFrstPerc':gmt,'icwFrstPerc':icw}
    
    return packages
    
def get_second_year_perc(roy_perc_df):
    """Returns the first year percent for each package"""
    
    itc = perc_col_val(roy_perc_df,'ITC','Renewal Royalty (%)')
    pa = perc_col_val(roy_perc_df,'Packages','Renewal Royalty (%)')
    beps = perc_col_val(roy_perc_df,'BEPS','Renewal Royalty (%)')
    dac6 = perc_col_val(roy_perc_df,'DAC6','Renewal Royalty (%)')
    gmt = perc_col_val(roy_perc_df,'GMT','Renewal Royalty (%)')
    icw = perc_col_val(roy_perc_df,'ICW','Renewal Royalty (%)')
    
    packages = {'itcScndPerc':itc,'paScndPerc':pa,'bepsScndPerc':beps,'dac6ScndPerc':dac6,'gmtScndPerc':gmt,'icwScndPerc':icw}
    
    return packages

if __name__ == '__main__':
    """TO debug features"""
    
    roy_perc_path = r'./../royalty_percents.csv'
    
    roy_perc_df = pd.read_csv(roy_perc_path)
    
    usedPckgs = get_used_packages(roy_perc_df)