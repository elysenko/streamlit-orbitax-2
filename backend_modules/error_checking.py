#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct  2 13:29:41 2024

@author: ericlysenko
"""

def acv_error_checking(df):
    """Checks that all ACV fields are correctly formatted before creating a report"""
    
    bad_cols = df_check(df)
    
    if len(df[df['Client']=='']) > 0:
        bad_cols.append('Client')
    
    return bad_cols
    
def mat_codes_error_checking(df):
    """Checks that all Material Code fields are correctly formatted before creating a report"""
    
    bad_cols = df_check(df)
    
    return bad_cols

def roy_perc_error_checking(df):
    """Checks that all Material Code fields are correctly formatted before creating a report"""
    
    bad_cols = df_check(df)
    
    return bad_cols

def df_check(df):
    
    bad_cols = []
    
    for col  in df.columns:
        if df[col].isna().any():
            bad_cols.append(col)
    
    return bad_cols

