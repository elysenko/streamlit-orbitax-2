#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct  2 13:29:41 2024

@author: ericlysenko
"""

def acv_error_checking(df):
    """Checks that all ACV fields are correctly formatted before creating a report"""
    
    bad_dic = df_check(df)
    
    
    if len(df[df['Client']=='']) > 0:
        bad_dic['Client'] = df[df['Client']==''].index.tolist()
    
    return bad_dic
    
def mat_codes_error_checking(df):
    """Checks that all Material Code fields are correctly formatted before creating a report"""
    
    bad_dic = df_check(df)
    
    return bad_dic

def roy_perc_error_checking(df):
    """Checks that all Material Code fields are correctly formatted before creating a report"""
    
    bad_dic = df_check(df)
    
    return bad_dic

def df_check(df):
    
    bad_dic = {}
    
    for col  in df.columns:
        na_df = df[df[col].isna()]
        if len(na_df) > 0:
            bad_dic[col] = na_df.index.tolist()
    
    return bad_dic

