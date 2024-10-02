#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct  2 12:35:18 2024

@author: ericlysenko
"""
import pandas as pd
import os
from datetime import datetime
 
def get_curr_qtr(folder='./session_state',filename='quarter.csv'):
    """Returns the current quarter of the user state"""
    
    
    try:
        # read tge dataframe
        df = read_data(folder,filename)
        
        # get the quarter
        qtr = df.loc[0,'qtr']
    except:
        # default to Q1
        qtr = 'Q1'
        
    return qtr
    
def get_curr_year(folder='./session_state',filename='year.csv'):
    """Returns the current year of the session state"""
    
    try:
        df = read_data(folder,filename)
        
        year = df['year'].iloc[0]
    except:
        year = datetime.now().year
        
    return year
    
def get_curr_acv(folder='./session_state',filename='current_df.csv'):
    """Returns the ucrrent acv of the session state"""
    
    acv = read_data(folder,filename)
    
    return acv

def get_curr_report_type(folder='./session_state',filename='report_type.csv'):
    """Returns the current report type"""
    
    try:
        # read tge dataframe
        df = read_data(folder,filename)
        
        # get the report type
        report_type = df.loc[0,'report_type']
    except:
        # default to Domestic
        report_type = 'Domestic'
        
    return report_type
    

def write_curr_qtr(qtr,folder='./session_state',filename='quarter.csv'):
    """Returns the current quarter of the user state"""
    
    data = {"qtr":qtr}
    
    write_data(data,folder,filename)
    
    return
    
def write_curr_year(year,folder='./session_state',filename='year.csv'):
    """Returns the current year of the session state"""
    
    data = {"year":year}
    
    write_data(data,folder,filename)
    
    return
    
def write_curr_acv(df,folder='./session_state',filename='current_df.csv'):
    """Returns the ucrrent acv of the session state"""
    
    write_data(df,folder,filename)
    
    return

def write_curr_report_type(report_type,folder='./session_state',filename='report_type.csv'):
    """Returns the current quarter of the user state"""
    
    data = {"report_type":report_type}
    
    write_data(data,folder,filename)
    
    return
    
def write_data(data,folder,filename):
    """Writes the dataframe to the location"""
    
    if isinstance(data,dict):
        df = pd.DataFrame(data,index=range(len(data.keys())))
    else:
        df = data
    
    full_path = os.path.join(folder,filename)
    
    # make the directory if it does not exist
    if not os.path.isdir(folder):
        os.mkdir(folder)
    
    df.to_csv(full_path)
    
    return

def read_data(folder,filename):
    """Reads the dataframe from the folder"""
    
    full_path = os.path.join(folder,filename)
    
    df = pd.read_csv(full_path,index_col='Index')
    
    return df