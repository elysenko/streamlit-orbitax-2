#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Sep 30 14:07:20 2024

@author: ericlysenko
"""

import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# configuration
st.set_page_config(layout="wide")

# Set GLobals
if not 'royalty_perc_path' in st.session_state:
    st.session_state.royalty_perc_path = r'./session_state/royalty_percents.csv'

st.title("Material Codes")

try:
    df = pd.read_csv(st.session_state.royalty_perc_path)
except:
    data = {"Package Select":['ITC','Packages','BEPS','DAC6','GMT','ICW'],
            "Include?": [True, True, True, True, True, True],
            'First Year Royalty (%)':[35,35,35,50,50,50],
            'Renewal Royalty (%)':[30,30,30,50,50,50],
            }
    df = pd.DataFrame(data)

def clean_df(df):
    """takes a df for royalty percent values and formats it correctly"""
    
    return df

df = clean_df(df)

# sort the df
col1,col2,col3 = st.columns([1,1,1])
with col1:
    sort_column = st.selectbox('Select column to sort by:', df.columns)
    
df_sorted = df.sort_values(by=sort_column, ascending=False)

# Display the Data
def change_state(edited_df):
      st.session_state['roy_perc']=edited_df
      
      return

def save_mat_codes(df,mat_codes_path):
    """Saves the material codes"""
    df.to_csv(mat_codes_path,index=False)
    
    # mycode = "<script>alert('Codes Saved')</script>"
    # components.html(mycode, height=0, width=0)
    
    return

st.session_state.roy_perc = df_sorted
df_sorted = st.data_editor(df_sorted,
               hide_index=True,
               column_config={
                    "Include?": st.column_config.CheckboxColumn(
                        help="Select which codes to include in report",
                        default=False,
                    ),
                    'First Year Royalty (%)':st.column_config.NumberColumn(
                        format='%f %%'
                    ),
                    'Renewal Royalty (%)':st.column_config.NumberColumn(
                        format='%f %%'
                    ),
                },
               on_change=change_state, args=(df_sorted,)
            )

save_mat_codes(df_sorted,st.session_state.royalty_perc_path)