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
if not 'mat_codes_path' in st.session_state:
    st.session_state.mat_codes_path = r'./session_state/mat_codes.csv'

st.title("Material Codes")

df = pd.read_csv(st.session_state.mat_codes_path)

def clean_df(df):
    """takes a df for material values and formats it correctly"""
    df['% ACV Subject to Royalty'] = df['% ACV Subject to Royalty'].astype(str).str.replace('%','').astype(float)
    
    # configure the max royalty allotment
    df['MAX ROYALTY'] = df['MAX ROYALTY'].astype(str).str.replace('$',"")
    df['MAX ROYALTY'] = df['MAX ROYALTY'].str.replace(',','')
    df['MAX ROYALTY'] = df['MAX ROYALTY'].str.replace('na','0')
    df['MAX ROYALTY'] = df['MAX ROYALTY'].astype(float)
    
    return df

df = clean_df(df)

# sort the df
col1,col2,col3 = st.columns([1,1,1])
with col1:
    sort_column = st.selectbox('Select column to sort by:', df.columns)
df_sorted = df.sort_values(by=sort_column, ascending=False)

st.session_state.mat_codes = st.data_editor(df_sorted,
               num_rows='dynamic',
               hide_index=True,
               column_config={
                    "Material Code": st.column_config.NumberColumn(
                        format="%f"
                    ),
                    '% ACV Subject to Royalty':st.column_config.NumberColumn(
                        format='%f %%'
                    ),
                    'MAX ROYALTY':'Max Royalty ($$)',
                    'PCKG':st.column_config.SelectboxColumn(
                        "Package Selection",
                        help="The category of the app",
                        width="medium",
                        options=[
                            "itc",
                            "icw",
                            "gmt",
                            "beps",
                            "pa",
                        ],
                        )
                }
            )

def save_mat_codes(df,mat_codes_path):
    """Saves the material codes"""
    df.to_csv(mat_codes_path,index=False)
    
    mycode = "<script>alert('Codes Saved')</script>"
    components.html(mycode, height=0, width=0)
    
    return
    
st.button("Save Material Codes", on_click=save_mat_codes,args=(st.session_state.mat_codes,st.session_state.mat_codes_path,))
