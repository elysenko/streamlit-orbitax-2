import streamlit as st
import pandas as pd
from backend_modules.rprtGen import rprtGenerator

st.title("Orbitax Calculator")


required_cols = ['Client','SP Customer Number','Contract Number','Material Code','SP AM Rep Name','SP OWM AM Rep Name', 'SP CSM Rep Name']

def file_uploader():
    uploader_key = 1
    
    uploaded_file = st.file_uploader("Upload an ACV (.csv)",type=['csv'],key=uploader_key)
    
    if uploaded_file is not None:
        df = pd.read_csv(uploaded_file)
        
        base_df = df[[col for col in df.columns if col in required_cols]]
    else:
        base_df = pd.DataFrame(columns = required_cols)
    
    return base_df

base_df = file_uploader()

col1,col2,col3 = st.columns([80,1,1])
with col1:
    input_table = st.data_editor(base_df,
                             hide_index=True,
                             num_rows='dynamic')

