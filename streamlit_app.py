import streamlit as st
import pandas as pd
from backend_modules.rprtGen import rprtGenerator
import backend_modules.session_state_manager as ssm
import backend_modules.package_creator as pc
import backend_modules.error_checking as ec
from datetime import datetime
import chardet
from io import BytesIO
import numpy as np

# configuration
st.set_page_config(layout="wide")

# Set GLobal Vars

if not 'royalty_perc_path' in st.session_state:
    st.session_state.royalty_perc_path = r'./session_state/royalty_percents.csv'

if not 'roy_perc' in st.session_state:
    st.session_state.roy_perc = pd.read_csv(st.session_state.royalty_perc_path)

if not 'mat_codes_path' in st.session_state:
    st.session_state.mat_codes_path = r'./session_state/mat_codes.csv'

if not 'mat_codes' in st.session_state:
    st.session_state.mat_codes = pd.read_csv(st.session_state.mat_codes_path)

if not 'working_df_path' in st.session_state:
    st.session_state.working_df_path = r'./session_state/current_df.csv'

if not 'base_cols' in st.session_state:
    st.session_state.base_cols = ['Client','SP Customer Number','Contract Number','Material Code','ACV','Contract Start Date','Contract End Date','SP AM Rep Name','SP OWM AM Rep Name', 'SP CSM Rep Name']

if not 'upload_key' in st.session_state:
    st.session_state.upload_key = 0

if not 'qtr' in st.session_state:
    st.session_state.qtr = ssm.get_curr_qtr()

if not 'yr' in st.session_state:
    st.session_state.yr = ssm.get_curr_year()

if not 'acv_df' in st.session_state:
    st.session_state.acv_df = ssm.get_curr_acv()

if not 'acv_df_view' in st.session_state:
    st.session_state.acv_df_view = ssm.get_curr_acv()

if not 'report_type' in st.session_state:
    st.session_state.report_type = ssm.get_curr_report_type()
    
if not 'qtrs' in st.session_state:
    st.session_state.qtrs = ["Q1", "Q2", "Q3","Q4"]
    
if not 'yrs' in st.session_state:
    st.session_state.yrs = list(reversed(range(2000,datetime.now().year + 1)))
    
if not 'report_types' in st.session_state:
    st.session_state.report_types = ["Domestic", "Australia"]
# ************************* Functions ****************************************
def clear_data():
    """Clears the working directory"""
    
    df = pd.DataFrame(columns=st.session_state.base_cols)
    
    ssm.write_curr_acv(df)
    
    st.session_state.upload_key += 1
    
    return

def change_state(edited_df,ss_key):
      st.session_state[ss_key]=edited_df
      
      return

## Rename Table headers
def convert_header(df):
    dom_conv_dic = {'SP CUSTOMER NAME': 'Client',
                    'SP Customer Name': 'Client',
                'SAP Contract Start Date':'Contract Start Date',
                'SAP Contract End Date':'Contract End Date',
                'Sub Material Num (Numeric)': 'Material Code',
        }
    aus_conv_dic = {'SP Name':'Client',
                    'Sub Material Num (Numeric)':'Material Code',
                    'Amount in USD':'ACV',
                    'Billing Plan Start Date':'Contract Start Date',
                    'Billing Plan End Date':'Contract End Date',
                    }
    
    df.rename(columns=dom_conv_dic, inplace=True)  
    df.rename(columns=aus_conv_dic, inplace=True)     

    return df       

def set_str_from_filename(filename):
    """Sets the Quarter if found in filename"""
    
    # Remove Suffix
    filename = filename.split('.')[0]
    
    filename_list = filename.split('_')
    
    for qtr in st.session_state.qtrs:
        if str(qtr) in filename_list:
            ssm.write_curr_qtr(qtr)
            st.session_state.qtr = qtr
            break
    
    return

def set_yr_from_filename(filename):
    """Sets the Year if found in filename"""
    
    # Remove Suffix
    filename = filename.split('.')[0]
    
    filename_list = filename.split('_')
    
    print(filename_list)
    for yr in st.session_state.yrs:
        if str(yr) in filename_list:
            ssm.write_curr_year(yr)
            st.session_state.yr = yr
            break
    return

def set_rt_from_filename(filename):
    """Sets the Report Type if found in filename"""
    
    # Remove Suffix
    filename = filename.split('.')[0]
    
    filename_list = filename.split('_')
    
    print(filename_list)
    for rt in st.session_state.report_types:
        if str(rt) in filename_list:
            ssm.write_curr_report_type(rt)
            st.session_state.report_type = rt
            break
    return

## Uploade an ACV
def file_uploader(report_type):
    """Upload an ACV"""
    uploader_key = f"upload_box_{st.session_state.upload_key}"
    
    uploaded_file = st.file_uploader("Upload an ACV (.csv)",type=['csv'],key=uploader_key)
    
    base_df = pd.DataFrame(columns = st.session_state.base_cols)
    
        
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        
        result = chardet.detect(file_bytes)
        
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, encoding=result['encoding'])
        
        # Clean Column Headers
        df.columns = df.columns.str.strip()
        
        # Normalize Headers
        base_df = convert_header(base_df)
        
        # get only columns in the report type header
        for col in df.columns:
            if col in base_df.columns:
                base_df[col] = df[col]
        
        # set year, quater, and report type
        filename = uploaded_file.name
        set_str_from_filename(filename)
        set_yr_from_filename(filename)
        st.session_state.upload_key += 1
        
        # set base df
        ssm.write_curr_acv(base_df)
        
        st.rerun()
    else:
        try:
            base_df = st.session_state.acv_df
        except:
            clear_data()
    
    # set the date cols accordingsly
    base_df['Client'] = base_df['Client'].fillna('').astype(str)    
    base_df['Contract Start Date'] = pd.to_datetime(base_df['Contract Start Date'], errors='coerce')
    base_df['Contract End Date'] = pd.to_datetime(base_df['Contract End Date'], errors='coerce')
    
    return base_df

## Create the report
def create_report(rprtGen,acv_df,filename):
    """Creates a report and checks for missing fields"""
    
    # generate the workbook
    wb = rprtGen.gen_rep(acv_df)
    
    # create a stream object
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    st.download_button(
        label="Click here to download your Excel file",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # write the df
    ssm.write_curr_acv(st.session_state.acv_df)

def is_not_empty(obj):
    """verifies if changes are empty or not"""
    
    # Check for empty lists or empty dictionaries
    if isinstance(obj, list):
        return len(obj) > 0 and any(is_not_empty(item) for item in obj)
    elif isinstance(obj, dict):
        return len(obj) > 0 and any(is_not_empty(value) for value in obj.values())
    return True  # Non-empty objects or other types are considered valid


def incorporate_df_changes(df_key,all_changes_key):
    """updates a df with changes"""
    
    all_changes = st.session_state[all_changes_key]
    df = st.session_state[df_key]
    
    print('all_changes: (next)')
    print(all_changes)
    
    print('df BEFORE: (next)')
    print(df)
    
    df = df.reset_index(drop=True)
    
    if is_not_empty(all_changes):
        for row in all_changes['edited_rows']:
            for col in all_changes['edited_rows'][row].keys():
                change = all_changes['edited_rows'][row][col]
                df.loc[row,col]  = change
                
        for deletions in all_changes['deleted_rows']:
            df.drop(deletions,inplace=True)
        
        if(len(all_changes['added_rows'])) > 0:
            for row in all_changes['added_rows']:
                if '_index' in row.keys():
                    idx = row[list(row.keys())[0]]
                    cols = list(row.keys())[1:]
                else:
                    idx = len(df)
                    cols = list(row.keys())
                for col in cols:
                    val = row[col]
                    df.loc[idx,col] = val 
    
    df = df.reset_index(drop=True)
    print('df AFTER: (next)')
    print(df)
    
    st.session_state[df_key] = df
    
    return 
# ****************************************************************************

# Application Title
st.title("Orbitax Calculator")

# Display the dataframe 
col1,col2,col3 = st.columns([1,1,1])
with col1:
    sub_col1,sub_col2,sub_col3 = st.columns([1,1,1])
    with sub_col1:
        report_type = ssm.get_curr_report_type()
        types = st.session_state.report_types
        report_idx = types.index(report_type)
        report_type = st.radio(
            "Report Type",
            types,
            captions=[
                "TR Domestic",
                "TR Australia",
            ],
            index=report_idx,
            on_change=change_state,args=(report_type,'report_type',)
        )
        ssm.write_curr_report_type(report_type)

    with sub_col2:
        qtr = ssm.get_curr_qtr()
        qtrs = st.session_state.qtrs
        qtr_idx = qtrs.index(qtr)
        qtr = st.selectbox(
            "Report Quarter",
            qtrs,
            index=qtr_idx,
            on_change=change_state,args=(qtr,'qtr',),
            key='qtr_selectbox'
        )
        ssm.write_curr_qtr(qtr)
    with sub_col3:
        yr = ssm.get_curr_year()
        yr_rng = st.session_state.yrs
        yr_idx = yr_rng.index(yr)
        yr = st.selectbox(
            "Report Year",
            yr_rng,
            index=yr_idx,
            on_change=change_state,args=(yr,'yr'),
            key='year_selectbox'
        )
        ssm.write_curr_year(yr)
with col2:
    st.write(f"Report Type: {report_type}")
with col3:
    pckgDic = pc.get_used_packages(st.session_state.roy_perc)
    frstYrDic = pc.get_first_year_perc(st.session_state.roy_perc)
    scndYrDic = pc.get_second_year_perc(st.session_state.roy_perc)
    
    # set the australia flag
    if report_type=="Australia":
        australia_flag = True
    else:
        australia_flag = False
        
    # create filenames for reports
    if australia_flag:
        # base_folder = base_folder +  r"\..\AUSTRALIA ACV - Last 5 Quarters"
        filename = f"Australia RR {qtr} {yr}.csv"
    else:
        # normal ACV's
        filename = rf"Domestic RR {qtr} {yr}.csv"
    dataPckg = {
                'qtr':ssm.get_curr_qtr(),
                'yr':ssm.get_curr_year(),
                'packages': pckgDic,
                
                'sngl': True, 'test': False,
                "filename": filename,
                'australia':australia_flag,
                }
    for key in frstYrDic.keys():
        dataPckg[key] = frstYrDic[key]
        
    for key in scndYrDic.keys():
        dataPckg[key] = scndYrDic[key]
    
    rprtGen = rprtGenerator(st.session_state.mat_codes,dataPckg)
    st.write("")
    
   
col1,col2,col3 = st.columns([1,1,1])
with col1:
    st.session_state.acv_df = file_uploader(report_type)
    st.session_state.acv_df_view = st.session_state.acv_df.copy()
    sort_column = st.selectbox('Select column to sort by:', [None]+st.session_state.acv_df.columns.tolist())
    if not sort_column is None:
        st.write('Table Locked for Editing')
        st.write('Set Filter to None to Make Changes')
with col2:
    st.write("")
    st.button("Clear Data",on_click=clear_data)
    
## Display the table of data
acv_df = st.session_state.acv_df_view
if not sort_column is None:
    acv_df = acv_df.sort_values(by=sort_column, ascending=True)
    table_disabled = False # can be set to true to disable of filter
else:
    acv_df = acv_df.sort_index()
    table_disabled = False
acv_df = st.data_editor(acv_df,
                             num_rows='dynamic',
                             column_config={
                                  "Client": st.column_config.TextColumn(disabled=table_disabled),
                                  "SP Customer Number": st.column_config.NumberColumn( format="%f",disabled=table_disabled),
                                  "Contract Number": st.column_config.NumberColumn(format="%f",disabled=table_disabled),
                                  'ACV': st.column_config.NumberColumn("ACV ($)", format="%f",disabled=table_disabled),
                                  'Material Code':st.column_config.NumberColumn('Material Code', format="%f",disabled=table_disabled),
                                  'Contract Start Date': st.column_config.DateColumn(format='MM/DD/YYYY',disabled=table_disabled),
                                  'Contract End Date': st.column_config.DateColumn(format='MM/DD/YYYY',disabled=table_disabled),
                                  
                             },
                             on_change=incorporate_df_changes, args=('acv_df','acv_df_changes',),
                             hide_index=False,
                             key='acv_df_changes'
                        )

# Create Report Button
acv_errors = ec.acv_error_checking(st.session_state.acv_df) 
roy_perc_errors = ec.roy_perc_error_checking(st.session_state.roy_perc)
mat_code_errors = ec.mat_codes_error_checking(st.session_state.mat_codes)
if len(acv_errors.keys()) + len(roy_perc_errors.keys()) + len(mat_code_errors.keys()) == 0:
    err_disabled = False
else:
    err_disabled = True
st.button("Create Report",on_click=create_report,args=(rprtGen,st.session_state.acv_df,filename,),disabled=err_disabled)

# Error Checking
st.title('Error Checking')
st.write('ACV cols missing data')
# for key in acv_errors.keys():
#     st.markdown(f'##### {key}')
#     st.markdown('###### Indices:')
#     lst = acv_errors[key]
#     s = '' 
#     for i in lst: s += "- " + str(i) + "\n"
#     st.markdown(s)
st.write(acv_errors)
st.write('Royalty Percent cols missing data')
st.write(roy_perc_errors)
st.write('Material Code cols missing data')
st.write(mat_code_errors)

