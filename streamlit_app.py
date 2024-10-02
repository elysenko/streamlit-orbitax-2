import streamlit as st
import pandas as pd
from backend_modules.rprtGen import rprtGenerator
import backend_modules.session_state_manager as ssm
import backend_modules.package_creator as pc
import backend_modules.error_checking as ec
from datetime import datetime
import chardet
from io import BytesIO

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

if not 'report_type' in st.session_state:
    st.session_state.report_type = ssm.get_curr_report_type()

# ************************* Functions ****************************************
def clear_data():
    """Clears the working directory"""
    
    df = pd.DataFrame(columns=st.session_state.base_cols)
    
    ssm.write_curr_acv(df)
    
    st.session_state.upload_key += 1
    
    return

## Display the table of data
def save_working_df(df):
    """Saves the working df you are working on"""
    
    ssm.write_curr_acv(df)
    
    return

def change_state(edited_df,ss_key):
      st.session_state[ss_key]=edited_df
      
      return

## Rename Table headers
def convert_header(df):
    dom_conv_dic = {'SP CUSTOMER NAME': 'Client',
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
    else:
        try:
            base_df = ssm.get_curr_acv()
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
    

# ****************************************************************************

# Application Title
st.title("Orbitax Calculator")

# Display the dataframe 
col1,col2,col3 = st.columns([1,1,1])
with col1:
    sub_col1,sub_col2,sub_col3 = st.columns([1,1,1])
    with sub_col1:
        report_type = ssm.get_curr_report_type()
        types = ["Domestic", "Australia"]
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
        qtrs = ["Q1", "Q2", "Q3","Q4"]
        qtr_idx = qtrs.index(qtr)
        qtr = st.selectbox(
            "Report Quarter",
            qtrs,
            index=qtr_idx,
            on_change=change_state,args=(qtr,'qtr',)
        )
        ssm.write_curr_qtr(qtr)
    with sub_col3:
        yr = ssm.get_curr_year()
        yr_rng = list(reversed(range(2000,datetime.now().year + 1)))
        yr_idx = yr_rng.index(yr)
        yr = st.selectbox(
            "Report Year",
            yr_rng,
            index=yr_idx,
            on_change=change_state,args=(yr,'yr')
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
        filename = f"Australia ACV {yr} {qtr}.csv"
    else:
        # normal ACV's
        filename = rf"{qtr}_{yr}_ACV.csv"
    dataPckg = {
                'qtr':qtr,
                'yr':yr,
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
    sort_column = st.selectbox('Select column to sort by:', [None]+st.session_state.acv_df.columns.tolist())
    if not sort_column is None:
        st.write('Table Locked for Editing')
        st.write('Set Filter to None to Make Changes')
with col2:
    st.write("")
    st.button("Clear Data",on_click=clear_data)
    
## Display the table of data
acv_df = st.session_state.acv_df
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
                             on_change=change_state, args=(acv_df,'acv_df',),
                             hide_index=True,
                        )

ssm.write_curr_acv(acv_df)
st.session_state.acv_df = acv_df

# Erro checking
st.title('Error Checking')
acv_errors = ec.acv_error_checking(st.session_state.acv_df) 
roy_perc_errors = ec.roy_perc_error_checking(st.session_state.roy_perc)
mat_code_errors = ec.mat_codes_error_checking(st.session_state.mat_codes)
if len(acv_errors + roy_perc_errors + mat_code_errors) == 0:
    err_disabled = False
else:
    err_disabled = True
    
st.button("Create Report",on_click=create_report,args=(rprtGen,ssm.get_curr_acv(),filename,),disabled=err_disabled)
st.write('ACV cols missing data')
st.write(acv_errors)
st.write('Royalty Percent cols missing data')
st.write(roy_perc_errors)
st.write('Material Code cols missing data')
st.write(mat_code_errors)

