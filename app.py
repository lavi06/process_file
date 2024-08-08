import streamlit as st
import json
import pandas as pd
import time
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

#####################
#### VERSION 1.0 ####
#####################


################################################################################

st.set_page_config(page_title="Process File", page_icon=":bar_chart:")
# st.image("Logo.png", caption=None, width=250, use_column_width=None, clamp=False, channels="RGB", output_format="auto")





def convert_datetime_to_str(series):
    return series
    
    def safe_strftime(x):
        try:
            if pd.isna(x):
                return None
            else:
                return x.strftime('%m/%d/%Y')
        except Exception as e:
            # Log the error if needed, for debugging purposes
            print(f"Error converting value {x}: {e}")
            return x  # Return the original value in case of an error
    
    return series.apply(safe_strftime)


# @st.cache_data
def read_excel(uploaded_file):
    df = pd.read_excel(uploaded_file)

    df["State-County"] = df["Filing_County"] + "-" + df["Filing_State"]

    column_type_mapping = {}

    for column_name, data_type in df.dtypes.items():
        
        column_type_mapping[column_name] = str(data_type)

        if "date" in str(data_type):
            # df[column_name] = pd.to_datetime(df[column_name], errors='coerce')
            # df[column_name] = df[column_name].dt.strftime('%m/%d/%Y')
            df[column_name] = convert_datetime_to_str(df[column_name])

    return df


uploaded_file = st.file_uploader("Choose a .csv/.xlsx file", type = ["csv", "xlsx"])

if uploaded_file is not None:

    # df = pd.read_excel(uploaded_file)
    df = read_excel(uploaded_file)

    df.to_excel("Test.xlsx", index = False)
    with open("Test.xlsx", "rb") as zipped:
        encoded_string = zipped.read()
    
    download = st.sidebar.download_button(label="Download Files", data = encoded_string, file_name= "Test.xlsx", mime='application/octet-stream')

    st.write("Data Preview")
    st.write(df)


