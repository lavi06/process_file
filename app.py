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



def add_filter():
    st.session_state.num_filter += 1


def generate_files():

    df = st.session_state.df
    num_filter = st.session_state.num_filter

    i = 1
    while i <= num_filter:

        filters = st.session_state[f"File-{i}"]
        filename = st.session_state[f"Filename-{i}"]
        if filename.endswith(".xlsx"):
            pass
        else:
            filename = filename + ".xlsx"

        sub_df = df[df["State-County"].isin(filters)]
        sub_df = sub_df.drop("State-County", axis = 1)

        sub_df.to_excel(filename, index = False)
        # with pd.ExcelWriter(filename, datetime_format='YYYY-MM-DD') as writer:
        #     sub_df.to_excel(writer, index=False)

        ####
        wb = load_workbook(filename)
        ws = wb.active

        header_font = Font(name='Calibri', size=11, bold=False)
        header_alignment = Alignment(horizontal='left', vertical='center')

        # Apply formatting to the header row
        for cell in ws[1]:  # Header is in the first row
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = None  # No border

        wb.save(filename)
        ####
        
        i += 1


    _zip = zipfile.ZipFile("Processed.zip", "w", zipfile.ZIP_DEFLATED)

    i = 1
    while i <= num_filter:

        filename = st.session_state[f"Filename-{i}"]
        if filename.endswith(".xlsx"):
            pass
        else:
            filename = filename + ".xlsx"
        _zip.write(filename)

        i += 1

    _zip.close()    

    with open("Processed.zip", "rb") as zipped:
        encoded_string = zipped.read()

    st.session_state.zipped = encoded_string




@st.cache_data
def read_excel(uploaded_file):
    df = pd.read_excel(uploaded_file)

    df["State-County"] = df["Filing_County"] + "-" + df["Filing_State"]

    column_type_mapping = {}

    for column_name, data_type in df.dtypes.items():
        st.write(column_name, ":", data_type)
        
        column_type_mapping[column_name] = str(data_type)

        if "date" in str(data_type):

            df[column_name] = pd.to_datetime(df[column_name], errors='coerce')

            df[column_name] = df[column_name].dt.strftime('%m/%d/%Y')

    return df




if "num_filter" not in st.session_state:
    st.session_state.num_filter = 1

if "df" not in st.session_state:
    df = pd.DataFrame([],columns=["State-County"])

    st.session_state.df = df
    st.session_state.filters = list(df["State-County"].unique())



uploaded_file = st.file_uploader("Choose a .csv/.xlsx file", type = ["csv", "xlsx"])


if uploaded_file is not None:

    # df = pd.read_excel(uploaded_file)
    df = read_excel(uploaded_file)

    cols_required = ["Filing_County", "Filing_State"] 
    cols_not_found = []
    for col in cols_required:
        if col not in df.columns:
            cols_not_found.append(col)

    if len(cols_not_found) > 0:
        errors = {"Columns Not Found" : cols_not_found}

        st.error(json.dumps(errors, indent = 3)) 


    else:
        # df["State-County"] = df["Filing_County"] + "-" + df["Filing_State"]

        st.session_state.df = df
        st.session_state.filters = list(df["State-County"].unique())


        add_row = st.sidebar.button("Add Filter", on_click = add_filter)

        i = 1

        while i <= st.session_state.num_filter:

            st.sidebar.multiselect(f"File {i}", options = st.session_state.filters, default = [], key = f"File-{i}")
            st.sidebar.text_input("Random", value = f"File-{i}.xlsx", placeholder = f"File {i}.xlsx" , label_visibility="collapsed", key = f"Filename-{i}")


            i = i + 1


        zipped_file = st.sidebar.button("Process Files", key='but_generate', on_click = generate_files)

        if zipped_file:
            download = st.sidebar.download_button(label="Download Files", data = st.session_state.zipped, file_name= "Processed.zip", mime='application/octet-stream')


        st.write("Data Preview")
        st.write(df.drop("State-County", axis = 1))


