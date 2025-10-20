# wip_report_app_fixed.py
import streamlit as st
import pandas as pd
import numpy as np
import re

st.set_page_config(page_title="WIP Report Generator", layout="wide")

# ----------------------------
# Helper functions
# ----------------------------
def calculate_units(group, day_col, hour_col):
    total = 0
    for day, g in group.groupby(day_col):
        g = g.copy()
        # ensure hour_col is integer series (if empty will be zeros)
        g[hour_col] = g[hour_col].fillna(0).astype(int)

        if day in ['Fri', 'Mon']:
            base_cutoff = 18
            cutoff_index = None
            # look for 18,19,...,23 (first occurrence wins)
            for h in range(base_cutoff, 24):
                cutoff_rows = g[g[hour_col] == h]
                if not cutoff_rows.empty:
                    cutoff_index = cutoff_rows.index[0]
                    break
            if cutoff_index is not None:
                total += g.loc[g.index < cutoff_index, '#Units'].sum()

        elif day in ['Tue', 'Wed', 'Thu', 'Sat']:
            base_cutoff = 17
            cutoff_index = None
            # look for 17,18,...,23 (first occurrence wins)
            for h in range(base_cutoff, 24):
                cutoff_rows = g[g[hour_col] == h]
                if not cutoff_rows.empty:
                    cutoff_index = cutoff_rows.index[0]
                    break
            if cutoff_index is not None:
                total += g.loc[g.index < cutoff_index, '#Units'].sum()

        elif day == 'Sun':
            total += g['#Units'].sum()

    return total


def load_wip_file(file):
    if file.name.endswith(".xls") or file.name.endswith(".xlsx"):
        try:
            df_list = pd.read_html(file)
            return df_list[0]
        except ValueError:
            return pd.read_excel(file)
    else:
        st.error("Please upload a valid .xls or .xlsx file")
        return None

def safe_num(x):
    """Return a safe numeric (0 if NaN/None). Convert numpy scalar to python type."""
    if x is None or pd.isna(x):
        return 0
    if isinstance(x, (np.generic,)):
        return x.item()
    try:
        # prefer ints when possible
        xf = float(x)
        if xf.is_integer():
            return int(xf)
        return xf
    except:
        return x

# ----------------------------
# UI
# ----------------------------
st.title("WIP-Hourly Report Generator")
st.write("Upload your `.xls` / `.xlsx` file")

uploaded_file = st.file_uploader("Choose a WIP file", type=["xls", "xlsx"])

if uploaded_file is not None:
    df = load_wip_file(uploaded_file)
    if df is None:
        st.stop()

    # basic validation
    if df.shape[0] < 8:
        st.error("Uploaded file doesn't look like the expected WIP export (too few rows). Check the file and retry.")
        st.stop()

    # slice 
    df1 = df.iloc[6:-1].reset_index(drop=True)

    if df1.shape[0] == 0:
        st.error("No rows found after slicing. Check the input file format.")
        st.stop()

    # set header row
    df1.columns = df1.iloc[0]
    df1 = df1.drop(df1.index[0]).reset_index(drop=True)

    # ==== CRITICAL FIX: sanitize column names (replace NaN with safe strings) ====
    df1.columns = [str(c).strip() if pd.notna(c) else "" for c in df1.columns]

    # convert datetimes (keep as datetime NaT for .dt accessor)
    if "Case Uploaded Time" in df1.columns:
        df1["Case Uploaded Time"] = pd.to_datetime(df1["Case Uploaded Time"], errors="coerce")
    else:
        df1["Case Uploaded Time"] = pd.NaT

    if "Restart Date" in df1.columns:
        df1["Restart Date"] = pd.to_datetime(df1["Restart Date"], errors="coerce")
    else:
        df1["Restart Date"] = pd.NaT

    # remove rows with 'test' in Skill Level (same as yours)
    if "Skill Level" in df1.columns:
        df1 = df1[~df1["Skill Level"].astype(str).str.contains("test", case=False, na=False)].reset_index(drop=True)

    # Make sure #Units column exists and is numeric (fill NaN -> 0)
    if "#Units" in df1.columns:
        df1["#Units"] = pd.to_numeric(df1["#Units"], errors="coerce").fillna(0)
    else:
        df1["#Units"] = 0

    # ----------------------------
    # iOS QC WIP
    # ----------------------------
    df_IOSQC_WIP = df1[df1.get('Pending For', '').isin(['IOS QC', 'Scan QC', 'Hold by Scan QC'])].copy()
    df_IOSQC_WIP['#Units'] = pd.to_numeric(df_IOSQC_WIP['#Units'], errors='coerce').fillna(0)

    # create day/hour columns safely (use .dt while dtype is datetime)
    df_IOSQC_WIP['Restart Date_Day'] = df_IOSQC_WIP['Restart Date'].dt.strftime('%a').fillna('')
    df_IOSQC_WIP['Restart Date_Hour'] = df_IOSQC_WIP['Restart Date'].dt.hour.fillna(0).astype(int)
    df_IOSQC_WIP['Case Uploaded_Day'] = df_IOSQC_WIP['Case Uploaded Time'].dt.strftime('%a').fillna('')
    df_IOSQC_WIP['Case Uploaded_Hour'] = df_IOSQC_WIP['Case Uploaded Time'].dt.hour.fillna(0).astype(int)

    df_restart_valid = df_IOSQC_WIP[df_IOSQC_WIP['Restart Date'].notna()].sort_values(by='Restart Date').reset_index(drop=True)
    total_units_restart = calculate_units(df_restart_valid, 'Restart Date_Day', 'Restart Date_Hour')

    df_restart_blank = df_IOSQC_WIP[df_IOSQC_WIP['Restart Date'].isna()].sort_values(by='Case Uploaded Time').reset_index(drop=True)
    total_units_case_uploaded = calculate_units(df_restart_blank, 'Case Uploaded_Day', 'Case Uploaded_Hour')

    total_IOSQC_WIP_Previous = total_units_restart + total_units_case_uploaded
    total_units_IOSQC = df_IOSQC_WIP["#Units"].sum()
    total_IOSQC_WIP_Advance = total_units_IOSQC - total_IOSQC_WIP_Previous

    # ----------------------------
    # Exocad WIP
    # ----------------------------
    exo_filter = ['Exocad(1)', 'Exocad(2)', 'Exocad(3)', 'Exocad(4)', 'Exocad(5)', 'Exocad(6)', 'Exocad(8)']
    df_ExoCad = df1[df1.get('Skill Level', '').isin(exo_filter)].copy()
    df_ExoCad = df_ExoCad[~df_ExoCad['Pending For'].isin(['IOS QC', 'Scan QC','Hold by Scan QC'])]

    
    # List of IDs to check
    Appliance_list = [
    "L00140", "D00156", "D00251", "D00280", "D00272", "D00190", "D00881", "D00119", "D01053",
    "L00796", "L00799", "D01058", "D01056", "D01057", "D01219", "A00796", "A00799", "M00799",
    "L0251", "A01058", "A01053", "M01056", "A01056", "M01186", "M00119", "M01219", "AJAY251",
    "L01426", "A01426", "N00119", "L01218", "EMLA185", "ELA185", "MLA186E", "LA181M", "LA136",
    "SB1426", "MA181M", "MLA186", "L00156", "L00251", "L00190", "MD001", "L00272", "L00280",
    "M00251", "L00441", "M00280", "A00156", "A00190", "L00448", "M00448", "M00441", "L00730",
    "M0021", "A00881", "M0050", "J00730", "M00881", "A00119", "M01053", "M01057", "ABH280",
    "M01058", "A01057", "M00796", "L01186", "L01192", "M01192", "L01205", "EA001", "L01219",
    "M01205", "SBM1426", "LA187", "L01647", "M01647", "LA136A", "LA185M", "LA181EM+", "MLA187",
    "LA127", "MLA184", "LA181EM", "LA184", "LA183", "MLA127", "LA186EM", "LA185EM", "LA128",
    "MLA128", "MLA183", "LA186M", "A01218", "LA182", "MLA182"
    ]

    Appliance_pattern = '|'.join(map(re.escape, Appliance_list))
    df_ExoCad = df_ExoCad[~df_ExoCad['Hold By'].str.contains(Appliance_pattern, na=False)].copy()


    df_ExoCad['#Units'] = pd.to_numeric(df_ExoCad['#Units'], errors='coerce').fillna(0)

    df_ExoCad_Previous = df1[df1.get('Skill Level', '').isin(exo_filter)].copy()
    df_ExoCad_Previous = df_ExoCad_Previous[~df_ExoCad_Previous['Pending For'].isin(['IOS QC', 'Scan QC'])]
    df_ExoCad_Previous = df_ExoCad_Previous[~df_ExoCad_Previous['Destination'].isin(['In-House'])]
    df_ExoCad_Previous['#Units'] = pd.to_numeric(df_ExoCad_Previous['#Units'], errors='coerce').fillna(0)

    df_ExoCad_Previous['Restart Date_Day'] = df_ExoCad_Previous['Restart Date'].dt.strftime('%a').fillna('')
    df_ExoCad_Previous['Restart Date_Hour'] = df_ExoCad_Previous['Restart Date'].dt.hour.fillna(0).astype(int)
    df_ExoCad_Previous['Case Uploaded_Day'] = df_ExoCad_Previous['Case Uploaded Time'].dt.strftime('%a').fillna('')
    df_ExoCad_Previous['Case Uploaded_Hour'] = df_ExoCad_Previous['Case Uploaded Time'].dt.hour.fillna(0).astype(int)

    df_restart_valid = df_ExoCad_Previous[df_ExoCad_Previous['Restart Date'].notna()].sort_values(by='Restart Date').reset_index(drop=True)
    total_units_restart = calculate_units(df_restart_valid, 'Restart Date_Day', 'Restart Date_Hour')

    df_restart_blank = df_ExoCad_Previous[df_ExoCad_Previous['Restart Date'].isna()].sort_values(by='Case Uploaded Time').reset_index(drop=True)
    total_units_case_uploaded = calculate_units(df_restart_blank, 'Case Uploaded_Day', 'Case Uploaded_Hour')

    total_ExoCad_WIP_Previous = total_units_restart + total_units_case_uploaded
    total_units_ExoCad = df_ExoCad["#Units"].sum()
    total_ExoCad_WIP_Advance = total_units_ExoCad - total_ExoCad_WIP_Previous

    # ----------------------------
    # 3Shape WIP
    # ----------------------------
    shape_filter = ['3Shape(1)', '3Shape(2)', '3Shape(3)', '3Shape(4)', '3Shape(5)', '3Shape(6)', '3Shape(8)']
    df_3Shape = df1[df1.get('Skill Level', '').isin(shape_filter)].copy()
    df_3Shape = df_3Shape[~df_3Shape['Pending For'].isin(['IOS QC', 'Scan QC', 'Hold by Scan QC'])]
    df_3Shape = df_3Shape[~df_3Shape['Hold By'].str.contains(Appliance_pattern, na=False)].copy()
    df_3Shape['#Units'] = pd.to_numeric(df_3Shape['#Units'], errors='coerce').fillna(0)
    total_units_3Shape = df_3Shape["#Units"].sum()

    # ----------------------------
    # Appliance WIP 
    # ----------------------------
    # df_App = df1[df1.get("Case Type", "").isin(["M","M+"])].copy()


    # df_App = df1[df1['Hold By'].str.contains(Appliance_pattern, na=False)].copy()
    # #df_App = df_App.drop_duplicates(subset='Order No', keep='first').copy()



    # df_App = df_App[~df_App.get("Destination", "").isin(["Easydent Dental Lab"])]
    # df_App['#Units'] = pd.to_numeric(df_App['#Units'], errors='coerce').fillna(0)
    # df_App["#Units"] = df_App["#Units"].apply(lambda x: 1 if x < 20 else 2)
    # total_units_App = df_App["#Units"].sum()



    df_App = df1[
        df1.get("Case Type", "").isin(["M","M+"]) | 
        df1['Hold By'].str.contains(Appliance_pattern, na=False)
        ].copy()

    # Exclude certain Destination
    df_App = df_App[~df_App.get("Destination", "").isin(["Easydent Dental Lab"])]

    df_App = df_App.drop_duplicates(subset='Order No', keep='first').copy()

    # Convert #Units to numeric and apply logic
    df_App['#Units'] = pd.to_numeric(df_App['#Units'], errors='coerce').fillna(0)
    df_App['#Units'] = df_App['#Units'].apply(lambda x: 1 if x <= 20 else 2)

    # Calculate total units
    total_units_App = df_App['#Units'].sum()





    # ----------------------------
    # EDDL WIP
    # ----------------------------
    df_EDDL = df1[df1.get('Destination', '').isin(['Easydent Dental Lab'])].copy()
    df_EDDL = df_EDDL[~df_EDDL['Pending For'].isin(['IOS QC', 'Scan QC', 'Hold by Scan QC'])]

    df_EDDL.loc[df_EDDL['Skill Level'].isin(['3Shape(7)', 'Exocad(7)']), '#Units'] = df_EDDL.loc[df_EDDL['Skill Level'].isin(['3Shape(7)', 'Exocad(7)']), '#Units'].apply(lambda x: 1 if x < 20 else 2)
  

    # Function to update #Units
    def update_units(row):
        if pd.isna(row['#Units']):
            return row['#Units']  # Keep NaN as is
        if re.search(Appliance_pattern, str(row['Hold By'])):
            return 1 if row['#Units'] <= 20 else 2
        return row['#Units']

    # Apply function row-wise
    df_EDDL['#Units'] = df_EDDL.apply(update_units, axis=1)



    df_EDDL['#Units'] = pd.to_numeric(df_EDDL['#Units'], errors='coerce').fillna(0)

    df_EDDL['Restart Date_Day'] = df_EDDL['Restart Date'].dt.strftime('%a').fillna('')
    df_EDDL['Restart Date_Hour'] = df_EDDL['Restart Date'].dt.hour.fillna(0).astype(int)
    df_EDDL['Case Uploaded_Day'] = df_EDDL['Case Uploaded Time'].dt.strftime('%a').fillna('')
    df_EDDL['Case Uploaded_Hour'] = df_EDDL['Case Uploaded Time'].dt.hour.fillna(0).astype(int)

    df_restart_valid = df_EDDL[df_EDDL['Restart Date'].notna()].sort_values(by='Restart Date').reset_index(drop=True)
    total_units_restart = calculate_units(df_restart_valid, 'Restart Date_Day', 'Restart Date_Hour')

    df_restart_blank = df_EDDL[df_EDDL['Restart Date'].isna()].sort_values(by='Case Uploaded Time').reset_index(drop=True)
    total_units_case_uploaded = calculate_units(df_restart_blank, 'Case Uploaded_Day', 'Case Uploaded_Hour')

    total_EDDL_WIP_Previous = total_units_restart + total_units_case_uploaded
    total_units_EDDL = df_EDDL["#Units"].sum()
    total_EDDL_WIP_Advance = total_units_EDDL - total_EDDL_WIP_Previous

    # ------------- Display (use safe_num to avoid NaN in JSON) -------------
    st.success("✅ Report Generated Successfully!")

    col1, col2, col3 = st.columns(3)
    col1.metric("iOS QC WIP Previous", safe_num(total_IOSQC_WIP_Previous))
    col1.metric("iOS QC WIP Advance", safe_num(total_IOSQC_WIP_Advance))
    col2.metric("ExoCad WIP Previous", safe_num(total_ExoCad_WIP_Previous))
    col2.metric("ExoCad WIP Advance", safe_num(total_ExoCad_WIP_Advance))
    col3.metric("EDDL WIP Previous", safe_num(total_EDDL_WIP_Previous))
    col3.metric("EDDL WIP Advance", safe_num(total_EDDL_WIP_Advance))

    st.divider()

    col5, col6, col7 = st.columns(3)
    col5.metric("3Shape Total Units", safe_num(total_units_3Shape))
    col6.metric("Appliance Units", safe_num(total_units_App))
    col7.metric("", safe_num(""))
    

    st.divider()

    # # Show sample of the sanitized dataframe (no NaN names)
    # st.subheader("Preview (first 10 rows)")
    # st.dataframe(df1.head(10))

else:
    st.info("📥 Please upload your Excel file to start.")






#####################################################################

