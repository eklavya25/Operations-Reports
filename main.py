#############################################################
import streamlit as st

st.set_page_config(page_title="Operations Workbench", layout="wide")

st.title("Operations Workbench")
st.write("Use the button below to go directly to the report page.")

# When user clicks the button, switch to another page
if st.button("➡ Go to WIP Hourly Report"):
    st.switch_page("pages/wip-hourly.py")





