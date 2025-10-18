#############################################################
import streamlit as st

st.set_page_config(page_title="Operation's Workbench", layout="wide")

st.title("Operation's Workbench")
st.write("Use the button below to go directly to the report page.")

# When user clicks the button, switch to another page
if st.button("➡ WIP Hourly"):
    st.switch_page("pages/wip-hourly.py")

if st.button("➡ Daily Units Report"):
    st.switch_page("pages/daily-units.py")




