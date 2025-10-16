# import streamlit as st

# st.set_page_config(page_title="Operations Workbench", layout="wide")

# st.title("Operations Workbench")
# st.write("Use the buttons below to Generator Reports.")

# # Navigation button
# if st.button("➡ Go to WIP Hourly Report"):
#     st.switch_page("pages/wip-hourly.py")  






import streamlit as st

st.set_page_config(page_title="Operations Workbench", layout="wide")

st.title("Operations Workbench")
st.write("Use the buttons below to generate reports.")

# Navigation button
if st.button("➡ Go to WIP Hourly Report"):
    st.switch_page("wip-hourly")  # <-- Use just the page name (no .py, no folders)



