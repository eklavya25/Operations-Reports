import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Lab Summary Report", layout="wide")

st.title("📊 Lab Summary Report Generator")

# Step 1: Upload file
uploaded_file = st.file_uploader("📂 Upload the Booked Data file (.xls or .xlsx)", type=["xls", "xlsx"])

# Step 2: Date input
user_date_input = st.text_input("📅 Enter cutoff date (DD/MM)", placeholder="e.g. 20/10")

if uploaded_file and user_date_input:
    try:
        cutoff_date = datetime.strptime(user_date_input, "%d/%m").replace(year=2025)
        cutoff_time = datetime.strptime("05:30", "%H:%M").time()

        df = pd.read_html(uploaded_file)
        df = df[0]
        df1 = df.iloc[6:-1].reset_index(drop=True)
        df1.columns = df1.iloc[0]
        df1 = df1.drop(df1.index[0])

        df1['#Units'] = pd.to_numeric(df1['#Units'], errors='coerce').fillna(0)
        df1.loc[df1['#Units'] == 0, '#Units'] = 1
        df1.loc[df1['Case Type'].isin(['M', 'M+']), '#Units'] = (
            df1.loc[df1['Case Type'].isin(['M', 'M+']), '#Units']
            .apply(lambda x: 1 if x < 20 else 2)
        )

        df1['Order ID'] = df1['Order ID'].astype(str).apply(
            lambda x: re.search(r'"(\d+)"', x).group(1) if re.search(r'"(\d+)"', x) else x
        )
        df1['Order ID'] = pd.to_numeric(df1['Order ID'], errors='ignore')

        Redesign = df1[df1['Order ID'].astype(str).str.contains('r', case=False, na=False)]
        Redesign_count = len(Redesign)
        Redesign_sum = Redesign['#Units'].sum()

        Restarted = df1[df1['Restart Date'].notna() & (df1['Restart Date'].astype(str).str.strip() != '')]
        Restarted_count = len(Restarted)
        Restarted_sum = Restarted['#Units'].sum()

        df1 = df1[~df1.index.isin(Redesign.index)]

        df1['Case In'] = pd.to_datetime(df1['Case In'], errors='coerce')
        df1 = df1.sort_values(by='Case In', ascending=True).reset_index(drop=True)
        df1['date'] = df1['Case In'].dt.strftime('%d-%m')
        df1['time'] = df1['Case In'].dt.strftime('%H:%M')

        df1['date_dt'] = pd.to_datetime(df1['date'] + '-2025', format='%d-%m-%Y', errors='coerce')
        df1['time_dt'] = pd.to_datetime(df1['time'], format='%H:%M', errors='coerce').dt.time

        df1 = df1[
            (df1['date_dt'] > cutoff_date) |
            ((df1['date_dt'] == cutoff_date) & (df1['time_dt'] >= cutoff_time))
        ].reset_index(drop=True)

        df1 = df1.drop(columns=['date_dt', 'time_dt'])

        summary_df = (
            df1.groupby('Lab Name', as_index=False)
            .agg(
                First_Order_ID=('Order ID', 'min'),
                Last_Order_ID=('Order ID', 'max'),
                Count=('Order ID', 'count'),
                Sum=('#Units', 'sum')
            )
        )

        hold_sum = (
            df1[df1['Hold'] == 'Y']
            .groupby('Lab Name')['#Units']
            .sum()
            .rename('Hold')
        )

        cancel_sum = (
            df1[df1['Cancel'] == 'Y']
            .groupby('Lab Name')['#Units']
            .sum()
            .rename('Cancel')
        )

        summary_df = summary_df.merge(hold_sum, on='Lab Name', how='left')
        summary_df = summary_df.merge(cancel_sum, on='Lab Name', how='left')
        summary_df[['Hold', 'Cancel']] = summary_df[['Hold', 'Cancel']].fillna(0)

        st.subheader("📄 Summary Preview")
        st.dataframe(summary_df, use_container_width=True)

        st.write(f"**Redesign count:** {Redesign_count} | **Sum:** {Redesign_sum}")
        st.write(f"**Restart count:** {Restarted_count} | **Sum:** {Restarted_sum}")

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='Summary')
            writer.save()

        st.download_button(
            label="📥 Download Summary Excel",
            data=buffer.getvalue(),
            file_name="Lab_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")

else:
    st.info("👆 Please upload your `.xls` file and enter a cutoff date (DD/MM) to begin.")
