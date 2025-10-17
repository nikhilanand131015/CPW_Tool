import streamlit as st
import pandas as pd

st.title("Step 1: Process Current Week Raw Data")

uploaded_file = st.file_uploader("Upload Current Week Raw Data (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("File loaded successfully!")
    st.dataframe(df.head())
else:
    st.info("Please upload the Excel file to proceed.")
