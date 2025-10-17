import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Your existing custom functions
# from your_script import clean_path, first_time_unique_code_run_pfp, first_time_run_pfp, simple_gba_tab, simple_team_tab

process_tabs = st.tabs(["1st Time Run", "Maintenance"])

with process_tabs[0]:
    tabs = st.tabs(["PFP Processing", "GBA Extraction", "Team Extraction"])

    with tabs[0]:
        st.write("### PFP Processing")

        # Info section
        st.info("""
        📌 **First Run – PFP Processing**

        1. Upload the latest **RAW PFP file** from your respective  
           `CPW FINAL PACKAGE \\ 01 Data Processing \\ Project Financial Plan (PFP)` folder.

        2. Once uploaded, the file will be automatically saved locally.

        3. Then click **Add Unique Code** and **Clean & Save** sequentially to complete processing.
        """)

        # File uploader (replaces manual path)
        uploaded_file = st.file_uploader("📂 Upload RAW PFP Excel File", type=["xlsx", "xls"], key="pfp_file_uploader")

        if uploaded_file is not None:
            # Save the uploaded file to a fixed folder
            save_dir = os.path.join(os.getcwd(), "Uploaded_PFP_Files")
            os.makedirs(save_dir, exist_ok=True)
            saved_path = os.path.join(save_dir, uploaded_file.name)

            with open(saved_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # ✅ Display actual saved path
            st.success(f"✅ File uploaded and saved at:\n`{saved_path}`")

            try:
                df_raw = pd.read_excel(saved_path)
                st.write(f"📊 Raw Data: {df_raw.shape[0]} rows")
                st.dataframe(df_raw.head(3))

                # Add Unique Code button
                if st.button("Add Unique Code", key="create_project_plan_btn"):
                    unique_df = first_time_unique_code_run_pfp(df_raw)

                    project_plan_path = os.path.join(save_dir, "Project Plan Analysis-continuous.xlsx")
                    unique_df.to_excel(project_plan_path, index=False)

                    st.success("✅ Project Plan Analysis created!")
                    st.write("**Project Plan Analysis Preview (with Unique Code):**")
                    st.dataframe(unique_df.head(3))

                    st.session_state["add_unique_clicked"] = True
                    st.session_state["unique_df"] = unique_df

                # Clean & Save step
                if st.session_state.get("add_unique_clicked", False):
                    if st.button("Clean & Save", key="clean_pfp_btn"):
                        unique_df = st.session_state.get("unique_df")
                        cleaned_df = first_time_run_pfp(unique_df)

                        final_date_str = datetime.now().strftime('%Y-%m-%d')
                        cleaned_file_name = f"Project Plan Analysis-continuous-{final_date_str}.xlsx"
                        cleaned_file_path = os.path.join(save_dir, cleaned_file_name)
                        cleaned_df.to_excel(cleaned_file_path, index=False)

                        st.success(f"🧹 Cleaned data saved: `{cleaned_file_name}`")
                        st.dataframe(cleaned_df.head(3))

                        st.info("✅ Ready for GBA and Team extraction!")
                        st.session_state.pop("add_unique_clicked", None)
                        st.session_state.pop("unique_df", None)

            except Exception as e:
                st.error(f"❌ Error processing file: {e}")

    # Other tabs
    with tabs[1]:
        print("a")

    with tabs[2]:
        print("b")
