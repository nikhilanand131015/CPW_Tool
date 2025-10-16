import streamlit as st
import os
import pandas as pd
import io
from datetime import datetime, date
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File
import xlwings as xw
import time

# === Global state ===
selected_file: str = ""
gba_file_path: str = ""
team_file_path: str = ""
wb = None
ws = None
last_row: int = 0
last_col: int = 0
var_start_row: int = 2

load_dotenv()

# === PFP Functions ===
def first_time_unique_code_run_pfp(df):
    df['Unique Code'] = df['Project Number'].astype(str) + ' - ' + df['Employee Name'].astype(str)
    cols = ['Unique Code'] + [col for col in df.columns if col != 'Unique Code']
    return df[cols]

def first_time_run_pfp(df):
    # CHANGE: Track cleaning statistics for detailed reporting
    original_count = len(df)
    
    # Remove duplicates based on Unique Code
    df_unique = df.drop_duplicates(subset=['Unique Code'], keep='first')
    duplicates_removed = original_count - len(df_unique)
    
    # Remove rows with missing Employee Name
    df_no_missing = df_unique.dropna(subset=['Employee Name'])
    blank_employees_removed = len(df_unique) - len(df_no_missing)
    
    # Remove 'Labor Cost, Conversion Employee' entries
    df_final = df_no_missing[df_no_missing['Employee Name'] != 'Labor Cost, Conversion Employee']
    labor_cost_removed = len(df_no_missing) - len(df_final)
    
    # CHANGE: Store cleaning statistics in session state for display
    st.session_state["cleaning_stats"] = {
        "original_count": original_count,
        "duplicates_removed": duplicates_removed,
        "blank_employees_removed": blank_employees_removed,
        "labor_cost_removed": labor_cost_removed,
        "final_count": len(df_final)
    }
    
    return df_final

# === Utils Functions ===
def clean_path(path: str) -> str:
    if not path:
        return ""
    return path.strip().strip('"').strip("'")

def clean_file_name(fn: str) -> str:
    """Clean filename by removing invalid characters"""
    invalid = r'\/:*?"<>|'
    for ch in invalid:
        fn = fn.replace(ch, '_')
    return fn.strip()

def derive_gba_file_path(selected_file: str) -> str:
    target_folder = "01 Data Processing"
    pos = selected_file.find(target_folder)
    if pos == -1:
        raise ValueError(f"Target folder '{target_folder}' not found in path.")
    directory_path = selected_file[: pos + len(target_folder)]
    if len(directory_path) - 19 > 0:
        gba_file_path = directory_path[: len(directory_path) - 19]
    else:
        gba_file_path = os.path.dirname(directory_path)
    return gba_file_path

def derive_team_file_path(selected_file: str) -> str:
    target_folder = "02 GBA Workbooks"
    pos = selected_file.find(target_folder)
    if pos == -1:
        raise ValueError(f"Target folder '{target_folder}' not found in path.")
    directory_path = selected_file[: pos + len(target_folder)]
    if len(directory_path) - 17 > 0:
        team_file_path = directory_path[: len(directory_path) - 17]
    else:
        team_file_path = os.path.dirname(directory_path)
    return team_file_path

def open_workbook(path: str):
    return xw.Book(path)

def find_last_row(sheet):
    return sheet.api.Cells(sheet.api.Rows.Count, 1).End(-4162).Row

def find_last_col(sheet):
    return sheet.api.Cells(1, sheet.api.Columns.Count).End(-4159).Column

def find_column_index_from_headers(headers, column_name):
    for j, h in enumerate(headers, start=1):
        if (h or "").strip() == column_name:
            return j
    return 0

def find_first_empty_row_in_col(sheet, col=5, start=5, search_limit=5000):
    """
    Finds the first empty row in a specified column within a given range.
    
    Args:
        sheet: Excel worksheet object
        col: Column number to check (default: 5)  
        start: Starting row to check from (default: 4)
        search_limit: Maximum row to search up to (default: 5000)
    
    Returns:
        int: Row number of first empty cell, or next available row after last used row
    """
    for r in range(start, search_limit + 1):
        val = sheet.range((r, col)).value
        if val is None or (str(val).strip() == ""):
            return r
    lr = sheet.api.Cells(sheet.api.Rows.Count, col).End(-4162).Row
    return max(start, lr + 1)

def read_block(sheet, nrows, ncols):
    rng = sheet.range((1, 1), (nrows, ncols))
    return rng.value

# === GBA Export Functions ===
def format_project_number(number):
    try:
        f = float(number)
        if f.is_integer():
            return str(int(f))
        return str(number)
    except Exception:
        return str(number)

def get_gba_project_details(sheet):
    last_row_ = find_last_row(sheet)
    last_col_ = find_last_col(sheet)
    block = read_block(sheet, last_row_, last_col_)
    headers = block[0]

    colProject = find_column_index_from_headers(headers, "Project Number") - 1
    colProjectName = find_column_index_from_headers(headers, "Project Name") - 1
    colEmployeeName = find_column_index_from_headers(headers, "Employee Name") - 1
    colDepartmentName = find_column_index_from_headers(headers, "Expenditure Organization Name") - 1

    gba_dict = {}
    for row in block[1:]:
        department_name = (row[colDepartmentName] if colDepartmentName >= 0 else "") or ""
        project_number = (row[colProject] if colProject >= 0 else "") or ""
        project_name = (row[colProjectName] if colProject >= 0 else "") or ""
        resource_name = (row[colEmployeeName] if colEmployeeName >= 0 else "")

        tokens = str(department_name).split()
        gba_value = ""
        for tt in (t.strip() for t in tokens):
            if tt in ("MOB:", "Mobility:"):
                gba_value = "Mobility"; break
            elif tt in ("PLA:", "Places:"):
                gba_value = "Places"; break
            elif tt in ("RES:", "Resilience:"):
                gba_value = "Resilience"; break
            elif tt == "EF:":
                gba_value = "Enabling Function"; break
            elif tt == "SSC:":
                gba_value = "Shared Services"; break

        if gba_value:
            project_number_fmt = format_project_number(project_number)
            gba_dict.setdefault(gba_value, []).append(
                [project_number_fmt, project_name, resource_name, department_name]
            )

    return gba_dict if gba_dict else None

def build_resource_lookup(ws_resource):
    last_row2 = find_last_row(ws_resource)
    block = ws_resource.range((1, 1), (last_row2, 4)).value
    lookup = {}
    for row in block:
        if row[1]:
            lookup[str(row[1]).strip()] = (row[0], row[2], row[3])
    return lookup

def clear_content(workbook):
    ws_ = workbook.sheets["Project Plan Analysis"]
    last_row_ = find_last_row(ws_)
    last_col_ = find_last_col(ws_)
    if last_row_ >= 3:
        ws_.range((3, 1), (last_row_, last_col_)).value = None
    ws_.range((3, 1), (50000, 20)).value = None

def export_gba_data_to_files():
    global wb, ws, gba_file_path
    book = wb
    sheet = ws
    gba_projects = get_gba_project_details(sheet)
    if not gba_projects:
        st.error("No data found!")
        return

    # CHANGE: Track which files were created/updated for progress reporting
    created_files = []
    updated_files = []
    progress_placeholder = st.empty()

    for idx, (gba_value, projects) in enumerate(gba_projects.items()):
        # CHANGE: Show real-time progress during file processing
        progress_placeholder.info(f"Processing GBA: {gba_value} ({idx + 1}/{len(gba_projects)})...")
        
        # Clean the GBA value for filename
        clean_gba_value = clean_file_name(gba_value)
        target_file = os.path.join(
            gba_file_path, "02 GBA Workbooks", f"CPW Tool_{clean_gba_value}_Main.xlsm"
        )
        target_file = target_file.replace("/", os.sep)
        file_exists = os.path.exists(target_file)

        if file_exists:
            target_wb = xw.Book(target_file)
            updated_files.append(f"CPW Tool_{clean_gba_value}_Main.xlsm")
            # CHANGE: Show immediate feedback when updating existing file
            st.success(f"✅ Updating existing file: CPW Tool_{clean_gba_value}_Main.xlsm")
        else:
            template_path = os.path.join(
                gba_file_path, "02 GBA Workbooks", "CPW GBA Specific Template.xlsm"
            )
            target_wb = xw.Book(template_path)
            clear_content(target_wb)
            created_files.append(f"CPW Tool_{clean_gba_value}_Main.xlsm")
            # CHANGE: Show immediate feedback when creating new file
            st.success(f"🆕 Creating new file: CPW Tool_{clean_gba_value}_Main.xlsm")

        try:
            ws_target = target_wb.sheets["Project Plan Analysis"]
        except Exception:
            ws_target = target_wb.sheets.add()
            ws_target.name = "Project Plan Analysis"

        try:
            ws_resource = target_wb.sheets["Resource List"]
            resource_lookup = build_resource_lookup(ws_resource)
        except Exception:
            resource_lookup = {}

        if file_exists:
            next_row = ws_target.api.Cells(ws_target.api.Rows.Count, 1).End(-4162).Row + 1
        else:
            next_row = 2

        output_rows = []
        for idx, proj in enumerate(projects):
            date_val = datetime.today().strftime("%d-%b-%Y")
            serial = next_row + idx - 1
            d_val, e_val, f_val, g_val = proj
            c_val = f"{d_val} - {f_val}"
            h_val = i_val = j_val = ""
            res_info = resource_lookup.get(str(f_val).strip())
            if res_info:
                h_val, i_val, j_val = res_info
            row = [date_val, serial, c_val, d_val, e_val, f_val, g_val, h_val, i_val, j_val]
            output_rows.append(row)

        if output_rows:
            ws_target.range((next_row, 1),
                            (next_row + len(output_rows) - 1, 10)).value = output_rows

        # CHANGE: Enhanced error handling for Excel save operations
        try:
            if not file_exists:
                os.makedirs(os.path.dirname(target_file), exist_ok=True)
                target_wb.save(target_file)
            else:
                target_wb.save()
            target_wb.close()
            # CHANGE: Show save confirmation for each file
            st.info(f"💾 Saved: {len(projects)} entries to {clean_gba_value}")
        except Exception as e:
            st.error(f"❌ Error saving {clean_gba_value}: {e}")
            try:
                target_wb.close()
            except:
                pass

    # CHANGE: Clear progress message and show final summary
    progress_placeholder.empty()
    st.success("🎉 GBA-wise project data processing completed!")
    
    # CHANGE: Show detailed file creation/update summary
    if created_files:
        st.info(f"📁 **New GBA Files Created:** {', '.join(created_files)}")
    
    if updated_files:
        st.info(f"🔄 **GBA Files Updated:** {', '.join(updated_files)}")

# === Team Export Functions ===
def get_team_project_details(sheet):
    project_sheet = sheet.book.sheets['Project Plan Analysis']
    last_row_ = find_last_row(project_sheet)
    last_col_ = find_last_col(project_sheet)
    block = read_block(project_sheet, last_row_, last_col_)
    headers = block[0]

    colOracleDate = find_column_index_from_headers(headers, "Oracle Date")
    colIndex = find_column_index_from_headers(headers, "Index")
    colUniqueCode = find_column_index_from_headers(headers, "Unique Code")
    colProject = find_column_index_from_headers(headers, "Project Number")
    colProjectName = find_column_index_from_headers(headers, "Project Name")
    colEmployeeName = find_column_index_from_headers(headers, "Resource Name")
    colTeamName = find_column_index_from_headers(headers, "Department Name")

    team_dict = {}
    start = var_start_row if var_start_row > 1 else 2
    for r in range(start - 1, last_row_):
        row = block[r]
        oracle_date = row[colOracleDate - 1] if colOracleDate else None
        index = row[colIndex - 1] if colIndex else None
        unique_code = row[colUniqueCode - 1] if colUniqueCode else None
        project_number = (row[colProject - 1] if colProject else "") or ""
        project_name = (row[colProjectName - 1] if colProject else "") or ""
        resource_name = (row[colEmployeeName - 1] if colEmployeeName else "") or ""
        team_name = (row[colTeamName - 1] if colTeamName else "") or ""
        if str(team_name).strip() != "":
            team_dict.setdefault(team_name, []).append(
                [oracle_date, index, unique_code, project_number, project_name, resource_name]
            )
    return team_dict if team_dict else None

def hide_and_protect(workbook):
    today_week = f"Week {date.today().isocalendar()[1]:02d}"

    def hide_columns_for_table(ws, table_name):
        try:
            tbl = ws.api.ListObjects(table_name)
        except Exception:
            return
        hide_flag = False
        for col in tbl.ListColumns:
            header = col.Name
            if "Week" in str(header):
                if today_week in str(header):
                    hide_flag = False
                if hide_flag:
                    col.Range.EntireColumn.Hidden = True
            if "Week 01" in str(header):
                hide_flag = True
                col.Range.EntireColumn.Hidden = True

    try:
        ws = workbook.sheets["Oracle"]
        hide_columns_for_table(ws, "ProjectRaw6")
        ws.api.Unprotect("1234")
        ws.api.Cells.Locked = False
        ws.api.Range("A:AA").Locked = True
        ws.api.Protect("1234", True, True, True)
        try:
            tbl = ws.api.ListObjects("ProjectRaw6")
            tbl.Range.Sort(Key1=tbl.ListColumns("Resource Name").Range, Order1=1)
        except Exception:
            pass
    except Exception:
        pass

    try:
        ws = workbook.sheets["Opportunity | Leaves | Others"]
        hide_columns_for_table(ws, "ProjectRaw6312")
        ws.api.Unprotect("1234")
        ws.api.Protect("1234", True, True, True)
    except Exception:
        pass

    try:
        ws = workbook.sheets["Summary Table"]
        hide_columns_for_table(ws, "Combined")
        ws.api.Unprotect("1234")
        ws.api.Protect("1234", True, True, True)
    except Exception:
        pass

    try:
        ws = workbook.sheets["Capacity Forecast %"]
        hide_flag = False
        for i in range(1, 101):
            header = ws.range((1, i)).value
            if header and "Week" in str(header):
                if today_week in str(header):
                    hide_flag = False
                if hide_flag:
                    ws.range((1, i)).entire_column.hidden = True
            if header and "Week 01" in str(header):
                hide_flag = True
                ws.range((1, i)).entire_column.hidden = True
        ws.api.Unprotect("1234")
        ws.api.Protect("1234", True, True, True)
    except Exception:
        pass

def export_team_data_to_files():
    global wb, ws, team_file_path
    book = wb
    sheet = ws
    team_projects = get_team_project_details(sheet)
    if not team_projects:
        st.error("No data found!")
        return

    # CHANGE: Track which team files were created/updated for progress reporting
    created_team_files = []
    updated_team_files = []
    progress_placeholder = st.empty()

    for idx, (team, projects) in enumerate(team_projects.items()):
        # CHANGE: Show real-time progress during team file processing
        progress_placeholder.info(f"Processing Team: {team} ({idx + 1}/{len(team_projects)})...")
        
        # Clean the team name for filename
        clean_team_name = clean_file_name(team)
        target_file = os.path.join(
            team_file_path, "03 Department Workbooks", f"CPW Tool_{clean_team_name}_Team.xlsm"
        )
        file_exists = os.path.exists(target_file)

        if file_exists:
            target_wb = xw.Book(target_file)
            updated_team_files.append(f"CPW Tool_{clean_team_name}_Team.xlsm")
            # CHANGE: Show immediate feedback when updating existing team file
            st.success(f"✅ Updating existing team file: CPW Tool_{clean_team_name}_Team.xlsm")
        else:
            template_path = os.path.join(
                team_file_path, "03 Department Workbooks", "CPW Team Specific Template.xlsm"
            )
            target_wb = xw.Book(template_path)
            os.makedirs(os.path.dirname(target_file), exist_ok=True)
            created_team_files.append(f"CPW Tool_{clean_team_name}_Team.xlsm")
            # CHANGE: Show immediate feedback when creating new team file
            st.success(f"🆕 Creating new team file: CPW Tool_{clean_team_name}_Team.xlsm")
            
            # CHANGE: Enhanced error handling for new team file creation
            try:
                target_wb.save(target_file)
            except Exception as e:
                st.error(f"❌ Error saving new team workbook {clean_team_name}: {e}")
                continue

        try:
            ws_target = target_wb.sheets["Oracle"]
        except Exception:
            ws_target = target_wb.sheets.add()
            ws_target.name = "Oracle"
        if ws_target is None:
            raise Exception("Failed to create or access 'Oracle' sheet")

        if file_exists:
            next_row = find_first_empty_row_in_col(ws_target, col=5, start=5)
        else:
            next_row = 5

        ws_target.api.Unprotect("1234")
        if projects:
            ws_target.range((next_row, 2), (next_row + len(projects) - 1, 7)).value = projects
        ws_target.api.Protect("1234", True, True, True)

        hide_and_protect(target_wb)

        # CHANGE: Enhanced error handling for team file save operations
        try:
            target_wb.save()
            # CHANGE: Show save confirmation for each team file
            st.info(f"💾 Saved: {len(projects)} entries to {clean_team_name}")
        except Exception as e:
            st.error(f"❌ Error saving team workbook {clean_team_name}: {e}")
        
        try:
            target_wb.close()
        except Exception as e:
            st.error(f"❌ Error closing team workbook {clean_team_name}: {e}")

    # CHANGE: Clear progress message and show final summary
    progress_placeholder.empty()
    st.success("🎉 Team-wise project data processing completed!")
    
    # CHANGE: Show detailed team file creation/update summary
    if created_team_files:
        st.info(f"📁 **New Team Files Created:** {', '.join(created_team_files)}")
    
    if updated_team_files:
        st.info(f"🔄 **Team Files Updated:** {', '.join(updated_team_files)}")

# === Simple Streamlit UI ===
def simple_gba_tab():
    st.write("GBA Wise Extract")

    st.info("""
            📌 **First Run – GBA Extraction**
            
            1. The Latest generated *Project Plan Analysis–continuous-YYYY-MM-DD.xlsx* file  
               from OLD PFP folder. Right-click the file and select **Copy as path**.

            2. Paste the path in the input box below and press **Enter** and then click the **Run GBA Export** button.
            
            3. The tool will create **GBA Workbooks** in:  
               `CPW FINAL PACKAGE \\ 02 GBA Workbooks` folder.  
               (e.g., *CPW Tool_Places_Main.xlsm*),  
               using the respective GBA-specific template when required.
            """)
    manual_path = st.text_input("Enter Excel file path:", key="gba_manual_path")
    
    if st.button("Run GBA Export", key="gba_export_btn"):
        global gba_file_path, wb, ws, last_row, last_col
        if not manual_path:
            st.warning("Please enter a file path.")
            return
        
        selected_file = clean_path(manual_path)
        try:
            gba_file_path = derive_gba_file_path(selected_file)
            wb = xw.Book(selected_file)
            ws = wb.sheets[0]
            last_row = find_last_row(ws)
            last_col = find_last_col(ws)
            export_gba_data_to_files()  # CHANGE: Now shows real-time progress and file creation info
        except Exception as e:
            st.error(f"Error: {e}")
        finally:
            if wb:
                wb.close()
            wb, ws = None, None

def simple_team_tab():
    global var_start_row, team_file_path, wb, ws, last_row, last_col
    st.write("Team Wise Extract")
    st.info("""
           📌 **First Run – Team-wise Extraction**
           1. Locate the respective **GBA Workbook**.  
              For example, if working on *PLA BE*, select *CPW Tool_Places_Main.xlsm* from `CPW FINAL PACKAGE \\ 02 GBA Workbooks` folder.  
              Right-click the file and choose **Copy as path**.
           
           2. Paste the path in the input box below, keep the **Starting Row** as `2`,  
              and then click the **Run Team Export** button.
           
           3. The tool will generate **Team Workbooks** in:  
              `CPW FINAL PACKAGE \\ 03 Department Workbooks` folder.
           
              Each Team Workbook will include:  
              - Updated **Oracle** sheets.  
              - Protections applied to critical sheets.  
              - Older week columns auto-hidden for clarity.
           """)

    manual_path = st.text_input("Enter GBA workbook path:", key="team_manual_path")
    start_row = st.number_input("Start row", min_value=1, value=var_start_row, key="team_start_row")
    
    if st.button("Run Team Export", key="team_export_btn"):
        if not manual_path:
            st.warning("Please enter a file path.")
            return
            
        var_start_row = int(start_row)
        selected_file = clean_path(manual_path)
        
        try:
            team_file_path = derive_team_file_path(selected_file)
            wb = xw.Book(selected_file)
            ws = wb.sheets[0]
            last_row = find_last_row(ws)
            last_col = find_last_col(ws)
            export_team_data_to_files()  # CHANGE: Now shows real-time progress and file creation info
        except Exception as e:
            st.error(f"Error: {e}")
        finally:
            if wb:
                wb.close()
            wb, ws = None, None

def simple_maintenance_gba_tab():
    st.write("GBA Wise Extract (Maintenance)")
    st.info("""
            📌 **Maintenance – GBA Extraction**
            
            1. The Latest generated *New_PFP_YYYY-MM-DD.xlsx* file  
               from NEW PFP folder. Right-click the file and select **Copy as path**.

            2. Paste the path in the input box below and press **Enter** and then click the **Run GBA Export** button.
            
            3. The tool will create/update **GBA Workbooks** in:  
               `CPW FINAL PACKAGE \\ 02 GBA Workbooks` folder.  
               (e.g., *CPW Tool_Places_Main.xlsm*),  
               using the respective GBA-specific template when required.
            """)
    manual_path = st.text_input("Enter Excel file path:", key="maintenance_gba_manual_path")
    
    if st.button("Run GBA Export", key="maintenance_gba_export_btn"):
        global gba_file_path, wb, ws, last_row, last_col
        if not manual_path:
            st.warning("Please enter a file path.")
            return
        
        selected_file = clean_path(manual_path)
        try:
            gba_file_path = derive_gba_file_path(selected_file)
            wb = xw.Book(selected_file)
            ws = wb.sheets[0]
            last_row = find_last_row(ws)
            last_col = find_last_col(ws)
            export_gba_data_to_files()  # CHANGE: Now shows real-time progress and file creation info
        except Exception as e:
            st.error(f"Error: {e}")
        finally:
            if wb:
                wb.close()
            wb, ws = None, None

def simple_maintenance_team_tab():
    global var_start_row, team_file_path, wb, ws, last_row, last_col
    st.write("Team Wise Extract (Maintenance)")
    st.info("""
            📌 **Maintenance – Team-wise Extraction**
            
            1. Locate the respective **GBA Workbook**.  
               For example, if working on *PLA BE*, select *CPW Tool_Places_Main.xlsm*  
               from the `CPW FINAL PACKAGE \\ 02 GBA Workbooks` folder.  
            
               - Open the workbook and check the **Oracle Date** column.  
               - Identify the row where the **current week’s date** (today’s date) begins.  
               - Note down this row number as the **Starting Row**.  
               - After confirming, ensure the GBA workbook is **closed**.  
            
               Finally, right-click the file and select **Copy as path**.

            2. Paste the copied path in the input below and press **Enter**.  
               Then, enter the **Starting Row** value and click the **Run Team Export** button.
            
            3. The tool will generate or update the respective **Team Workbooks** in:  
               `CPW FINAL PACKAGE \\ 03 Department Workbooks`  
            
               Each Team Workbook will include:  
               - Updated **Oracle** sheets.  
               - Protections applied to critical sheets.  
               - Auto-hiding of older week columns for clarity.
            """)


    manual_path = st.text_input("Enter GBA workbook path:", key="maintenance_team_manual_path")
    start_row = st.number_input("Start row", min_value=1, value=var_start_row, key="maintenance_team_start_row")
    
    if st.button("Run Team Export", key="maintenance_team_export_btn"):
        if not manual_path:
            st.warning("Please enter a file path.")
            return
            
        var_start_row = int(start_row)
        selected_file = clean_path(manual_path)
        
        try:
            team_file_path = derive_team_file_path(selected_file)
            wb = xw.Book(selected_file)
            ws = wb.sheets[0]
            last_row = find_last_row(ws)
            last_col = find_last_col(ws)
            export_team_data_to_files()  # CHANGE: Now shows real-time progress and file creation info
        except Exception as e:
            st.error(f"Error: {e}")
        finally:
            if wb:
                wb.close()
            wb, ws = None, None

def selection_page():
    st.title("Capacity Planning Workbook (CPW) Tool")
    st.info(""" 
            - The Capacity Planning Workbook (CPW) Tool is designed to support teams in making informed, forward-looking decisions about resource allocation.
            - It offers a structured and consistent approach to capacity planning, helping departments identify which resources are available, when, and for which types of work — whether billable projects, internal initiatives, or new opportunities.
            - This tool is not intended for financial tracking or budgeting. Instead, it focuses on headcount and effort availability, giving visibility into workforce readiness across time horizons — from current commitments to future demand.
         """)
    
    if "ba_selected" not in st.session_state:
        st.session_state["ba_selected"] = ""
    if "gba_selected" not in st.session_state:
        st.session_state["gba_selected"] = ""

    col1, col2 = st.columns(2)
    
    with col1:
        ba_options = ["", "A&U", "APM", "Australia", "Belgium", "Brazil", "Canada", "Chile", "France","Germany", "Italy", "Netherlands", "Peru", "Philippines", "Poland", "Spain & Portugal", "UK & I","US"]
        ba_selected = st.selectbox("Business Area (BA):", ba_options, key="ba_selectbox")
        st.session_state["ba_selected"] = ba_selected

    with col2:
        gba_options = ["", "Places", "Mobility", "Resilience"]
        gba_selected = st.selectbox("Global Business Area (GBA):", gba_options, key="gba_selectbox")
        st.session_state["gba_selected"] = gba_selected

    if ba_selected and gba_selected:
        if st.button("Proceed to Processing", key="proceed_btn"):
            st.session_state["current_page"] = "processing"
            st.rerun()
    else:
        st.warning("Please select both BA and GBA")

def processing_page():
    ba = st.session_state.get("ba_selected", "")
    gba = st.session_state.get("gba_selected", "")
    
    st.title("Capacity Planning Workbook (CPW) Tool")
    st.title(f"Processing: {ba} - {gba}")
    
    if st.button("← Back", key="back_btn"):
        st.session_state["current_page"] = "selection"
        st.rerun()
    
    process_tabs = st.tabs(["1st Time Run", "Maintenance"])
    
    with process_tabs[0]:
        tabs = st.tabs(["PFP Processing", "GBA Extraction", "Team Extraction"])

        with tabs[0]:
            st.write("PFP Processing")
            # 1) FIRST RUN — PFP Processing
            st.info("""
            📌 **First Run – PFP Processing**
            
            1. Navigate to your respective **{GBA - BA} folder** and open:  
               `CPW FINAL PACKAGE \\ 01 Data Processing \\ Project Financial Plan (PFP)`
            
            2. In the **Raw Data Folder**, locate the latest **RAW PFP file**.  
               Right-click the file and select **Copy as path**.
            
            3. Paste the file path in the input box below and press **Enter**.
            
            4. When the **Add Unique Code** button appears, click it.  
               This step assigns unique codes to the RAW PFP and saves it as:  
               *Project Plan Analysis–continuous.xlsx* in the `CPW FINAL PACKAGE \\ 01 Data Processing \\ Project Financial Plan (PFP)` folder.
            
            5. Next, click the **Clean & Save** button when prompted.  
               This will:  
               - Remove duplicates and blanks (Employee Names, Labor Cost, Conversion Employee Names).  
               - Save the cleaned version as *Project Plan Analysis–continuous-YYYY-MM-DD.xlsx* in the **OLD PFP** folder.
            """)


            manual_path = st.text_input("Raw Data File path:", key="pfp_manual_path")
            
            if manual_path:
                selected_file = clean_path(manual_path)
                try:
                    raw_data_folder = os.path.dirname(selected_file)
                    pfp_folder = os.path.dirname(raw_data_folder)
                    project_plan_path = os.path.join(pfp_folder, "Project Plan Analysis-continuous.xlsx")
                    old_pfp_folder = os.path.join(pfp_folder, "OLD PFP")
                    
                    df_raw = pd.read_excel(selected_file)
                    st.write(f"Raw Data: {df_raw.shape[0]} rows")
                    
                    if st.button("Add Unique Code", key="create_project_plan_btn"):
                        unique_df = first_time_unique_code_run_pfp(df_raw)
                        unique_df.to_excel(project_plan_path, index=False)
                        st.success("Project Plan Analysis created!")
                        # CHANGE: Show preview of data with unique codes
                        st.write("**Project Plan Analysis Preview (with Unique Code):**")
                        st.dataframe(unique_df.head(3))
                        st.session_state["add_unique_clicked"] = True
                        st.session_state["unique_df"] = unique_df
                        st.session_state["old_pfp_folder"] = old_pfp_folder

                    if st.session_state.get("add_unique_clicked", False):
                        if st.button("Clean & Save", key="clean_pfp_btn"):
                            unique_df = st.session_state.get("unique_df")
                            cleaned_df = first_time_run_pfp(unique_df)
                            final_date_str = datetime.now().strftime('%Y-%m-%d')
                            cleaned_file_name = f"Project Plan Analysis-continuous-{final_date_str}.xlsx"
                            cleaned_file_path = os.path.join(old_pfp_folder, cleaned_file_name)
                            os.makedirs(old_pfp_folder, exist_ok=True)
                            cleaned_df.to_excel(cleaned_file_path, index=False)
                            st.success(f"Cleaned data saved to OLD PFP: {cleaned_file_name}")
                            
                            # CHANGE: Show cleaning statistics
                            if "cleaning_stats" in st.session_state:
                                stats = st.session_state["cleaning_stats"]
                                st.info(f"""
                                **Data Cleaning Summary:**
                                - Original rows: {stats['original_count']:,}
                                - Duplicates removed: {stats['duplicates_removed']:,}
                                - Blank employees removed: {stats['blank_employees_removed']:,}
                                - Labor Cost, Conversion Employee entries removed: {stats['labor_cost_removed']:,}
                                - Final clean rows: {stats['final_count']:,}
                                """)

                            # CHANGE: Show preview of cleaned data
                            st.write("**Cleaned Data Preview:**")
                            st.dataframe(cleaned_df.head(3))

                            # CHANGE: Show final summary
                            st.info(f"""
                            **Final Summary:**
                            - Ready for GBA and Team extraction ✅
                            """)
                            # CHANGE: Clear session state after successful completion
                            st.session_state.pop("add_unique_clicked", None)
                            st.session_state.pop("unique_df", None)
                            st.session_state.pop("old_pfp_folder", None)
                            
                except Exception as e:
                    st.error(f"Error: {e}")

        with tabs[1]:
            simple_gba_tab()

        with tabs[2]:
            simple_team_tab()
    
    with process_tabs[1]:
        tabs = st.tabs(["PFP Processing", "GBA Extraction", "Team Extraction"])

        with tabs[0]:
            st.write("Maintenance PFP Processing")
            st.info("""
                    📌 **Maintenance – PFP Update**
                    
                    **Step 1 – Current Week**  
                    1. Navigate to your respective **{GBA - BA} folder** and open:  
                       `CPW FINAL PACKAGE \\ 01 Data Processing \\ Project Financial Plan (PFP)`
                    
                    2. In the **Raw Data** folder, locate the latest **RAW PFP file**.  
                       Right-click the file and select **Copy as path**.
                    
                    3. Paste the file path in the input box below and press **Enter**.
                    
                    4. Click the **Process Current Week** button.  
                       This will:  
                       - Assign unique codes to the current week’s RAW PFP.  
                       - Clean the data by removing duplicates and blanks.  
                       - Save the cleaned version as *Project Plan Analysis–continuous-YYYY-MM-DD.xlsx*  
                         in the **OLD PFP** folder.  
                    
                    ---
                    
                    **Step 2 – Comparison**  
                    1. Provide both the **last week’s cleaned file** and the **current week’s file**  
                       from the **OLD PFP** folder.  
                    
                    2. If new *Unique Code* rows are identified:  
                       - Click **Generate New PFP**, and then click **Save New PFP**.  
                       - The updated file will be saved in the **NEW PFP** folder as:  
                         *New_PFP_YYYY-MM-DD.xlsx*  
                    
                    3. This newly generated file will be used for the next step: **GBA Extraction**.
                    """)

            
            # CHANGE: Added current week raw data processing step with data cleaning
            st.subheader("Step 1: Process Current Week Raw Data")
            current_raw_path = st.text_input("Current Week Raw Data File path:", key="maintenance_current_raw_path")
            
            if current_raw_path:
                current_raw_file = clean_path(current_raw_path)
                try:
                    raw_data_folder = os.path.dirname(current_raw_file)
                    pfp_folder = os.path.dirname(raw_data_folder)
                    project_plan_path = os.path.join(pfp_folder, "Project Plan Analysis-continuous.xlsx")
                    old_pfp_folder = os.path.join(pfp_folder, "OLD PFP")
                    
                    df_current_raw = pd.read_excel(current_raw_file)
                    st.write(f"Current Week Raw Data: {df_current_raw.shape[0]} rows")
                    
                    if st.button("Process Current Week", key="process_current_week_btn"):
                        current_unique_df = first_time_unique_code_run_pfp(df_current_raw)
                        current_unique_df.to_excel(project_plan_path, index=False)
                        
                        current_cleaned_df = first_time_run_pfp(current_unique_df)
                        
                        current_date_str = datetime.now().strftime('%Y-%m-%d')
                        current_cleaned_file_name = f"Project Plan Analysis-continuous-{current_date_str}.xlsx"
                        current_cleaned_file_path = os.path.join(old_pfp_folder, current_cleaned_file_name)
                        os.makedirs(old_pfp_folder, exist_ok=True)
                        current_cleaned_df.to_excel(current_cleaned_file_path, index=False)
                        
                        st.success(f"Current week processed and saved: {current_cleaned_file_name}")
                        
                        # CHANGE: Show processing statistics for current week
                        if "cleaning_stats" in st.session_state:
                            stats = st.session_state["cleaning_stats"]
                            st.info(f"""
                            **Current Week Processing Summary:**
                            - Raw data rows: {stats['original_count']:,}
                            - Duplicates removed: {stats['duplicates_removed']:,}
                            - Blank employees removed: {stats['blank_employees_removed']:,}
                            - Labor Cost, Conversion Employee entries removed: {stats['labor_cost_removed']:,}
                            - Final processed rows: {stats['final_count']:,}
                            """)
                        
                        st.session_state["current_week_processed"] = True
                        st.session_state["current_cleaned_path"] = current_cleaned_file_path
                        
                except Exception as e:
                    st.error(f"Error processing current week: {e}")
            
            st.divider()
            st.subheader("Step 2: Compare with Previous Week")
            
            col1, col2 = st.columns(2)
            
            with col1:
                prev_week_path = st.text_input("Previous Week File:", key="maintenance_prev_path")
            
            with col2:
                if st.session_state.get("current_week_processed", False):
                    current_week_path = st.session_state.get("current_cleaned_path", "")
                    st.text_input("Current Week File:", value=current_week_path, disabled=True, key="maintenance_current_path_display")
                else:
                    current_week_path = st.text_input("Current Week File:", key="maintenance_current_path")

            if prev_week_path and current_week_path:
                try:
                    df_prev_week = pd.read_excel(clean_path(prev_week_path))
                    df_current_week = pd.read_excel(clean_path(current_week_path))
                    
                    st.write(f"Previous: {len(df_prev_week)} rows, Current: {len(df_current_week)} rows")
                    
                    # CHANGE: Show preview of both weeks' data
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**Previous Week Preview:**")
                        st.dataframe(df_prev_week.head(2))
                    with col2:
                        st.write("**Current Week Preview:**")
                        st.dataframe(df_current_week.head(2))

                    if st.button("Generate New PFP", key="generate_new_pfp_btn"):
                        prev_week_unique_codes = df_prev_week['Unique Code']
                        current_week_unique_codes = df_current_week['Unique Code']
                        is_new_row = ~current_week_unique_codes.isin(prev_week_unique_codes)
                        df_new_pfp = df_current_week[is_new_row].copy()
                        
                        if len(df_new_pfp) > 0:
                            st.success(f"Found {len(df_new_pfp)} new entries!")

                            # CHANGE: Show detailed comparison statistics
                            st.info(f"""
                            **Comparison Results:**
                            - Previous week unique codes: {len(prev_week_unique_codes.unique()):,}
                            - Current week unique codes: {len(current_week_unique_codes.unique()):,}
                            - New entries (not in previous): {len(df_new_pfp):,}
                            - Duplicate entries (already existed): {len(df_current_week) - len(df_new_pfp):,}
                            """)
                            # CHANGE: Show preview of new entries
                            st.write("**New PFP Entries Preview:**")
                            st.dataframe(df_new_pfp.head(3))
                            
                            st.session_state["df_new_pfp_ready"] = df_new_pfp
                            st.session_state["new_pfp_entries_found"] = True
                            
                            old_pfp_folder = os.path.dirname(clean_path(prev_week_path))
                            pfp_base_folder = os.path.dirname(old_pfp_folder)
                            new_pfp_folder = os.path.join(pfp_base_folder, "NEW PFP")
                            st.session_state["new_pfp_folder"] = new_pfp_folder
                        else:
                            st.warning("No new entries found")
                            st.info("All current week entries already existed in previous week")

                except Exception as e:
                    st.error(f"Error: {e}")

            if st.session_state.get("new_pfp_entries_found", False):
                df_new_pfp = st.session_state.get("df_new_pfp_ready")
                if df_new_pfp is not None:
                    st.write(f"Ready to save: {len(df_new_pfp)} entries")
                    
                    if st.button("Save New PFP", key="save_new_pfp_btn"):
                        try:
                            new_pfp_folder = st.session_state.get("new_pfp_folder")
                            os.makedirs(new_pfp_folder, exist_ok=True)
                            timestamp = datetime.now().strftime('%Y-%m-%d')
                            new_pfp_filename = f"New_PFP_{timestamp}.xlsx"
                            new_pfp_path = os.path.join(new_pfp_folder, new_pfp_filename)
                            df_new_pfp.to_excel(new_pfp_path, index=False)
                            st.success(f"New PFP saved to NEW PFP folder: {new_pfp_filename}")
                            
                            # CHANGE: Show final summary
                            st.info(f"""
                            **Final Summary:**
                            - New entries saved: {len(df_new_pfp):,}
                            - File location: NEW PFP folder
                            - Ready for GBA and Team extraction ✅
                            """)
                            
                            st.session_state.pop("df_new_pfp_ready", None)
                            st.session_state["new_pfp_entries_found"] = False
                            st.rerun()
                        except Exception as e:
                            st.error(f"Save error: {e}")

        with tabs[1]:
            simple_maintenance_gba_tab()

        with tabs[2]:
            simple_maintenance_team_tab()

def main():
    if "current_page" not in st.session_state:
        st.session_state["current_page"] = "selection"
    
    if st.session_state["current_page"] == "selection":
        selection_page()
    elif st.session_state["current_page"] == "processing":
        processing_page()

if __name__ == "__main__":
    st.set_page_config(page_title="Workforce Planning Tool", layout="wide")
    main()