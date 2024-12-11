import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
import os

# Dropbox Access Token
ACCESS_TOKEN = "sl.u.AFbSi0u1JEt41HfRfC_wgf5aX0jwZCocifSqL6t8JaJvCL1Q86YPZMr9VcQevwbYrsqY87MzcGxefFW3kL_acbhKoOuqZ2dwrNKTP-2HhBFJdZm9NZWgrLgqD_qMBJBlbCF7IjI0nRNeUzgm44Q5fBZrUGFvYbDznCb_E3NdsYUpdm5vML2362RLKifpuRK2Sh0K6da6MbyysftZCf11oEaAHS-jY45uP21ps5geNPm5JRSTffJ0Z2zj_GZ1AWpSwJCFPsLXYJRmYxLwTNh8OxFSCgQU_gRCrG3ymr99wWNNN-njybj47lR27dCmPKj6jhhQT14Tu7nG9E-w8Jt9OywK3msYI18tCO_SZ21A3GqOFwdVghSazBn6s3uXzeDN8qmNteTSi1RWkHIOyTa_ItYkkWbY7xts_mIkWQpUhdsxb9PfA-jJQtEy7yt2cMD-sEekDJhHBtejqHIVNSBA1QJbA7Dc_lmu34i7Jf_1CxPfqZdsWdrbqtQbl6VAFS8xR0FtBtnf4il1FHsEV4eyoc93OLGeWYf-ikLE_KPS5vwXyHxXxWO3DNWu3tVu8BmXiH26B7BYVkRaD6PP-ANRtqS4Vi27Ve21N0GMopBqEyiwz63lck8Fxu2NgdWvPntYqcIjjgHyQVGsP-jBXgFvja-KJHMzHC3V9hMiQ4CPJMnnB9kFU9nHWVLpHxhQvMKM99XW8hqgIzl_RMXDajmvxeXG2fvTrT3ktEl8vXRrqg2maIN3-zq5ef1dhelmkfGvV7I7nI0Zv3stGW3F0WYJFX33hrCEl7MdToABlpPqlMA0b80njb5nF_YoAuayKWUCemaIYtAOn5PnsIzbE6tshtb0iIVldF4vStfgjIG6FuOh7IiYaU1FrIWHGBhs0yey_-IDGt_RAmuHRTp2IW_Jt5ByxrR748H9JignxTmGSfJUKQuaX6vSkzThIoJ6elkcj6JFXOphbJ33_DwuPVm1cXfk9NGz8Fagb56TQdb8MtDaP1wqg2ERur_PIOv3vuHHsKqrdm5Y4bgXryrwEUDKf-mfXhDivk1ATGUfghfgVnIBLznpMgoWoML_TCnuUaPrfa5Ajyo2goHHNeuZQ8nzHB4AMo6d7Tb_HxB9AVkwddoJ2O-sQm-ks0rv5VpdiiApjEy7Hz9YFwgqC7bMf-Ex5BMiFmVyrrvT_BinF5KWuXS_dORZsjdQwwPZsNr2TpX2rIZ76dvJAtPWSqiVo6gSouAnKFlrNmTc5JrHPBTV_GH4H81lYGLVj2HycAY7c3xD_wQ37HVs_R3F7BGcVWoEUm1fzgD3Darj3oQz4MS75hZGEDP1YKcmuM3s6J6O6kSrv_fwcF3oPilsRWBQYYwKiVeJLamanRdxxnFndMEPhZ0s5n5EuXg6_v3rsm0FOmTyVNKTOJ1cYAZ5Wo5KHkhpDRPMX4qkrDdV4_F7v59tRDd5iA"
DROPBOX_FOLDER_PATH = "/Project_Data"

# Dropbox Functions
def upload_to_dropbox(file_path, dropbox_path, access_token):
    """Upload a file to Dropbox."""
    try:
        dbx = dropbox.Dropbox(access_token)
        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
        print(f"Uploaded to Dropbox: {dropbox_path}")
    except Exception as e:
        print(f"Error uploading to Dropbox: {e}")

def download_from_dropbox(dropbox_path, local_path, access_token):
    """Download a file from Dropbox."""
    try:
        dbx = dropbox.Dropbox(access_token)
        metadata, res = dbx.files_download(dropbox_path)
        with open(local_path, "wb") as f:
            f.write(res.content)
        print(f"Downloaded from Dropbox: {dropbox_path}")
    except dropbox.exceptions.ApiError as e:
        print(f"Dropbox file not found: {dropbox_path}")
    except Exception as e:
        print(f"Error downloading from Dropbox: {e}")

# Ensure File Exists
def ensure_file_exists(file_path, dropbox_path, columns):
    """Ensure the file exists locally or in Dropbox. Create it if not."""
    if not os.path.exists(file_path):
        try:
            # Attempt to download from Dropbox
            download_from_dropbox(dropbox_path, file_path, ACCESS_TOKEN)
        except Exception:
            # Create a new file if Dropbox download fails
            pd.DataFrame(columns=columns).to_excel(file_path, index=False)
            upload_to_dropbox(file_path, dropbox_path, ACCESS_TOKEN)
            print(f"File created and uploaded: {file_path}")

# Main Application
def main():
    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    # File paths
    local_projects_path = "projects_data_weekly.xlsx"
    dropbox_projects_path = f"{DROPBOX_FOLDER_PATH}/projects_data_weekly.xlsx"

    # Ensure the projects file exists
    ensure_file_exists(
        local_projects_path,
        dropbox_projects_path,
        ["Project ID", "Project Name", "Personnel", "Week", "Year", "Month", "Budgeted Hrs", "Spent Hrs"]
    )

    # Load resources
    hr_file_path = 'Human Resources.xlsx'
    hr_sections = pd.ExcelFile(hr_file_path).sheet_names

    # Step 1: User Selection
    st.subheader("Step 1: Choose Action")
    action = st.radio("What would you like to do?", ["Create New Project", "Update Existing Project"])

    # Section Selection
    st.subheader("Step 2: Select Section")
    selected_section = st.selectbox("Choose a section", hr_sections)

    engineers_data = pd.read_excel(hr_file_path, sheet_name=selected_section)
    engineer_names = engineers_data["Name"].dropna().tolist()

    if action == "Create New Project":
        st.subheader("Step 3: Project Details")
        project_id = st.text_input("Project ID")
        project_name = st.text_input("Project Name")
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
        selected_month = st.selectbox("Month", list(month_name)[1:])
        selected_month_index = list(month_name).index(selected_month)

        weeks = [(f"Week {i}", i) for i in range(1, 5)]
        st.subheader("Step 4: Assign Engineers and Weekly Hours")
        assigned_engineers = st.multiselect("Select Engineers", engineer_names)

        if "engineer_allocation" not in st.session_state:
            st.session_state.engineer_allocation = {}

        for engineer in assigned_engineers:
            st.markdown(f"**Engineer: {engineer}**")
            for week_label, week_number in weeks:
                budgeted_hours = st.number_input(
                    f"Budgeted Hours ({week_label})", min_value=0, step=1, key=f"{engineer}_{week_label}"
                )
                if st.button("Add Hours", key=f"add_{engineer}_{week_label}"):
                    unique_key = f"{engineer}_{week_label}"
                    st.session_state.engineer_allocation[unique_key] = {
                        "Project ID": project_id,
                        "Project Name": project_name,
                        "Personnel": engineer,
                        "Week": week_label,
                        "Year": selected_year,
                        "Month": selected_month,
                        "Budgeted Hrs": budgeted_hours,
                        "Spent Hrs": 0,
                    }
                    st.success(f"Added hours for {engineer} in {week_label}.")

        if st.session_state.engineer_allocation:
            summary_data = pd.DataFrame(st.session_state.engineer_allocation.values())
            st.dataframe(summary_data)
            if st.button("Submit Project"):
                try:
                    existing_data = pd.read_excel(local_projects_path)
                    final_data = pd.concat([existing_data, summary_data], ignore_index=True)
                except FileNotFoundError:
                    final_data = summary_data
                final_data.to_excel(local_projects_path, index=False)
                upload_to_dropbox(local_projects_path, dropbox_projects_path, ACCESS_TOKEN)
                st.success(f"Project '{project_name}' submitted!")

    elif action == "Update Existing Project":
        try:
            projects_data = pd.read_excel(local_projects_path)
        except FileNotFoundError:
            st.warning("No existing projects found.")
            return

        project_names = projects_data["Project Name"].unique().tolist()
        selected_project = st.selectbox("Select a project to update", project_names)

        project_data = projects_data[projects_data["Project Name"] == selected_project]
        updated_data = []

        for _, row in project_data.iterrows():
            budgeted_hours = st.number_input(f"Budgeted Hours for {row['Personnel']} ({row['Week']})", value=row["Budgeted Hrs"])
            spent_hours = st.number_input(f"Spent Hours for {row['Personnel']} ({row['Week']})", value=row["Spent Hrs"])
            updated_data.append({
                **row,
                "Budgeted Hrs": budgeted_hours,
                "Spent Hrs": spent_hours,
            })

        if st.button("Save Updates"):
            remaining_data = projects_data[projects_data["Project Name"] != selected_project]
            final_data = pd.concat([remaining_data, pd.DataFrame(updated_data)], ignore_index=True)
            final_data.to_excel(local_projects_path, index=False)
            upload_to_dropbox(local_projects_path, dropbox_projects_path, ACCESS_TOKEN)
            st.success(f"Updates to project '{selected_project}' saved!")

if __name__ == "__main__":
    main()
