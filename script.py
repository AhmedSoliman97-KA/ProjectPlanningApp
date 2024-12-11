import streamlit as st
import pandas as pd
from datetime import datetime
import dropbox
import os

# Dropbox Access Token and Paths
ACCESS_TOKEN = "sl.u.AFZfAbFllA29neJgjpRx63DedZKVU5jYc9_ze4-Y4HhSdTfnCW3s6NqQVy-8JwH5sa0rS1SKhdJUeFdRk6vVBPF3brz0sjQmPX7W-H_7lSoZj2224z8wzUWgYmHA0tIS6V9eRd34rj-bXqUSQ4zzrFflowgqXxFbt12e6bxsdkzSVSvXO8f7G1MRBQC3lln3XFbYYMRpJwnjO1nNs3HrPuo2tToG7VuQps4yPrilW1XaucBx5g2kDM83VmW0VSfIb24RO2XxbrJB_x11XXigsy1ES0hV0J7vVW6tD1mrg13wmwx92imzWePsiskx0lMKQ5bHGnI3_XfqR1YW18rt8nHgg983LJQ24Iu_iMiWiX6gcyMMaB5iv6DvFjF2AGSvuB-oe-yzXWq35ATuvhGxL1iBTNyKZGAcY7fgNBB8Km__d61aY-wttgWNG35Dp0VdLfSdDiZT-r_3M1MA9uTK7yYhLKmJg-MvPsVLp1Vh_AF7HRCiYJDulcTCzTNyvVMDlEFac5NidFL4shCOLXP3oMKT4_z0HcdxSnwQ0k9-uFSg1BV8XQaP7u0UiE_16fLtvKZaa9H2YFO6FiN14rYt6czeTb1Jqnv9UqJ-P1RLSEma2BZ-kL04ntXBdwsOPReHxAKD2TTpbDQZk8nXXwyeC-l1Ho2FMaR2lx7lp8QN1PNCTjSFS794uPUnGH_OuDwkSKI6uP7KfSgQFgl-BNXhVkdjHuzyomLnrKuKu-S53UNTf74p-FOYO_RT7cC_rtI3g3TD8MlIEpNfCyT_IQaXI3YFparHkpK27XjVlO1sJ7Bf6Hxg0gGnLMzdJl2HBXrAdr684QFpYscWagvTBX-UoofSHnotdhIMc65fHaSH3mYlerKwYx_IafYnWvhN5L4lGnMhlK4Kp6tj7_6fPqEJN9PtsPO1Ob3YFdN4IMovTGB-Z4kcDq3EsQZMBIITaFbIhWgEjnSVGruXPB0IM0m1nZxTGxbbNB8F02Dy2P3uwZqNRXL4mwSlg5H-jUtGQw4opTeAKRp2lycSNHlp8lLVDrekXHmvOyg7VQhgAVEfm6n51NX3KlUjSkKva3N3QySFiwUhMpMBUX8eRyGuMvSpcfIMKvmJc0vAhIagiG50u9LjPixrPIQp0KY8-z6wOikBRzA6CjefibaS5Cw4BA1yvGrEeSQPvGg6ADOv1HgDeoHuMk78ZT1daAyvfCCHGWiBHQ1kmUw6V0egKDiO-92FEOZwxxwOJKlyhxJ7T5NrnFVKcT5Tu9GBW-CeaHN-MN4EhJFU7LM5ej6mNrBYCVKa1L_s02g6OjfaRb8EZFDgcBIOvij85rYoZNDr-x4jlPbCDXYuYDSYLjstTUa7WQidObkHh-JGnHfoCyJPuJusk8rvsqUwzKVSHM-VTrZgUjUE39u9BSi98HWQWaRMhvsksaR1z_rJ2WJWDMo7EZNHf3QXXg"
DROPBOX_FOLDER_PATH = "/Project_Data"
HR_FILE_NAME = "Human Resources.xlsx"
LOCAL_HR_FILE = HR_FILE_NAME
DROPBOX_HR_PATH = f"{DROPBOX_FOLDER_PATH}/{HR_FILE_NAME}"
LOCAL_PROJECTS_FILE = "projects_data_weekly.xlsx"
DROPBOX_PROJECTS_PATH = f"{DROPBOX_FOLDER_PATH}/projects_data_weekly.xlsx"

# Dropbox Functions
def upload_to_dropbox(file_path, dropbox_path, access_token):
    """Upload a file to Dropbox."""
    try:
        dbx = dropbox.Dropbox(access_token)
        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
        st.success(f"File successfully uploaded to Dropbox: {dropbox_path}")
    except dropbox.exceptions.AuthError as e:
        st.error(f"Authentication error: {e}")
    except Exception as e:
        st.error(f"Error uploading to Dropbox: {e}")

def download_from_dropbox(dropbox_path, local_path, access_token):
    """Download a file from Dropbox."""
    try:
        dbx = dropbox.Dropbox(access_token)
        metadata, res = dbx.files_download(dropbox_path)
        with open(local_path, "wb") as f:
            f.write(res.content)
        print(f"Downloaded from Dropbox: {dropbox_path}")
    except dropbox.exceptions.ApiError as e:
        print(f"Dropbox file not found: {e}")
    except Exception as e:
        print(f"Error downloading file: {e}")

# Ensure File Exists or Create One
def ensure_file_exists(file_path, dropbox_path, columns=None):
    """Ensure the file exists locally or in Dropbox. Create it if not."""
    if not os.path.exists(file_path):
        try:
            download_from_dropbox(dropbox_path, file_path, ACCESS_TOKEN)
        except Exception:
            if columns:
                pd.DataFrame(columns=columns).to_excel(file_path, index=False)
                upload_to_dropbox(file_path, dropbox_path, ACCESS_TOKEN)
                print(f"File created and uploaded: {file_path}")

# Load Engineers Data
def load_engineers(file_path, selected_section):
    """Load engineer names from the Human Resources file."""
    try:
        data = pd.read_excel(file_path, sheet_name=selected_section)
        return data["Name"].dropna().tolist()
    except Exception as e:
        st.error(f"Error loading Human Resources data: {e}")
        return []

# Main Application
def main():
    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    # Ensure project and HR files exist
    ensure_file_exists(
        LOCAL_PROJECTS_FILE,
        DROPBOX_PROJECTS_PATH,
        ["Project ID", "Project Name", "Personnel", "Week", "Year", "Month", "Budgeted Hrs", "Spent Hrs"]
    )
    ensure_file_exists(LOCAL_HR_FILE, DROPBOX_HR_PATH)

    # Load available sections from Human Resources file
    try:
        hr_file = pd.ExcelFile(LOCAL_HR_FILE)
        sections = hr_file.sheet_names
    except Exception as e:
        st.error("Unable to load sections from Human Resources file.")
        return

    # Step 1: User Selection
    st.subheader("Step 1: Choose Action")
    action = st.radio("What would you like to do?", ["Create New Project", "Update Existing Project"])

    # Section Selection
    st.subheader("Step 2: Select Section")
    selected_section = st.selectbox("Choose a section", sections)

    # Load Engineers
    engineers = load_engineers(LOCAL_HR_FILE, selected_section)
    if not engineers:
        st.warning("No engineers found in the selected section.")
        return

    if action == "Create New Project":
        # Project Details
        st.subheader("Step 3: Enter Project Details")
        project_id = st.text_input("Project ID", help="Enter a unique ID for the project.")
        project_name = st.text_input("Project Name")
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
        selected_month = st.selectbox("Month", range(1, 13))
        weeks = [f"Week {i}" for i in range(1, 5)]

        st.subheader("Step 4: Assign Engineers and Weekly Hours")
        selected_engineers = st.multiselect("Select Engineers", engineers)

        # Initialize Allocation Data
        if "engineer_allocation" not in st.session_state:
            st.session_state.engineer_allocation = {}

        for engineer in selected_engineers:
            st.markdown(f"**Engineer: {engineer}**")
            for week in weeks:
                budgeted_hours = st.number_input(
                    f"Budgeted Hours ({week})", min_value=0, step=1, key=f"{engineer}_{week}_budgeted"
                )
                spent_hours = st.number_input(
                    f"Spent Hours ({week})", min_value=0, step=1, key=f"{engineer}_{week}_spent"
                )
                if st.button("Add Hours", key=f"add_{engineer}_{week}"):
                    unique_key = f"{engineer}_{week}"
                    st.session_state.engineer_allocation[unique_key] = {
                        "Project ID": project_id,
                        "Project Name": project_name,
                        "Personnel": engineer,
                        "Week": week,
                        "Year": selected_year,
                        "Month": selected_month,
                        "Budgeted Hrs": budgeted_hours,
                        "Spent Hrs": spent_hours,
                    }
                    st.success(f"Added hours for {engineer} in {week}.")

        # Display Current Data
        if st.session_state.engineer_allocation:
            st.subheader("Allocation Summary")
            summary_data = pd.DataFrame(st.session_state.engineer_allocation.values())
            st.dataframe(summary_data)

            if st.button("Submit Project"):
                try:
                    # Load existing data
                    try:
                        existing_data = pd.read_excel(LOCAL_PROJECTS_FILE)
                    except FileNotFoundError:
                        existing_data = pd.DataFrame(columns=summary_data.columns)

                    # Append and save
                    final_data = pd.concat([existing_data, summary_data], ignore_index=True)
                    final_data.to_excel(LOCAL_PROJECTS_FILE, index=False)

                    # Upload to Dropbox
                    upload_to_dropbox(LOCAL_PROJECTS_FILE, DROPBOX_PROJECTS_PATH, ACCESS_TOKEN)
                    st.success(f"Project '{project_name}' submitted successfully!")
                except Exception as e:
                    st.error(f"Error submitting project: {e}")

    elif action == "Update Existing Project":
        st.subheader("Step 3: Update an Existing Project")
        try:
            projects_data = pd.read_excel(LOCAL_PROJECTS_FILE)
            project_ids = projects_data["Project ID"].unique().tolist()
            selected_project_id = st.selectbox("Select Project ID", project_ids)

            if selected_project_id:
                project_data = projects_data[projects_data["Project ID"] == selected_project_id]
                st.dataframe(project_data)

                updated_data = []
                for _, row in project_data.iterrows():
                    st.markdown(f"**Personnel: {row['Personnel']} | Week: {row['Week']}**")
                    budgeted_hours = st.number_input(
                        f"Budgeted Hours", min_value=0, value=int(row["Budgeted Hrs"]), step=1, key=f"{row['Week']}_budgeted"
                    )
                    spent_hours = st.number_input(
                        f"Spent Hours", min_value=0, value=int(row["Spent Hrs"]), step=1, key=f"{row['Week']}_spent"
                    )
                    updated_data.append({
                        **row,
                        "Budgeted Hrs": budgeted_hours,
                        "Spent Hrs": spent_hours,
                    })

                if st.button("Save Updates"):
                    remaining_data = projects_data[projects_data["Project ID"] != selected_project_id]
                    final_data = pd.concat([remaining_data, pd.DataFrame(updated_data)], ignore_index=True)
                    final_data.to_excel(LOCAL_PROJECTS_FILE, index=False)
                    upload_to_dropbox(LOCAL_PROJECTS_FILE, DROPBOX_PROJECTS_PATH, ACCESS_TOKEN)
                    st.success(f"Project '{selected_project_id}' updated successfully!")

        except Exception as e:
            st.error(f"Error loading projects: {e}")

if __name__ == "__main__":
    main()
