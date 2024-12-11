import streamlit as st
import pandas as pd
from datetime import datetime
import dropbox
import os

# Dropbox Access Token and Paths
ACCESS_TOKEN = "sl.CCZ5A6jITRh7Hxzl_mQ2IwVGcgzjmZCVsnPf10S50iSQDBGp-aGyUVWRce98IEcMwoQQnfoY2oh1M0ehD2XsYXyIAvP0pTm-ZXBa7uMQBIPRInyfp81VtCkAg_NTvKxHQ89RWFV5qUHJKM0EUXQdXoQ"
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

    # Download Latest File Button
    st.subheader("Download Latest File")
    if st.button("Download Latest File"):
        try:
            download_from_dropbox(DROPBOX_PROJECTS_PATH, LOCAL_PROJECTS_FILE, ACCESS_TOKEN)
            with open(LOCAL_PROJECTS_FILE, "rb") as file:
                st.download_button(
                    label="Download Project Data",
                    data=file,
                    file_name="projects_data_weekly.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error downloading file: {e}")

if __name__ == "__main__":
    main()
