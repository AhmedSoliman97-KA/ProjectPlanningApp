import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import dropbox
import os

# Dropbox Access Token and Folder Path
ACCESS_TOKEN = "sl.CCbBAzoF5cwNWFQ_Tu_FMjqrfzi_6gxHrjcPrlxvQmTuZMmGS5uBmDE7RQV8-NtwT3BwJRaQlQmuOnQtK54mZ16YX7bhpR7ufM4Y1lE7ZwkvgptEqIQQTt6RwBt4XrgOuAF1GO2gK7KpHSiR2eR9"
DROPBOX_FOLDER_PATH = "/Project_Data"

# Dropbox Functions
def upload_to_dropbox(file_path, dropbox_path, access_token):
    """Upload a file to Dropbox."""
    try:
        dbx = dropbox.Dropbox(access_token)
        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
        print(f"Uploaded to Dropbox: {dropbox_path}")
    except dropbox.exceptions.AuthError:
        print("Authentication error: Check your ACCESS_TOKEN.")
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
        print(f"Error downloading file: {e}")

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

    # Project Details
    st.subheader("Step 1: Enter Project Details")
    project_id = st.text_input("Project ID")
    project_name = st.text_input("Project Name")
    selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
    selected_month = st.selectbox("Month", range(1, 13))
    weeks = [f"Week {i}" for i in range(1, 5)]

    st.subheader("Step 2: Assign Engineers and Weekly Hours")
    engineer_names = ["Alice", "Bob", "Charlie"]
    selected_engineers = st.multiselect("Select Engineers", engineer_names)

    # Initialize Data Storage
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
        st.subheader("Current Data")
        summary_data = pd.DataFrame(st.session_state.engineer_allocation.values())
        st.dataframe(summary_data)

        if st.button("Submit Project"):
            try:
                # Load existing data
                try:
                    existing_data = pd.read_excel(local_projects_path)
                except FileNotFoundError:
                    existing_data = pd.DataFrame(columns=summary_data.columns)

                # Append and save
                final_data = pd.concat([existing_data, summary_data], ignore_index=True)
                final_data.to_excel(local_projects_path, index=False)

                # Upload to Dropbox
                upload_to_dropbox(local_projects_path, dropbox_projects_path, ACCESS_TOKEN)
                st.success(f"Project '{project_name}' submitted!")
            except Exception as e:
                st.error(f"Error submitting project: {e}")

    # Download Button
    st.subheader("Download Latest File")
    if st.button("Download Latest File"):
        try:
            download_from_dropbox(dropbox_projects_path, local_projects_path, ACCESS_TOKEN)
            with open(local_projects_path, "rb") as file:
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
