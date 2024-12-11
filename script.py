import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import dropbox
import os

# Dropbox Access Token and Paths
ACCESS_TOKEN = "sl.u.AFYOnrUv7z6a7Ju2hl4nImVS7iBr1hpkEnaOnE7nm25My9O1XC_Nc3wwpmfOq37TKTczK1yJxUI184QuJgAYRaXZwqFK5J2cI1Ww9olBRaY6CTaBwoQlha8V5N-v7I1BUqHOtiIwIU8qteX4alDSS22ISDAcLpqJqiun4WfGgOz0MdNjLDc3661Tp5MwxfsvunsFXU7W5JeWQ9H9f8FBFaZH3Qg8zsHcoejykPOMt9jeMqgRXtTr5AYAEOqJyhWqsiOY5kHYnKGVimgA76K75tiDtOV37qQGqiYq71nBWn2zQdpaOPZZnF-Zj2EVy0v4AP66p8ZQljlvwftL1ExsRQKSF7LQy_RHojIezE3ds0BLXmlGkPnZOwV6zS0VgwzPiAUOYSEXdBT6j8A6LnOzGkQaofFIRFr8WPeXlg91vHxSP1nFeFAdnhJ7bRJlB5YMtywCTh9K6PDeB2_HwG1-bY84w0SrxhTIkCzQJBZ_GnJWxnjikIJFddk1k1jsYZMA_ITRDawf1oGeNYhCG-dQ11EJ44fuRo2SLhKbzYpvzhEXeGX5R_DFCTsMXq_vNNbsqOZMMPhL2Jv9oeBs5VfW8SIGXd98JNtPpr4A0cDrNU17lup_sUNT9b_M6aPpjRsLYDzl3Hb-7DzEiFAePXjKi73gMqm-IyDNrN-1qyNFiAKBo9SKWWSXPnPh7ismrRnswJso8Z2aVMSE9QAqetm-Oe5JGZqSZ68NnaL0sB7ig5KJhZVVYImrvP6GA_Ulvg8YUn_ejdS1kddRscIkjXL-F0FPfTbpvPiWMRCNLZ3Bkdx-hhW2fUZqq_MyssfpyTvoKtBI4VW3a7_3KM6ifzTtZIkJycyDv7H5VInydYOcpuUmv9kamkmbYWLeC8pJs-3UAuqq35Q6GAnOPqaNoVg8iHzUoMdDKc5pVnO94QWDJzAFXkEfRuZFt6xQh6hnb0Stx7mPJL6xXUHfMn82FnGZRltz7RIBWrp2_HHF3PxLM020i0lKY66zYZgMZb3mzdtg9jIQyETJzV6oYOjaqb4YVYBUhKN-Vp3vgANvqCaFxHtaeH7I6Blpawv4FI1g3BwxNiqQq5Y2HcO7ulPLBy3lT3XIvXyiNFr-q9RcDKMofYxYD-SzI1BfYBEp12k43X0tZ1-xkAoh0XvqjQ33WlsyhLncx87fWggVTJbxoxsfcYXi0vt1NwN4Q3PYMbZhHAfYudRuDfKI4SHsYiDlX8FZU9WpF5XhjEDLI-D_5j1vdbLT5JKwVkNmjcEYo-8AqLSVaAcCi6nE3Uwk40ZX85WBxGaRW8P4LO6nXs5dkxDICbrlcnSKZaib345ieiXybDKrem9OoU_ZF01_5onqhfXyn1nhbK9bhVRZQZJKAhgLoggefdQQtwuNTufR5V8nl4MIS0WdQe7o_frDNK_qaeRhtC80LO-IAzKCEfG_bUpzubsVnQ"
DROPBOX_FOLDER_PATH = "/Project_Data"
HR_FILE_NAME = "Human Resources.xlsx"
LOCAL_HR_FILE = HR_FILE_NAME
DROPBOX_HR_PATH = f"{DROPBOX_FOLDER_PATH}/{HR_FILE_NAME}"

# Dropbox Functions
def upload_to_dropbox(file_path, dropbox_path, access_token):
    """Upload a file to Dropbox."""
    try:
        dbx = dropbox.Dropbox(access_token)
        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
        print(f"Uploaded to Dropbox: {dropbox_path}")
    except dropbox.exceptions.AuthError:
        print("Authentication error: Check your access token.")
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
    except dropbox.exceptions.ApiError:
        print(f"Dropbox file not found: {dropbox_path}")
    except Exception as e:
        print(f"Error downloading file: {e}")

# Ensure File Exists or Create One
def ensure_file_exists(file_path, dropbox_path, columns=None):
    """Ensure the file exists locally or in Dropbox. Create it if not."""
    if not os.path.exists(file_path):
        try:
            # Attempt to download from Dropbox
            download_from_dropbox(dropbox_path, file_path, ACCESS_TOKEN)
        except Exception:
            # Create a new file if Dropbox download fails
            if columns:
                pd.DataFrame(columns=columns).to_excel(file_path, index=False)
                upload_to_dropbox(file_path, dropbox_path, ACCESS_TOKEN)
                print(f"File created and uploaded: {file_path}")

# Load Engineers Data from Human Resources File
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

    # File paths for project data
    local_projects_path = "projects_data_weekly.xlsx"
    dropbox_projects_path = f"{DROPBOX_FOLDER_PATH}/projects_data_weekly.xlsx"

    # Ensure project data file exists
    ensure_file_exists(
        local_projects_path,
        dropbox_projects_path,
        ["Project ID", "Project Name", "Personnel", "Week", "Year", "Month", "Budgeted Hrs", "Spent Hrs"]
    )

    # Ensure Human Resources file exists locally
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
    project_id = st.text_input("Project ID")
    project_name = st.text_input("Project Name")
    selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
    selected_month = st.selectbox("Month", range(1, 13))
    weeks = [f"Week {i}" for i in range(1, 5)]

    st.subheader("Step 4: Assign Engineers and Weekly Hours")
    selected_engineers = st.multiselect("Select Engineers", engineers)

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
