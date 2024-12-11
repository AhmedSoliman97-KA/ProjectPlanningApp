import streamlit as st
import pandas as pd
from datetime import datetime
import dropbox
import os

# Dropbox Access Token and Paths
ACCESS_TOKEN = "sl.u.AFZfAbFllA29neJgjpRx63DedZKVU5jYc9_ze4-Y4HhSdTfnCW3s6NqQVy-8JwH5sa0rS1SKhdJUeFdRk6vVBPF3brz0sjQmPX7W-H_7lSoZj2224z8wzUWgYmHA0tIS6V9eRd34rj-bXqUSQ4zzrFflowgqXxFbt12e6bxsdkzSVSvXO8f7G1MRBQC3lln3XFbYYMRpJwnjO1nNs3HrPuo2tToG7VuQps4yPrilW1XaucBx5g2kDM83VmW0VSfIb24RO2XxbrJB_x11XXigsy1ES0hV0J7vVW6tD1mrg13wmwx92imzWePsiskx0lMKQ5bHGnI3_XfqR1YW18rt8nHgg983LJQ24Iu_iMiWiX6gcyMMaB5iv6DvFjF2AGSvuB-oe-yzXWq35ATuvhGxL1iBTNyKZGAcY7fgNBB8Km__d61aY-wttgWNG35Dp0VdLfSdDiZT-r_3M1MA9uTK7yYhLKmJg-MvPsVLp1Vh_AF7HRCiYJDulcTCzTNyvVMDlEFac5NidFL4shCOLXP3oMKT4_z0HcdxSnwQ0k9-uFSg1BV8XQaP7u0UiE_16fLtvKZaa9H2YFO6FiN14rYt6czeTb1Jqnv9UqJ-P1RLSEma2BZ-kL04ntXBdwsOPReHxAKD2TTpbDQZk8nXXwyeC-l1Ho2FMaR2lx7lp8QN1PNCTjSFS794uPUnGH_OuDwkSKI6uP7KfSgQFgl-BNXhVkdjHuzyomLnrKuKu-S53UNTf74p-FOYO_RT7cC_rtI3g3TD8MlIEpNfCyT_IQaXI3YFparHkpK27XjVlO1sJ7Bf6Hxg0gGnLMzdJl2HBXrAdr684QFpYscWagvTBX-UoofSHnotdhIMc65fHaSH3mYlerKwYx_IafYnWvhN5L4lGnMhlK4Kp6tj7_6fPqEJN9PtsPO1Ob3YFdN4IMovTGB-Z4kcDq3EsQZMBIITaFbIhWgEjnSVGruXPB0IM0m1nZxTGxbbNB8F02Dy2P3uwZqNRXL4mwSlg5H-jUtGQw4opTeAKRp2lycSNHlp8lLVDrekXHmvOyg7VQhgAVEfm6n51NX3KlUjSkKva3N3QySFiwUhMpMBUX8eRyGuMvSpcfIMKvmJc0vAhIagiG50u9LjPixrPIQp0KY8-z6wOikBRzA6CjefibaS5Cw4BA1yvGrEeSQPvGg6ADOv1HgDeoHuMk78ZT1daAyvfCCHGWiBHQ1kmUw6V0egKDiO-92FEOZwxxwOJKlyhxJ7T5NrnFVKcT5Tu9GBW-CeaHN-MN4EhJFU7LM5ej6mNrBYCVKa1L_s02g6OjfaRb8EZFDgcBIOvij85rYoZNDr-x4jlPbCDXYuYDSYLjstTUa7WQidObkHh-JGnHfoCyJPuJusk8rvsqUwzKVSHM-VTrZgUjUE39u9BSi98HWQWaRMhvsksaR1z_rJ2WJWDMo7EZNHf3QXXg"
DROPBOX_FOLDER_PATH = "/Project_Data"
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
        st.warning("File not found! Initializing a new file...")
        if columns:
            pd.DataFrame(columns=columns).to_excel(file_path, index=False)
            upload_to_dropbox(file_path, dropbox_path, ACCESS_TOKEN)
            print(f"New file created and uploaded: {file_path}")

# Clear All Session Data
def clear_all_session_data():
    """Clear all session data."""
    for key in list(st.session_state.keys()):
        del st.session_state[key]

# Main Application
def main():
    # Ensure file initialization
    ensure_file_exists(
        LOCAL_PROJECTS_FILE,
        DROPBOX_PROJECTS_PATH,
        columns=["Project ID", "Project Name", "Personnel", "Week", "Year", "Month", "Budgeted Hrs", "Spent Hrs"]
    )

    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    # Step 1: Enter Project Details
    st.subheader("Enter Project Details")
    project_id = st.text_input("Project ID", help="Enter a unique ID for the project.")
    project_name = st.text_input("Project Name")
    selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
    selected_month = st.selectbox("Month", [datetime(1900, m, 1).strftime("%B") for m in range(1, 13)])
    weeks = [f"Week {i}" for i in range(1, 5)]
    personnel = st.text_input("Personnel")
    budgeted_hours = st.number_input("Budgeted Hours", min_value=0, step=1)
    spent_hours = st.number_input("Spent Hours", min_value=0, step=1)

    # Display and manage existing data
    if "project_data" not in st.session_state:
        try:
            st.session_state.project_data = pd.read_excel(LOCAL_PROJECTS_FILE)
        except FileNotFoundError:
            st.session_state.project_data = pd.DataFrame(columns=["Project ID", "Project Name", "Personnel", "Week", "Year", "Month", "Budgeted Hrs", "Spent Hrs"])

    if st.session_state.project_data.empty:
        st.warning("No data available yet.")
    else:
        st.subheader("Existing Project Data")
        st.dataframe(st.session_state.project_data)

    # Add a new record
    if st.button("Add Allocation"):
        new_entry = {
            "Project ID": project_id,
            "Project Name": project_name,
            "Personnel": personnel,
            "Week": weeks[0],  # Default week for simplicity
            "Year": selected_year,
            "Month": selected_month,
            "Budgeted Hrs": budgeted_hours,
            "Spent Hrs": spent_hours,
        }
        st.session_state.project_data = pd.concat(
            [st.session_state.project_data, pd.DataFrame([new_entry])], ignore_index=True
        )
        st.success("Allocation added.")

    # Submit all data
    if st.button("Submit Project"):
        try:
            st.session_state.project_data.to_excel(LOCAL_PROJECTS_FILE, index=False)
            upload_to_dropbox(LOCAL_PROJECTS_FILE, DROPBOX_PROJECTS_PATH, ACCESS_TOKEN)
            st.success("Project submitted successfully!")
            clear_all_session_data()  # Clear session data
        except Exception as e:
            st.error(f"Error submitting project: {e}")

    # Download latest file
    if st.button("Download Latest File"):
        with open(LOCAL_PROJECTS_FILE, "rb") as file:
            st.download_button(
                label="Download Project File",
                data=file,
                file_name="projects_data_weekly.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
