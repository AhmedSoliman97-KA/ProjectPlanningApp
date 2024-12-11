import streamlit as st
import pandas as pd
import dropbox
from datetime import datetime
import os

# Dropbox Access Token (replace with your token)
ACCESS_TOKEN = "sl.u.AFbSi0u1JEt41HfRfC_wgf5aX0jwZCocifSqL6t8JaJvCL1Q86YPZMr9VcQevwbYrsqY87MzcGxefFW3kL_acbhKoOuqZ2dwrNKTP-2HhBFJdZm9NZWgrLgqD_qMBJBlbCF7IjI0nRNeUzgm44Q5fBZrUGFvYbDznCb_E3NdsYUpdm5vML2362RLKifpuRK2Sh0K6da6MbyysftZCf11oEaAHS-jY45uP21ps5geNPm5JRSTffJ0Z2zj_GZ1AWpSwJCFPsLXYJRmYxLwTNh8OxFSCgQU_gRCrG3ymr99wWNNN-njybj47lR27dCmPKj6jhhQT14Tu7nG9E-w8Jt9OywK3msYI18tCO_SZ21A3GqOFwdVghSazBn6s3uXzeDN8qmNteTSi1RWkHIOyTa_ItYkkWbY7xts_mIkWQpUhdsxb9PfA-jJQtEy7yt2cMD-sEekDJhHBtejqHIVNSBA1QJbA7Dc_lmu34i7Jf_1CxPfqZdsWdrbqtQbl6VAFS8xR0FtBtnf4il1FHsEV4eyoc93OLGeWYf-ikLE_KPS5vwXyHxXxWO3DNWu3tVu8BmXiH26B7BYVkRaD6PP-ANRtqS4Vi27Ve21N0GMopBqEyiwz63lck8Fxu2NgdWvPntYqcIjjgHyQVGsP-jBXgFvja-KJHMzHC3V9hMiQ4CPJMnnB9kFU9nHWVLpHxhQvMKM99XW8hqgIzl_RMXDajmvxeXG2fvTrT3ktEl8vXRrqg2maIN3-zq5ef1dhelmkfGvV7I7nI0Zv3stGW3F0WYJFX33hrCEl7MdToABlpPqlMA0b80njb5nF_YoAuayKWUCemaIYtAOn5PnsIzbE6tshtb0iIVldF4vStfgjIG6FuOh7IiYaU1FrIWHGBhs0yey_-IDGt_RAmuHRTp2IW_Jt5ByxrR748H9JignxTmGSfJUKQuaX6vSkzThIoJ6elkcj6JFXOphbJ33_DwuPVm1cXfk9NGz8Fagb56TQdb8MtDaP1wqg2ERur_PIOv3vuHHsKqrdm5Y4bgXryrwEUDKf-mfXhDivk1ATGUfghfgVnIBLznpMgoWoML_TCnuUaPrfa5Ajyo2goHHNeuZQ8nzHB4AMo6d7Tb_HxB9AVkwddoJ2O-sQm-ks0rv5VpdiiApjEy7Hz9YFwgqC7bMf-Ex5BMiFmVyrrvT_BinF5KWuXS_dORZsjdQwwPZsNr2TpX2rIZ76dvJAtPWSqiVo6gSouAnKFlrNmTc5JrHPBTV_GH4H81lYGLVj2HycAY7c3xD_wQ37HVs_R3F7BGcVWoEUm1fzgD3Darj3oQz4MS75hZGEDP1YKcmuM3s6J6O6kSrv_fwcF3oPilsRWBQYYwKiVeJLamanRdxxnFndMEPhZ0s5n5EuXg6_v3rsm0FOmTyVNKTOJ1cYAZ5Wo5KHkhpDRPMX4qkrDdV4_F7v59tRDd5iA"

# Dropbox Functions
def save_to_dropbox(file_path, file_name, access_token):
    """Uploads a file to Dropbox."""
    try:
        dbx = dropbox.Dropbox(access_token)
        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), f"/{file_name}", mode=dropbox.files.WriteMode("overwrite"))
        print(f"File {file_name} uploaded to Dropbox.")
    except dropbox.exceptions.ApiError as e:
        print(f"Dropbox API error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

def download_from_dropbox(file_name, access_token):
    """Downloads a file from Dropbox."""
    try:
        dbx = dropbox.Dropbox(access_token)
        metadata, res = dbx.files_download(f"/{file_name}")
        with open(file_name, "wb") as f:
            f.write(res.content)
        print(f"File {file_name} downloaded from Dropbox.")
    except Exception as e:
        print(f"Error downloading file: {e}")

# Ensure File Exists or Create One
def ensure_file_exists(file_path, columns):
    """Ensure the file exists. Create it if it doesn't."""
    if not os.path.exists(file_path):
        df = pd.DataFrame(columns=columns)
        df.to_excel(file_path, index=False)
        print(f"File created: {file_path}")

# Save Project Data Function
def save_project_data(dataframe, file_path):
    """Save the project data locally and upload to Dropbox."""
    # Ensure the file exists
    ensure_file_exists(file_path, dataframe.columns)

    # Load existing data if the file exists
    try:
        existing_data = pd.read_excel(file_path)
    except FileNotFoundError:
        existing_data = pd.DataFrame(columns=dataframe.columns)

    # Append new data and save
    combined_data = pd.concat([existing_data, dataframe], ignore_index=True)
    combined_data.to_excel(file_path, index=False)

    # Upload to Dropbox
    save_to_dropbox(file_path, os.path.basename(file_path), ACCESS_TOKEN)

# Streamlit App
def main():
    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    # File path for storage
    local_file_path = "Project_Data.xlsx"

    # Project Details
    st.subheader("Step 1: Enter Project Details")
    project_id = st.text_input("Project ID", help="Enter the unique ID for the project.")
    project_name = st.text_input("Project Name", help="Enter the name of the project.")

    # Engineer and Hours Input
    st.subheader("Step 2: Assign Engineers and Hours")
    engineers = st.multiselect("Select Engineers", ["Alice", "Bob", "Charlie"])
    week = st.selectbox("Select Week", [f"Week {i}" for i in range(1, 53)])
    hours = st.number_input("Hours", min_value=0, step=1, help="Enter hours for the selected week.")

    # Initialize Data Storage
    if "project_data" not in st.session_state:
        st.session_state.project_data = []

    # Add Data Button
    if st.button("Add Hours"):
        for engineer in engineers:
            st.session_state.project_data.append({
                "Project ID": project_id,
                "Project Name": project_name,
                "Engineer": engineer,
                "Week": week,
                "Hours": hours,
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
        st.success("Hours added successfully!")

    # Display Current Data
    if st.session_state.project_data:
        st.subheader("Current Data")
        data = pd.DataFrame(st.session_state.project_data)
        st.dataframe(data)

    # Submit Project Button
    if st.button("Submit Project"):
        if st.session_state.project_data:
            # Save data to Dropbox
            data = pd.DataFrame(st.session_state.project_data)
            save_project_data(data, local_file_path)
            st.success("Project data saved and uploaded to Dropbox!")
        else:
            st.error("No data to submit!")

    # Download Button
    st.subheader("Download Latest File")
    if st.button("Download Latest File"):
        try:
            download_from_dropbox("Project_Data.xlsx", ACCESS_TOKEN)
            with open("Project_Data.xlsx", "rb") as file:
                st.download_button(
                    label="Download Project Data",
                    data=file,
                    file_name="Project_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error downloading file: {e}")

if __name__ == "__main__":
    main()
