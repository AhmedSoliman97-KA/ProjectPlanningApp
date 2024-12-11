import streamlit as st
import pandas as pd
import dropbox
from datetime import datetime

# Dropbox Access Token (replace with your token)
ACCESS_TOKEN = "sl.u.AFbSi0u1JEt41HfRfC_wgf5aX0jwZCocifSqL6t8JaJvCL1Q86YPZMr9VcQevwbYrsqY87MzcGxefFW3kL_acbhKoOuqZ2dwrNKTP-2HhBFJdZm9NZWgrLgqD_qMBJBlbCF7IjI0nRNeUzgm44Q5fBZrUGFvYbDznCb_E3NdsYUpdm5vML2362RLKifpuRK2Sh0K6da6MbyysftZCf11oEaAHS-jY45uP21ps5geNPm5JRSTffJ0Z2zj_GZ1AWpSwJCFPsLXYJRmYxLwTNh8OxFSCgQU_gRCrG3ymr99wWNNN-njybj47lR27dCmPKj6jhhQT14Tu7nG9E-w8Jt9OywK3msYI18tCO_SZ21A3GqOFwdVghSazBn6s3uXzeDN8qmNteTSi1RWkHIOyTa_ItYkkWbY7xts_mIkWQpUhdsxb9PfA-jJQtEy7yt2cMD-sEekDJhHBtejqHIVNSBA1QJbA7Dc_lmu34i7Jf_1CxPfqZdsWdrbqtQbl6VAFS8xR0FtBtnf4il1FHsEV4eyoc93OLGeWYf-ikLE_KPS5vwXyHxXxWO3DNWu3tVu8BmXiH26B7BYVkRaD6PP-ANRtqS4Vi27Ve21N0GMopBqEyiwz63lck8Fxu2NgdWvPntYqcIjjgHyQVGsP-jBXgFvja-KJHMzHC3V9hMiQ4CPJMnnB9kFU9nHWVLpHxhQvMKM99XW8hqgIzl_RMXDajmvxeXG2fvTrT3ktEl8vXRrqg2maIN3-zq5ef1dhelmkfGvV7I7nI0Zv3stGW3F0WYJFX33hrCEl7MdToABlpPqlMA0b80njb5nF_YoAuayKWUCemaIYtAOn5PnsIzbE6tshtb0iIVldF4vStfgjIG6FuOh7IiYaU1FrIWHGBhs0yey_-IDGt_RAmuHRTp2IW_Jt5ByxrR748H9JignxTmGSfJUKQuaX6vSkzThIoJ6elkcj6JFXOphbJ33_DwuPVm1cXfk9NGz8Fagb56TQdb8MtDaP1wqg2ERur_PIOv3vuHHsKqrdm5Y4bgXryrwEUDKf-mfXhDivk1ATGUfghfgVnIBLznpMgoWoML_TCnuUaPrfa5Ajyo2goHHNeuZQ8nzHB4AMo6d7Tb_HxB9AVkwddoJ2O-sQm-ks0rv5VpdiiApjEy7Hz9YFwgqC7bMf-Ex5BMiFmVyrrvT_BinF5KWuXS_dORZsjdQwwPZsNr2TpX2rIZ76dvJAtPWSqiVo6gSouAnKFlrNmTc5JrHPBTV_GH4H81lYGLVj2HycAY7c3xD_wQ37HVs_R3F7BGcVWoEUm1fzgD3Darj3oQz4MS75hZGEDP1YKcmuM3s6J6O6kSrv_fwcF3oPilsRWBQYYwKiVeJLamanRdxxnFndMEPhZ0s5n5EuXg6_v3rsm0FOmTyVNKTOJ1cYAZ5Wo5KHkhpDRPMX4qkrDdV4_F7v59tRDd5iA"

# Dropbox Functions
def save_to_dropbox(file_path, file_name, access_token):
    """Uploads a file to Dropbox."""
    dbx = dropbox.Dropbox(access_token)
    with open(file_path, "rb") as f:
        dbx.files_upload(f.read(), f"/{file_name}", mode=dropbox.files.WriteMode("overwrite"))
    print(f"File {file_name} uploaded to Dropbox.")

def download_from_dropbox(file_name, access_token):
    """Downloads a file from Dropbox."""
    dbx = dropbox.Dropbox(access_token)
    metadata, res = dbx.files_download(f"/{file_name}")
    with open(file_name, "wb") as f:
        f.write(res.content)
    print(f"File {file_name} downloaded from Dropbox.")

# Load Engineer Data from Human Resources File
def load_engineers(file_path):
    """Loads engineer names and sections from Human Resources.xlsx."""
    try:
        excel_file = pd.ExcelFile(file_path)
        sections = excel_file.sheet_names  # Get all sheet names (sections)
        return {section: excel_file.parse(sheet_name=section)["Name"].dropna().tolist() for section in sections}
    except Exception as e:
        st.error(f"Error loading Human Resources file: {e}")
        return {}

# Save Project Data Function
def save_project_data(dataframe):
    """Save the project data locally and upload to Dropbox."""
    file_path = "Project_Data.xlsx"
    file_name = "Project_Data.xlsx"

    # Save to local file
    dataframe.to_excel(file_path, index=False)

    # Upload to Dropbox
    save_to_dropbox(file_path, file_name, ACCESS_TOKEN)

# Streamlit App
def main():
    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    # Load Human Resources Data
    human_resources_path = "Human Resources.xlsx"
    engineers_by_section = load_engineers(human_resources_path)

    if not engineers_by_section:
        st.error("Could not load engineer data. Ensure Human Resources.xlsx is in the directory.")
        return

    # Section Selection
    st.subheader("Step 1: Select Section")
    selected_section = st.selectbox("Choose a section", list(engineers_by_section.keys()))

    # Filter Engineers
    engineers = engineers_by_section.get(selected_section, [])

    if not engineers:
        st.warning("No engineers found in the selected section.")
        return

    # Project Details
    st.subheader("Step 2: Enter Project Details")
    project_id = st.text_input("Project ID", help="Enter the unique ID for the project.")
    project_name = st.text_input("Project Name", help="Enter the name of the project.")

    # Engineer and Hours Input
    st.subheader("Step 3: Assign Engineers and Hours")
    selected_engineers = st.multiselect("Select Engineers", engineers)
    week = st.selectbox("Select Week", [f"Week {i}" for i in range(1, 53)])
    hours = st.number_input("Hours", min_value=0, step=1, help="Enter hours for the selected week.")

    # Initialize Data Storage
    if "project_data" not in st.session_state:
        st.session_state.project_data = []

    # Add Data Button
    if st.button("Add Hours"):
        for engineer in selected_engineers:
            st.session_state.project_data.append({
                "Project ID": project_id,
                "Project Name": project_name,
                "Engineer": engineer,
                "Week": week,
                "Hours": hours,
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Section": selected_section,
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
            save_project_data(data)
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
