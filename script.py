import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox

# Dropbox Settings
ACCESS_TOKEN = "sl.u.AFZfAbFllA29neJgjpRx63DedZKVU5jYc9_ze4-Y4HhSdTfnCW3s6NqQVy-8JwH5sa0rS1SKhdJUeFdRk6vVBPF3brz0sjQmPX7W-H_7lSoZj2224z8wzUWgYmHA0tIS6V9eRd34rj-bXqUSQ4zzrFflowgqXxFbt12e6bxsdkzSVSvXO8f7G1MRBQC3lln3XFbYYMRpJwnjO1nNs3HrPuo2tToG7VuQps4yPrilW1XaucBx5g2kDM83VmW0VSfIb24RO2XxbrJB_x11XXigsy1ES0hV0J7vVW6tD1mrg13wmwx92imzWePsiskx0lMKQ5bHGnI3_XfqR1YW18rt8nHgg983LJQ24Iu_iMiWiX6gcyMMaB5iv6DvFjF2AGSvuB-oe-yzXWq35ATuvhGxL1iBTNyKZGAcY7fgNBB8Km__d61aY-wttgWNG35Dp0VdLfSdDiZT-r_3M1MA9uTK7yYhLKmJg-MvPsVLp1Vh_AF7HRCiYJDulcTCzTNyvVMDlEFac5NidFL4shCOLXP3oMKT4_z0HcdxSnwQ0k9-uFSg1BV8XQaP7u0UiE_16fLtvKZaa9H2YFO6FiN14rYt6czeTb1Jqnv9UqJ-P1RLSEma2BZ-kL04ntXBdwsOPReHxAKD2TTpbDQZk8nXXwyeC-l1Ho2FMaR2lx7lp8QN1PNCTjSFS794uPUnGH_OuDwkSKI6uP7KfSgQFgl-BNXhVkdjHuzyomLnrKuKu-S53UNTf74p-FOYO_RT7cC_rtI3g3TD8MlIEpNfCyT_IQaXI3YFparHkpK27XjVlO1sJ7Bf6Hxg0gGnLMzdJl2HBXrAdr684QFpYscWagvTBX-UoofSHnotdhIMc65fHaSH3mYlerKwYx_IafYnWvhN5L4lGnMhlK4Kp6tj7_6fPqEJN9PtsPO1Ob3YFdN4IMovTGB-Z4kcDq3EsQZMBIITaFbIhWgEjnSVGruXPB0IM0m1nZxTGxbbNB8F02Dy2P3uwZqNRXL4mwSlg5H-jUtGQw4opTeAKRp2lycSNHlp8lLVDrekXHmvOyg7VQhgAVEfm6n51NX3KlUjSkKva3N3QySFiwUhMpMBUX8eRyGuMvSpcfIMKvmJc0vAhIagiG50u9LjPixrPIQp0KY8-z6wOikBRzA6CjefibaS5Cw4BA1yvGrEeSQPvGg6ADOv1HgDeoHuMk78ZT1daAyvfCCHGWiBHQ1kmUw6V0egKDiO-92FEOZwxxwOJKlyhxJ7T5NrnFVKcT5Tu9GBW-CeaHN-MN4EhJFU7LM5ej6mNrBYCVKa1L_s02g6OjfaRb8EZFDgcBIOvij85rYoZNDr-x4jlPbCDXYuYDSYLjstTUa7WQidObkHh-JGnHfoCyJPuJusk8rvsqUwzKVSHM-VTrZgUjUE39u9BSi98HWQWaRMhvsksaR1z_rJ2WJWDMo7EZNHf3QXXg"
PROJECTS_FILE_PATH = "/Project_Data/projects_data_weekly.xlsx"
HR_FILE_LOCAL = "Human Resources.xlsx"

# Dropbox Functions
def download_from_dropbox(file_path):
    """Download a file from Dropbox."""
    try:
        dbx = dropbox.Dropbox(ACCESS_TOKEN)
        metadata, res = dbx.files_download(file_path)
        return pd.ExcelFile(res.content)
    except dropbox.exceptions.ApiError as e:
        if e.error.is_path() and e.error.get_path().is_not_found():
            return None
        else:
            st.error(f"Error downloading file: {e}")
            return None

def upload_to_dropbox(df, dropbox_path):
    """Upload a DataFrame to Dropbox as an Excel file."""
    try:
        dbx = dropbox.Dropbox(ACCESS_TOKEN)
        with pd.ExcelWriter("temp.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        with open("temp.xlsx", "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
        st.success(f"Data successfully uploaded to {dropbox_path}.")
    except Exception as e:
        st.error(f"Error uploading data: {e}")

def ensure_dropbox_projects_file_exists(file_path):
    """Ensure the projects file exists in Dropbox, create if not."""
    existing_file = download_from_dropbox(file_path)
    if existing_file is None:
        st.warning(f"{file_path} not found in Dropbox. Creating a new file...")
        empty_df = pd.DataFrame(columns=[
            "Project ID", "Project Name", "Personnel", "Week", "Year", "Month",
            "Budgeted Hrs", "Spent Hrs", "Remaining Hrs", "Cost/Hour", "Budgeted Cost",
            "Remaining Cost", "Section", "Category"
        ])
        upload_to_dropbox(empty_df, file_path)

# Main Application
def main():
    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    # Ensure the projects file exists in Dropbox
    ensure_dropbox_projects_file_exists(PROJECTS_FILE_PATH)

    # Load Human Resources File
    try:
        hr_excel = pd.ExcelFile(HR_FILE_LOCAL)
    except Exception as e:
        st.error(f"Error loading Human Resources file: {e}")
        st.stop()

    # Load Projects Data from Dropbox
    projects_excel = download_from_dropbox(PROJECTS_FILE_PATH)
    if projects_excel is None:
        projects_data = pd.DataFrame(columns=[
            "Project ID", "Project Name", "Personnel", "Week", "Year", "Month",
            "Budgeted Hrs", "Spent Hrs", "Remaining Hrs", "Cost/Hour", "Budgeted Cost",
            "Remaining Cost", "Section", "Category"
        ])
    else:
        projects_data = pd.read_excel(projects_excel)

    # Action Selection
    st.sidebar.subheader("Actions")
    action = st.sidebar.radio("Choose an Action", ["Create New Project", "Update Existing Project"])

    if action == "Create New Project":
        st.subheader("Create a New Project")
        project_id = st.text_input("Project ID", help="Enter a unique ID for the project.")
        project_name = st.text_input("Project Name", help="Enter the name of the project.")
        st.write("Create project logic is untouched.")

    elif action == "Update Existing Project":
        st.subheader("Update an Existing Project")
        if projects_data.empty:
            st.warning("No existing projects found.")
            st.stop()

        # Step 1: Filter by Section
        st.subheader("Filter by Section")
        unique_sections = projects_data["Section"].dropna().str.strip().str.lower().unique().tolist()
        selected_section = st.selectbox("Choose a Section", unique_sections)

        # Normalize sections for comparison
        selected_section_normalized = selected_section.strip().lower()
        projects_data["Section Normalized"] = projects_data["Section"].str.strip().str.lower()

        # Step 2: Filter Projects by Section
        filtered_projects = projects_data[projects_data["Section Normalized"] == selected_section_normalized]
        if filtered_projects.empty:
            st.warning(f"No projects found for the selected section: {selected_section}")
            st.stop()

        # Step 3: Select Project
        st.subheader("Select a Project")
        selected_project = st.selectbox("Select a Project", filtered_projects["Project Name"].unique())
        project_details = filtered_projects[filtered_projects["Project Name"] == selected_project]

        # Display Project Details
        st.subheader(f"Summary of Current Allocations for '{selected_project}'")
        st.dataframe(project_details)

        # Step 4: Engineer Dropdown
        selected_engineer = st.selectbox("Select an Engineer", project_details["Personnel"].unique())
        engineer_details = project_details[project_details["Personnel"] == selected_engineer]

        st.subheader(f"Update Allocations for {selected_engineer}")
        updated_rows = []
        for _, row in engineer_details.iterrows():
            updated_budgeted = st.number_input(
                f"Budgeted Hours ({row['Week']})",
                min_value=0,
                value=int(row["Budgeted Hrs"]) if not pd.isna(row["Budgeted Hrs"]) else 0,
                step=1,
                key=f"update_budgeted_{row['Personnel']}_{row['Week']}"
            )
            updated_spent = st.number_input(
                f"Spent Hours ({row['Week']})",
                min_value=0,
                value=int(row["Spent Hrs"]) if not pd.isna(row["Spent Hrs"]) else 0,
                step=1,
                key=f"update_spent_{row['Personnel']}_{row['Week']}"
            )
            remaining_hours = updated_budgeted - updated_spent
            budgeted_cost = updated_budgeted * row["Cost/Hour"]
            remaining_cost = remaining_hours * row["Cost/Hour"]

            updated_rows.append({
                "Project ID": row["Project ID"],
                "Project Name": row["Project Name"],
                "Personnel": row["Personnel"],
                "Week": row["Week"],
                "Year": row["Year"],
                "Month": row["Month"],
                "Budgeted Hrs": updated_budgeted,
                "Spent Hrs": updated_spent,
                "Remaining Hrs": remaining_hours,
                "Cost/Hour": row["Cost/Hour"],
                "Budgeted Cost": budgeted_cost,
                "Remaining Cost": remaining_cost,
                "Section": row["Section"],
                "Category": row["Category"]
            })

        # Save Updates
        if st.button("Save Updates"):
            updated_rows_df = pd.DataFrame(updated_rows)
            remaining_data = projects_data[~projects_data.index.isin(project_details.index)]
            final_data = pd.concat([remaining_data, updated_rows_df], ignore_index=True)
            upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
            st.success(f"Updates saved successfully!")

if __name__ == "__main__":
    main()

