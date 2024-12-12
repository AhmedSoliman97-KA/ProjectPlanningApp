import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
from st_aggrid import AgGrid, GridOptionsBuilder

# Dropbox Settings
ACCESS_TOKEN = "sl.u.AFYUwp18eTgoFNK7TUsLEW18-UpWE80v1N9bqMD9uWJrv8hHZHQXAFL8oRen_WCym4SuhXCAkfaxmPFyPiuetz5pmbGot-4Q_niLdbuOwcWTUiolWsqLMwsO74vvUNVy0bI3yrALytVBriQ9wpxB5kNrj9L69qPZn1iTGetTDxJD3TTmW3naoDFg4RiiBRXKVyQw5-T6xSix8TfXv36S1opHOZLiCD3O_pIrzNSlrWI6oiARKTLUOM0GlQWUcsfGjltD8vhJu2KZPQ9bbuXj4r8UFqiLr5-Rv3MDUjLX8DFkcBP7Uy8LwXKqDcF2z0rML9pEe-yCdRQ6WvpApenPoszFOEnM9t8mJM4FzMRjGZm6J27TAzE2sM55-Hxu3NRBayy1cG3ouAw9vt_ucKF2uQ3nJIKHdrD8c1Nwl4xTpVYbhxmkVp1lFdOAh5qjXzKxOfxb0O-quOgZQMoYxJ_FDQLXEZQKI6HMmNHC4pjPNGvSvFTk6Z1n7Ltedyouaw9MkC3H4apZG2mhH2P-M6StJ0Uwfrmpn53ExceC7wGjkBcRWeuTYfYcbO4NZPfjfZf8Of4gey3rldtgW-FXRB8yj_MkFNA5yngF0n7eAZMFQboVkPZlWLrMV_XTP9EYQmSNNENsiL1hb311gRiI1lA6c2BPYn-RqJfe4Gt0MBfTACc-QonhzafTJCsmSdL5KWw0kk51K2BIuZSHMSfKPWwxae08N8aGQCHb1_CSLlId0tWaogkjoTQjd4aGS3OuM7AI6KcMCacDzdaBUBzJRiG2LjYi5zNkm3e10e63SKGmDFOKVaQ_pHMXJov3pSDfqs0L4JwO1OonFxBf8ZbyxtQJZwRD6XrTmGw0b9yMrhd8qKhlfz2aOGoxqYPx-0GMLod2yam279Fo5-qmzsWjNPxUqH-n8_63UvLI1qkfVo1VPRdYG_HXTF0C-p71oiMS9-iV8wWd1w0Ya1c2P7wGm2GuWpLogbqW0n4W_hgerMU5yyhSaNlQldYxsSWQrprIdur70DCoc1XF0BI-lorjjaYLqA2W0mzFZKp1Wi3UCXJmqvVFU7t3klA1d28S5CUbu57WzUiWg1cYxPtzumTsmxK3nkkhnUd0E1k-0sM0Z-QVprwJwGadwcc6bBieEeNylJN7aASIwSZSMXRjrwOXgMOS7YQ1s5ro2ZskXXyLritvjBJf6wuAWvdhgIy4i5QZimN56HzaxwOlFNk_fbe24NJLENUaUxP4rkGkEwu8rQ-vcthmGANL36HZ_NoEL46uSiv1kPTZuCN9_XRPTh2HKZ4R7_LDxkc__zL2G7CwDaMTeHSxAuX4o0j2hxIMEwmZlLqzyKPRiaolAuzBdXDsgpO7l9pMT1U_DFixfbFbllhswH2senCfI4T9A2FSj-6NF1wAvQXPxwklxIrER_1qZlNtgadScMJy3sZQheLS9fosdXNMW1BpVyeR_2RTvMLtcY17q9c"
PROJECTS_FILE_PATH = "/Project_Data/projects_data_weekly.xlsx"
HR_FILE_PATH = "/Project_Data/Human Resources.xlsx"

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
        raise

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

# Generate Weeks for a Given Month
def generate_weeks(year, month):
    start_date = datetime(year, month, 1)
    end_date = (start_date + timedelta(days=31)).replace(day=1)
    weeks = []
    while start_date < end_date:
        week_label = f"Week {start_date.isocalendar()[1]} ({start_date.strftime('%b')})"
        weeks.append((week_label, start_date))
        start_date += timedelta(days=7)
    return weeks

# Main Application
def main():
    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    # Ensure the projects file exists in Dropbox
    ensure_dropbox_projects_file_exists(PROJECTS_FILE_PATH)

    # Load Human Resources File from Dropbox
    hr_excel = download_from_dropbox(HR_FILE_PATH)
    if hr_excel is None:
        st.error(f"Human Resources file not found in Dropbox at {HR_FILE_PATH}.")
        st.stop()

    # Load Sections from HR File
    hr_sections = hr_excel.sheet_names

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

    # Create New Project
    if action == "Create New Project":
        st.subheader("Create a New Project")

        # Project details input
        project_id = st.text_input("Project ID", help="Enter a unique ID for the project.")
        project_name = st.text_input("Project Name", help="Enter the name of the project.")
        approved_budget = st.number_input("Approved Total Budget (in $)", min_value=0, step=1)

        # Year and month selection
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)

        # Engineer selection
        st.subheader("Select Engineers for Allocation")
        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers_data = hr_excel.parse(sheet_name=selected_section)
        engineers = engineers_data["Name"].dropna().tolist()

        # Ag-Grid Configuration for Tabular Input
        st.subheader("Allocate Weekly Hours")
        allocations_template = pd.DataFrame(columns=[
            "Engineer", "Week", "Budgeted Hours", "Spent Hours", "Cost/Hour", "Budgeted Cost", "Remaining Cost"
        ])

        # Add engineer details dynamically
        for engineer in engineers:
            engineer_details = engineers_data[engineers_data["Name"] == engineer].iloc[0]
            cost_per_hour = pd.to_numeric(engineer_details.get("Cost/Hour", 0), errors='coerce')
            cost_per_hour = cost_per_hour if not pd.isna(cost_per_hour) else 0
            allocations_template = pd.concat([allocations_template, pd.DataFrame([{
                "Engineer": engineer,
                "Week": "",
                "Budgeted Hours": 0,
                "Spent Hours": 0,
                "Cost/Hour": cost_per_hour,
                "Budgeted Cost": 0,
                "Remaining Cost": 0
            }])], ignore_index=True)

        gb = GridOptionsBuilder.from_dataframe(allocations_template)
        gb.configure_default_column(editable=True)
        gb.configure_column("Engineer", editable=False)
        grid_options = gb.build()

        response = AgGrid(allocations_template, gridOptions=grid_options, update_mode='MANUAL')
        updated_data = pd.DataFrame(response['data'])

        # Recalculate Budgeted and Remaining Costs dynamically
        updated_data["Budgeted Cost"] = updated_data["Budgeted Hours"] * updated_data["Cost/Hour"]
        updated_data["Remaining Cost"] = updated_data["Budgeted Cost"] - (updated_data["Spent Hours"] * updated_data["Cost/Hour"])

        st.subheader("Summary of Allocations")
        st.dataframe(updated_data)

        # Submit Button
        if st.button("Submit Project"):
            new_data = updated_data
            new_data["Composite Key"] = (
                project_id + "_" +
                project_name + "_" +
                new_data["Engineer"] + "_" +
                new_data["Week"]
            )
            projects_data["Composite Key"] = (
                projects_data["Project ID"] + "_" +
                projects_data["Project Name"] + "_" +
                projects_data["Personnel"] + "_" +
                projects_data["Week"]
            )
            updated_projects = projects_data[~projects_data["Composite Key"].isin(new_data["Composite Key"])]
            final_data = pd.concat([updated_projects, new_data], ignore_index=True)
            final_data.drop(columns=["Composite Key"], inplace=True)
            upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
            st.success("Project submitted successfully!")

    # Update Existing Project
    if action == "Update Existing Project":
        st.subheader("Update an Existing Project")

        # Step 1: Section Selection
        st.subheader("Select Section")
        selected_section = st.selectbox("Choose a Section", hr_sections)

        # Filter Projects by Selected Section
        filtered_projects = projects_data[projects_data["Section"] == selected_section]

        if filtered_projects.empty:
            st.warning(f"No projects found for the section: {selected_section}.")
            st.stop()

        # Step 2: Select Project
        st.subheader("Select a Project")
        selected_project = st.selectbox("Choose a Project", filtered_projects["Project Name"].unique())
        project_details = filtered_projects[filtered_projects["Project Name"] == selected_project]

        # Step 3: Display Current Allocations for the Selected Project
        st.subheader("Current Allocations for Selected Project")
        st.dataframe(project_details)

        # Ag-Grid for Updating Allocations
        gb = GridOptionsBuilder.from_dataframe(project_details)
        gb.configure_default_column(editable=True)
        grid_options = gb.build()
        response = AgGrid(project_details, gridOptions=grid_options, update_mode='MANUAL')
        updated_project_data = pd.DataFrame(response['data'])

        # Submit Updates
        if st.button("Save Updates"):
            updated_project_data["Composite Key"] = (
                updated_project_data["Project ID"] + "_" +
                updated_project_data["Project Name"] + "_" +
                updated_project_data["Personnel"] + "_" +
                updated_project_data["Week"]
            )
            projects_data["Composite Key"] = (
                projects_data["Project ID"] + "_" +
                projects_data["Project Name"] + "_" +
                projects_data["Personnel"] + "_" +
                projects_data["Week"]
            )
            remaining_data = projects_data[~projects_data["Composite Key"].isin(updated_project_data["Composite Key"])]
            final_data = pd.concat([remaining_data, updated_project_data], ignore_index=True)
            final_data.drop(columns=["Composite Key"], inplace=True)

            # Upload updated data back to Dropbox
            upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
            st.success(f"Updates to '{selected_project}' saved successfully!")

if __name__ == "__main__":
    main()
