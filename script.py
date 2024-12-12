import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox

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

        # Table Input for Engineers and Weeks
        if engineers:
            weeks = generate_weeks(selected_year, list(month_name).index(selected_month))
            table_data = pd.DataFrame(
                0,
                index=engineers,
                columns=[f"{week_label}" for week_label, _ in weeks]
            )

            st.subheader("Allocate Weekly Hours")
            edited_table = st.experimental_data_editor(
                table_data,
                use_container_width=True,
                num_rows="dynamic"
            )

            # Calculate Summary Metrics
            st.subheader("Summary Metrics")
            total_budgeted_hours = edited_table.sum().sum()
            st.metric("Total Budgeted Hours", f"{total_budgeted_hours} hrs")

            # Submit Button
            if st.button("Submit Project"):
                allocations = []
                for engineer in edited_table.index:
                    for week, hours in edited_table.loc[engineer].items():
                        if hours > 0:
                            week_label = week.split(" (")[0]
                            budgeted_hours = hours
                            spent_hours = 0
                            remaining_hours = budgeted_hours - spent_hours
                            cost_per_hour = pd.to_numeric(
                                engineers_data[engineers_data["Name"] == engineer]["Cost/Hour"].iloc[0],
                                errors='coerce'
                            )
                            cost_per_hour = cost_per_hour if not pd.isna(cost_per_hour) else 0
                            budgeted_cost = budgeted_hours * cost_per_hour
                            remaining_cost = remaining_hours * cost_per_hour

                            allocations.append({
                                "Project ID": project_id,
                                "Project Name": project_name,
                                "Personnel": engineer,
                                "Week": week_label,
                                "Year": selected_year,
                                "Month": selected_month,
                                "Budgeted Hrs": budgeted_hours,
                                "Spent Hrs": spent_hours,
                                "Remaining Hrs": remaining_hours,
                                "Cost/Hour": cost_per_hour,
                                "Budgeted Cost": budgeted_cost,
                                "Remaining Cost": remaining_cost,
                                "Section": selected_section,
                                "Category": ""
                            })

                # Save Data to Dropbox
                if allocations:
                    new_data = pd.DataFrame(allocations)
                    new_data["Composite Key"] = (
                        new_data["Project ID"] + "_" +
                        new_data["Project Name"] + "_" +
                        new_data["Personnel"] + "_" +
                        new_data["Week"]
                    )
                    projects_data["Composite Key"] = (
                        projects_data["Project ID"] + "_" +
                        projects_data["Project Name"] + "_" +
                        projects_data["Personnel"] + "_" +
                        projects_data["Week"]
                    )
                    updated_data = projects_data[~projects_data["Composite Key"].isin(new_data["Composite Key"])]
                    updated_data = pd.concat([updated_data, new_data], ignore_index=True)
                    updated_data.drop(columns=["Composite Key"], inplace=True)
                    upload_to_dropbox(updated_data, PROJECTS_FILE_PATH)
                    st.success("Project submitted successfully!")

    # Update Existing Project
    if action == "Update Existing Project":
        st.subheader("Update an Existing Project")

        # Section Selection
        selected_section = st.selectbox("Choose a Section", hr_sections)
        filtered_projects = projects_data[projects_data["Section"] == selected_section]

        if filtered_projects.empty:
            st.warning(f"No projects found for the section: {selected_section}.")
            st.stop()

        selected_project = st.selectbox("Choose a Project", filtered_projects["Project Name"].unique())
        project_details = filtered_projects[filtered_projects["Project Name"] == selected_project]

        # Display Allocations
        st.subheader("Current Allocations")
        current_allocations = project_details.pivot(
            index="Personnel",
            columns="Week",
            values="Budgeted Hrs"
        ).fillna(0)
        updated_allocations = st.experimental_data_editor(
            current_allocations,
            use_container_width=True,
            num_rows="dynamic"
        )

        st.subheader("Updated Metrics")
        total_updated_hours = updated_allocations.sum().sum()
        st.metric("Total Updated Hours", f"{total_updated_hours} hrs")

        if st.button("Save Updates"):
            for engineer, row in updated_allocations.iterrows():
                for week, updated_hours in row.items():
                    week_label = week
                    project_details.loc[
                        (project_details["Personnel"] == engineer) &
                        (project_details["Week"] == week_label), "Budgeted Hrs"
                    ] = updated_hours

            # Save back to Dropbox
            upload_to_dropbox(project_details, PROJECTS_FILE_PATH)
            st.success("Project updates saved successfully!")

if __name__ == "__main__":
    main()

