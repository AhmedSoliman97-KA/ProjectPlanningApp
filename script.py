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
        weeks.append(week_label)
        start_date += timedelta(days=7)
    return weeks

# Create Editable Table
def create_editable_table(engineers, weeks, existing_data=None):
    """Generate a dynamic editable table for weekly allocation."""
    table_data = []
    for engineer in engineers:
        row = {"Engineer": engineer}
        for week in weeks:
            # If existing data exists, prefill; else default to 0
            if existing_data is not None:
                matching_row = existing_data[
                    (existing_data["Personnel"] == engineer) & (existing_data["Week"] == week)
                ]
                row[week] = (
                    matching_row["Budgeted Hrs"].values[0] if not matching_row.empty else 0
                )
            else:
                row[week] = 0
        table_data.append(row)
    return pd.DataFrame(table_data)

def update_table_inputs(df):
    """Create input fields for editing the table."""
    updated_rows = []
    for i, row in df.iterrows():
        updated_row = {"Engineer": row["Engineer"]}
        for col in df.columns[1:]:  # Skip 'Engineer' column
            updated_row[col] = st.number_input(
                f"{col} ({row['Engineer']})",
                min_value=0,
                value=row[col],
                step=1,
                key=f"{row['Engineer']}_{col}",
            )
        updated_rows.append(updated_row)
    return pd.DataFrame(updated_rows)

# Main App
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

    # Sidebar Action Selection
    st.sidebar.subheader("Actions")
    action = st.sidebar.radio("Choose an Action", ["Create New Project", "Update Existing Project"])

    if action == "Create New Project":
        st.subheader("Create a New Project")

        # Input project details
        project_id = st.text_input("Project ID", help="Enter a unique ID for the project.")
        project_name = st.text_input("Project Name", help="Enter the name of the project.")
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)
        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers = hr_excel.parse(sheet_name=selected_section)["Name"].dropna().tolist()

        weeks = generate_weeks(selected_year, list(month_name).index(selected_month))

        st.subheader("Allocate Weekly Hours")
        table_data = create_editable_table(engineers, weeks)
        edited_table = update_table_inputs(table_data)

        # Display Summary
        st.subheader("Summary of Allocations")
        st.dataframe(edited_table)

        # Submit Button
        if st.button("Submit Project"):
            allocations = []
            for _, row in edited_table.iterrows():
                for week in weeks:
                    budgeted_hours = row[week]
                    if budgeted_hours > 0:
                        allocations.append({
                            "Project ID": project_id,
                            "Project Name": project_name,
                            "Personnel": row["Engineer"],
                            "Week": week,
                            "Year": selected_year,
                            "Month": selected_month,
                            "Budgeted Hrs": budgeted_hours,
                        })
            allocation_df = pd.DataFrame(allocations)
            projects_data = pd.concat([projects_data, allocation_df], ignore_index=True)
            upload_to_dropbox(projects_data, PROJECTS_FILE_PATH)
            st.success("Project submitted successfully!")

    elif action == "Update Existing Project":
        st.subheader("Update an Existing Project")
        selected_section = st.selectbox("Choose a Section", hr_sections)
        filtered_projects = projects_data[projects_data["Section"] == selected_section]

        if filtered_projects.empty:
            st.warning(f"No projects found for the section: {selected_section}.")
            return

        selected_project = st.selectbox("Choose a Project", filtered_projects["Project Name"].unique())
        project_details = filtered_projects[filtered_projects["Project Name"] == selected_project]
        engineers = project_details["Personnel"].unique()
        weeks = project_details["Week"].unique()

        st.subheader("Update Weekly Allocations")
        table_data = create_editable_table(engineers, weeks, project_details)
        edited_table = update_table_inputs(table_data)

        # Display Summary
        st.subheader("Updated Allocations")
        st.dataframe(edited_table)

        if st.button("Save Updates"):
            updated_rows = []
            for _, row in edited_table.iterrows():
                for week in weeks:
                    budgeted_hours = row[week]
                    if budgeted_hours > 0:
                        updated_rows.append({
                            "Project ID": selected_project,
                            "Project Name": selected_project,
                            "Personnel": row["Engineer"],
                            "Week": week,
                            "Budgeted Hrs": budgeted_hours,
                        })
            updated_df = pd.DataFrame(updated_rows)
            projects_data = pd.concat([projects_data, updated_df], ignore_index=True)
            upload_to_dropbox(projects_data, PROJECTS_FILE_PATH)
            st.success("Project updates saved!")

            # Display Total Allocated Budget
            st.metric("Total Allocated Budget (in $)", f"${total_allocated_budget:,.2f}")
            st.metric("Approved Total Budget (in $)", f"${approved_budget:,.2f}")
            st.metric("Difference (Remaining/Over-Allocated)", f"${approved_budget - total_allocated_budget:,.2f}")
        
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

        # Metrics for Current Allocations
        current_budgeted_hours = project_details["Budgeted Hrs"].sum()
        current_spent_hours = project_details["Spent Hrs"].sum()
        current_budgeted_cost = project_details["Budgeted Cost"].sum()
        current_spent_cost = (project_details["Spent Hrs"] * project_details["Cost/Hour"]).sum()

        st.metric("Current Budgeted Hours", f"{current_budgeted_hours} hrs")
        st.metric("Current Spent Hours", f"{current_spent_hours} hrs")
        st.metric("Current Budgeted Cost", f"${current_budgeted_cost:,.2f}")
        st.metric("Current Spent Cost", f"${current_spent_cost:,.2f}")

        # Step 4: Select Engineer for Update or Add New Engineer
        st.subheader("Select Engineer")
        engineer_options = list(project_details["Personnel"].unique()) + ["Add New Engineer"]
        selected_engineer = st.selectbox("Choose an Engineer", engineer_options)

        # Add New Engineer Logic
        if selected_engineer == "Add New Engineer":
            st.subheader("Add New Engineer")
            new_engineers = st.multiselect("Choose New Engineers", options=hr_excel.parse(sheet_name=selected_section)["Name"].dropna().tolist())

            for new_engineer in new_engineers:
                if new_engineer not in project_details["Personnel"].values:
                    project_details = pd.concat([project_details, pd.DataFrame([{
                        "Project ID": project_details.iloc[0]["Project ID"],
                        "Project Name": project_details.iloc[0]["Project Name"],
                        "Personnel": new_engineer,
                        "Week": None,
                        "Year": None,
                        "Month": None,
                        "Budgeted Hrs": 0,
                        "Spent Hrs": 0,
                        "Remaining Hrs": 0,
                        "Cost/Hour": 0,
                        "Budgeted Cost": 0,
                        "Remaining Cost": 0,
                        "Section": selected_section,
                        "Category": None
                    }])])

        # Step 5: Update Allocations for Each Week
        st.subheader(f"Update Allocations for {selected_engineer}")
        updated_rows = []

        # Only display unique weeks for the selected engineer
        engineer_details = project_details[project_details["Personnel"] == selected_engineer]
        weeks = engineer_details["Week"].unique()

        for week in weeks:
            existing_allocation = engineer_details[engineer_details["Week"] == week].iloc[0] if not engineer_details.empty else {}
            updated_budgeted = st.number_input(
                f"Budgeted Hours ({week})",
                min_value=0,
                value=int(existing_allocation.get("Budgeted Hrs", 0)),
                step=1,
                key=f"update_budgeted_{selected_engineer}_{week}"
            )
            updated_spent = st.number_input(
                f"Spent Hours ({week})",
                min_value=0,
                value=int(existing_allocation.get("Spent Hrs", 0)),
                step=1,
                key=f"update_spent_{selected_engineer}_{week}"
            )
            remaining_hours = updated_budgeted - updated_spent
            cost_per_hour = existing_allocation.get("Cost/Hour", 0)
            budgeted_cost = updated_budgeted * cost_per_hour
            spent_cost = updated_spent * cost_per_hour

            updated_rows.append({
                "Project ID": existing_allocation.get("Project ID", ""),
                "Project Name": existing_allocation.get("Project Name", ""),
                "Personnel": selected_engineer,
                "Week": week,
                "Year": existing_allocation.get("Year", None),
                "Month": existing_allocation.get("Month", None),
                "Budgeted Hrs": updated_budgeted,
                "Spent Hrs": updated_spent,
                "Remaining Hrs": remaining_hours,
                "Cost/Hour": cost_per_hour,
                "Budgeted Cost": budgeted_cost,
                "Remaining Cost": budgeted_cost - spent_cost,
                "Section": existing_allocation.get("Section", ""),
                "Category": existing_allocation.get("Category", "")
            })

        # Metrics for Updated Allocations
        st.subheader("Updated Allocations for Selected Project")
        updated_df = pd.DataFrame(updated_rows)
        st.dataframe(updated_df)

        updated_budgeted_hours = updated_df["Budgeted Hrs"].sum()
        updated_spent_hours = updated_df["Spent Hrs"].sum()
        updated_budgeted_cost = updated_df["Budgeted Cost"].sum()
        updated_spent_cost = updated_df["Spent Hrs"].sum() * updated_df["Cost/Hour"].sum()

        st.metric("Updated Budgeted Hours", f"{updated_budgeted_hours} hrs")
        st.metric("Updated Spent Hours", f"{updated_spent_hours} hrs")
        st.metric("Updated Budgeted Cost", f"${updated_budgeted_cost:,.2f}")
        st.metric("Updated Spent Cost", f"${updated_spent_cost:,.2f}")

        # Save Updates
        if st.button("Save Updates"):
            updated_rows_df = pd.DataFrame(updated_rows)
            updated_rows_df["Composite Key"] = (
                updated_rows_df["Project ID"] + "_" +
                updated_rows_df["Project Name"] + "_" +
                updated_rows_df["Personnel"] + "_" +
                updated_rows_df["Week"]
            )
            projects_data["Composite Key"] = (
                projects_data["Project ID"] + "_" +
                projects_data["Project Name"] + "_" +
                projects_data["Personnel"] + "_" +
                projects_data["Week"]
            )
            remaining_data = projects_data[~projects_data["Composite Key"].isin(updated_rows_df["Composite Key"])]
            final_data = pd.concat([remaining_data, updated_rows_df], ignore_index=True)
            final_data.drop(columns=["Composite Key"], inplace=True)

            # Upload updated data back to Dropbox
            upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
            st.success(f"Updates to '{selected_project}' saved successfully!")

if __name__ == "__main__":
    main()
