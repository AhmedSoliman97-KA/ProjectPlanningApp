import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox

# Dropbox Settings
ACCESS_TOKEN = "YOUR_DROPBOX_ACCESS_TOKEN"
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
        selected_engineers = st.multiselect("Choose Engineers", options=engineers, help="Select engineers to allocate hours.")

        allocations = []
        total_allocated_budget = 0

        if selected_engineers:
            st.subheader("Allocate Weekly Hours")

            # Generate Weeks
            weeks = generate_weeks(selected_year, list(month_name).index(selected_month))

            for engineer in selected_engineers:
                # Fetch engineer details from Human Resources file
                engineer_details = engineers_data[engineers_data["Name"] == engineer].iloc[0]
                section = engineer_details.get("Section", "Unknown")
                category = engineer_details.get("Category", "N/A")
                cost_per_hour = pd.to_numeric(engineer_details.get("Cost/Hour", 0), errors='coerce')
                cost_per_hour = cost_per_hour if not pd.isna(cost_per_hour) else 0

                st.markdown(f"### Engineer: {engineer}")
                for week_label, _ in weeks:
                    budgeted_hours = st.number_input(
                        f"Budgeted Hours ({week_label}) for {engineer}",
                        min_value=0,
                        step=1,
                        key=f"{engineer}_{week_label}"
                    )
                    if budgeted_hours > 0:
                        spent_hours = 0
                        remaining_hours = budgeted_hours - spent_hours
                        budgeted_cost = budgeted_hours * cost_per_hour
                        remaining_cost = remaining_hours * cost_per_hour

                        total_allocated_budget += budgeted_cost

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
                            "Section": section,
                            "Category": category
                        })

        # Display Allocation Summary and Comparison
        if allocations:
            st.subheader("Summary of Allocations")
            allocation_df = pd.DataFrame(allocations)
            st.dataframe(allocation_df)

            # Display Total Allocated Budget
            st.metric("Total Allocated Budget (in $)", f"${total_allocated_budget:,.2f}")
            st.metric("Approved Total Budget (in $)", f"${approved_budget:,.2f}")
            st.metric("Difference (Remaining/Over-Allocated)", f"${approved_budget - total_allocated_budget:,.2f}")

        # Submit Button
        if st.button("Submit Project"):
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

        # Step 4: Select Engineer for Update
        st.subheader("Select Engineer")
        selected_engineer = st.selectbox("Choose an Engineer", project_details["Personnel"].unique())
        engineer_details = project_details[project_details["Personnel"] == selected_engineer]

        st.subheader(f"Update Allocations for {selected_engineer}")
        updated_rows = []

        # Step 5: Update Allocations for Each Week
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

        # Display Summary of Updated Allocations
        if updated_rows:
            st.subheader("Summary of Updated Allocations")
            updated_df = pd.DataFrame(updated_rows)
            st.dataframe(updated_df)
            st.metric("Total Budgeted Hours", updated_df["Budgeted Hrs"].sum())
            st.metric("Total Spent Hours", updated_df["Spent Hrs"].sum())

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
