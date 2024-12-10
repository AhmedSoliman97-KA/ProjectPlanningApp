import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name

# Load data from the Excel file
def load_data(file_path, sheet_name=None):
    try:
        if sheet_name:
            return pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            return pd.ExcelFile(file_path).sheet_names
    except ImportError:
        st.error("Missing optional dependency 'openpyxl'. Use pip to install openpyxl.")
        return None
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None

# Save project details to Excel with the desired format
def save_to_excel(output_data, output_path):
    output_df = pd.DataFrame(output_data)
    output_df.to_excel(output_path, index=False)

# Generate weeks for a specific month and year
def generate_weeks_for_month(year, month):
    start_date = datetime(year, month, 1)
    end_date = (start_date + timedelta(days=31)).replace(day=1)  # Move to the next month
    weeks = []
    while start_date < end_date:
        week_label = f"Week {start_date.isocalendar()[1]} - {year}"
        weeks.append((week_label, start_date))
        start_date += timedelta(days=7)
    return weeks

# Main application
def main():
    # App header
    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    # Initialize paths
    resources_path = 'Work Load Sheet - Draft Rev 02 - Water Treatment.xlsx'
    hr_file_path = 'Human Resources.xlsx'
    projects_data_path = "projects_data_weekly.xlsx"

    # Load available sections from the HR file (tab names)
    hr_sections = load_data(hr_file_path)
    if hr_sections is None:
        st.stop()

    # Step 1: User Selection (New or Update Project)
    st.subheader("Step 1: Choose Action")
    action = st.radio("What would you like to do?", ["Create New Project", "Update Existing Project"])

    # Section Selection
    st.subheader("Step 2: Select Section")
    selected_section = st.selectbox("Choose a section", hr_sections)

    # Load data for the selected section
    engineers_data = load_data(hr_file_path, sheet_name=selected_section)
    if engineers_data is None or "Name" not in engineers_data.columns:
        st.error(f"The selected section '{selected_section}' does not have a valid 'Name' column.")
        st.stop()

    engineer_names = engineers_data["Name"].dropna().tolist()

    if action == "Create New Project":
        # Step 3: Enter Project Details
        st.subheader("Step 3: Project Details")
        with st.container():
            col1, col2 = st.columns(2)
            with col1:
                project_id = st.text_input("Project ID", help="Enter the unique ID for the project.")
            with col2:
                project_name = st.text_input("Project Name", help="Enter the unique name of the project.")

        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)
        selected_month_index = list(month_name).index(selected_month)

        # Generate weeks
        weeks = generate_weeks_for_month(selected_year, selected_month_index)

        # Step 4: Assign Engineers and Weekly Hours
        st.subheader("Step 4: Assign Engineers and Weekly Budgeted Hours")
        assigned_engineers = st.multiselect("Select Engineers", engineer_names)

        # Initialize dictionary in session state
        if "engineer_allocation" not in st.session_state:
            st.session_state.engineer_allocation = {}

        for engineer in assigned_engineers:
            st.markdown(f"**Engineer: {engineer}**")
            for week_label, _ in weeks:
                col1, col2 = st.columns(2)
                with col1:
                    budgeted_hours = st.number_input(
                        f"Budgeted Hours ({week_label})",
                        min_value=0,
                        step=1,
                        key=f"budget_{engineer}_{week_label}"
                    )
                with col2:
                    if st.button("Add Hours", key=f"add_hours_{engineer}_{week_label}"):
                        # Create a unique key for the allocation
                        unique_key = f"{engineer}_{week_label}"
                        # Add or update the allocation
                        st.session_state.engineer_allocation[unique_key] = {
                            "Project ID": project_id,
                            "Project Name": project_name,
                            "Personnel": engineer,
                            "Week": week_label,
                            "Year": selected_year,
                            "Month": selected_month,
                            "Budgeted Hrs": budgeted_hours,
                            "Spent Hrs": 0,  # No spent hours for new projects
                        }
                        st.success(f"Hours added for {engineer} ({week_label}).")

        # Summary Section
        st.subheader("Summary of Allocations")
        if st.session_state.engineer_allocation:
            summary_data = list(st.session_state.engineer_allocation.values())
            summary_df = pd.DataFrame(summary_data)

            st.dataframe(summary_df)
            st.write(f"**Total Budgeted Hours:** {summary_df['Budgeted Hrs'].sum()}")

        # Submit Button
        if st.button("Submit Project"):
            new_df = pd.DataFrame(st.session_state.engineer_allocation.values())
            try:
                existing_projects = pd.read_excel(projects_data_path)
                final_df = pd.concat([existing_projects, new_df], ignore_index=True)
            except FileNotFoundError:
                final_df = new_df
            final_df.to_excel(projects_data_path, index=False)
            st.success(f"Project '{project_name}' submitted successfully!")
            st.session_state.engineer_allocation = {}

    elif action == "Update Existing Project":
        # Load existing projects
        try:
            projects_data = pd.read_excel(projects_data_path)
        except FileNotFoundError:
            projects_data = pd.DataFrame(columns=["Project ID", "Project Name", "Personnel", "Budgeted Hrs", "Spent Hrs", "Week", "Year", "Month"])

        if projects_data.empty:
            st.warning("No existing projects found. Please create a new project first.")
            st.stop()

        # Step 3: Select Project to Update
        st.subheader("Step 3: Select Project to Update")
        project_names = projects_data["Project Name"].unique().tolist()
        selected_project = st.selectbox("Choose a project", project_names)

        if selected_project:
            # Filter allocations for the selected project
            project_data = projects_data[projects_data["Project Name"] == selected_project]
            st.subheader(f"Update Allocations for Project: {selected_project}")

            updated_data = []

            for index, row in project_data.iterrows():
                engineer = row["Personnel"]
                week = row["Week"]

                if engineer not in engineer_names:
                    st.warning(f"Engineer '{engineer}' is not part of the selected section '{selected_section}'. Skipping...")
                    continue

                col1, col2 = st.columns(2)
                with col1:
                    updated_budgeted_hours = st.number_input(
                        f"Budgeted Hours ({week}) for {engineer}",
                        min_value=0,
                        value=row["Budgeted Hrs"],
                        step=1,
                        key=f"update_budget_{index}"
                    )
                with col2:
                    updated_spent_hours = st.number_input(
                        f"Spent Hours ({week}) for {engineer}",
                        min_value=0,
                        value=row["Spent Hrs"],
                        step=1,
                        key=f"update_spent_{index}"
                    )

                updated_data.append({
                    "Project ID": row["Project ID"],
                    "Project Name": row["Project Name"],
                    "Personnel": engineer,
                    "Week": week,
                    "Year": row["Year"],
                    "Month": row["Month"],
                    "Budgeted Hrs": updated_budgeted_hours,
                    "Spent Hrs": updated_spent_hours,
                })

            # Submit Updates
            if st.button("Save Updates"):
                updated_df = pd.DataFrame(updated_data)
                remaining_data = projects_data[projects_data["Project Name"] != selected_project]
                final_data = pd.concat([remaining_data, updated_df], ignore_index=True)
                final_data.to_excel(projects_data_path, index=False)
                st.success(f"Updates to project '{selected_project}' have been saved successfully!")

if __name__ == "__main__":
    main()
