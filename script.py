import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name

# Paths to Excel files
PROJECTS_FILE = "projects_data_weekly.xlsx"  # To store project allocations
HR_FILE = "Human Resources.xlsx"  # To fetch engineers and their details

# Load data from an Excel file
def load_data(file_path, sheet_name=None):
    try:
        if sheet_name:
            return pd.read_excel(file_path, sheet_name=sheet_name)
        return pd.read_excel(file_path)
    except FileNotFoundError:
        st.error(f"The file '{file_path}' was not found.")
        return None
    except Exception as e:
        st.error(f"Error loading file '{file_path}': {e}")
        return None

# Save data to an Excel file
def save_to_excel(data, file_path):
    try:
        data.to_excel(file_path, index=False)
    except Exception as e:
        st.error(f"Failed to save data: {e}")

# Generate weeks for a specific month and year
def generate_weeks_for_month(year, month):
    start_date = datetime(year, month, 1)
    end_date = (start_date + timedelta(days=31)).replace(day=1)  # Move to next month
    weeks = []
    while start_date < end_date:
        week_label = f"Week {start_date.isocalendar()[1]} - {year}"
        weeks.append((week_label, start_date))
        start_date += timedelta(days=7)
    return weeks

# Main application
def main():
    st.title("Water & Environment Project Planning")

    # Load HR sections
    hr_sections = load_data(HR_FILE)
    if hr_sections is None:
        st.stop()

    # Load existing projects
    try:
        project_data = load_data(PROJECTS_FILE)
    except FileNotFoundError:
        project_data = pd.DataFrame(
            columns=[
                "Project ID", "Project Name", "Personnel", "Week",
                "Year", "Month", "Budgeted Hrs", "Spent Hrs"
            ]
        )

    # Step 1: Choose Action
    st.subheader("Step 1: Choose Action")
    action = st.radio("What would you like to do?", ["Create New Project", "Update Existing Project"])

    # Step 2: Select Section
    st.subheader("Step 2: Select Section")
    sections = load_data(HR_FILE)
    selected_section = st.selectbox("Select a Section", hr_sections)

    # Filter engineers based on the selected section
    engineers_data = load_data(HR_FILE, sheet_name=selected_section)
    if engineers_data is None or "Name" not in engineers_data.columns:
        st.error(f"Section {selected_section} doesn't have a valid 'Name' column")
        st.stop()
    engineer_names = engineers_data["Name"].dropna().tolist()

    if action == "Create New Project":
        st.subheader("Step 3: Create New Project")

        # Collect project details
        project_id = st.text_input("Project ID", help="Enter the unique ID for the project.")
        project_name = st.text_input("Project Name", help="Enter the unique name of the project.")
        year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
        month = st.selectbox("Month", list(month_name)[1:])
        month_index = list(month_name).index(month)

        weeks = generate_weeks_for_month(year, month_index)

        # Assign hours for engineers
        st.subheader("Step 4: Assign Weekly Budgeted Hours")
        if "new_project_allocations" not in st.session_state:
            st.session_state.new_project_allocations = []

        for engineer in engineer_names:
            st.markdown(f"**{engineer}**")
            for week_label, _ in weeks:
                hours = st.number_input(f"{week_label} Hours", min_value=0, step=1, key=f"{engineer}_{week_label}")
                if hours > 0:
                    st.session_state.new_project_allocations.append(
                        {"Project ID": project_id, "Project Name": project_name,
                         "Personnel": engineer, "Week": week_label,
                         "Year": year, "Month": month, "Budgeted Hrs": hours}
                    )

        # Save new project
        if st.button("Submit Project"):
            new_data = pd.DataFrame(st.session_state.new_project_allocations)
            project_data = pd.concat([project_data, new_data], ignore_index=True)
            save_to_excel(project_data, PROJECTS_FILE)
            st.success(f"Project '{project_name}' created successfully!")
            st.session_state.new_project_allocations = []

    elif action == "Update Existing Project":
        st.subheader("Step 3: Update Existing Project")
        project_names = project_data["Project Name"].unique()
        selected_project = st.selectbox("Select Project", project_names)

        if selected_project:
            project_allocations = project_data[project_data["Project Name"] == selected_project]
            for i, row in project_allocations.iterrows():
                st.markdown(f"**{row['Personnel']} ({row['Week']})**")
                col1, col2 = st.columns(2)
                with col1:
                    budgeted_hours = st.number_input(
                        "Budgeted Hours", min_value=0, value=row["Budgeted Hrs"], step=1, key=f"budget_{i}"
                    )
                with col2:
                    spent_hours = st.number_input(
                        "Spent Hours", min_value=0, value=row.get("Spent Hrs", 0), step=1, key=f"spent_{i}"
                    )
                project_data.loc[i, "Budgeted Hrs"] = budgeted_hours
                project_data.loc[i, "Spent Hrs"] = spent_hours

            # Save updates
            if st.button("Save Updates"):
                save_to_excel(project_data, PROJECTS_FILE)
                st.success(f"Project '{selected_project}' updated successfully!")

    # Display project summary
    st.subheader("Summary of Projects")
    if not project_data.empty:
        st.dataframe(project_data)
        st.write(f"**Total Budgeted Hours:** {project_data['Budgeted Hrs'].sum()}")
        st.write(f"**Total Spent Hours:** {project_data['Spent Hrs'].sum()}")

if __name__ == "__main__":
    main()
 
