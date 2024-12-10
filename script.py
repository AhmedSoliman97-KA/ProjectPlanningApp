import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name

# Path to Excel files
EXCEL_FILE = "projects_data_weekly.xlsx"  # For saving and loading project data
HR_FILE = "Human Resources.xlsx"  # For fetching engineers based on section

# Function to load data
def load_data(file_path, sheet_name=None):
    try:
        if sheet_name:
            return pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            return pd.read_excel(file_path)
    except FileNotFoundError:
        st.error(f"The file '{file_path}' was not found. Please ensure it's uploaded to your repository.")
        return None
    except ImportError:
        st.error("Missing optional dependency 'openpyxl'. Use pip to install openpyxl.")
        return None
    except Exception as e:
        st.error(f"Error loading file '{file_path}': {e}")
        return None

# Save data back to the Excel file
def save_to_excel(data, file_path):
    try:
        data.to_excel(file_path, index=False)
    except Exception as e:
        st.error(f"Failed to save data: {e}")

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

# Calculate additional fields for summary
def calculate_additional_fields(data, resources_db):
    enhanced_data = []
    for row in data:
        engineer_name = row["Personnel"]
        resource_details = resources_db[resources_db["Name"] == engineer_name]
        if not resource_details.empty:
            category = resource_details["Category"].values[0]
            section = resource_details["Section"].values[0]
            cost_per_hour = resource_details["Cost/Hour"].values[0]
            remaining_hours = max(0, row["Budgeted Hrs"] - row.get("Spent Hrs", 0))
            budgeted_cost = row["Budgeted Hrs"] * cost_per_hour
            remaining_cost = remaining_hours * cost_per_hour
            row.update({
                "Category": category,
                "Section": section,
                "Remaining Hrs": remaining_hours,
                "Cost/Hour": cost_per_hour,
                "Budgeted Cost": budgeted_cost,
                "Remaining Cost": remaining_cost,
            })
        else:
            # If no matching resource, add empty or default values
            row.update({
                "Category": "N/A",
                "Section": "N/A",
                "Remaining Hrs": 0,
                "Cost/Hour": 0,
                "Budgeted Cost": 0,
                "Remaining Cost": 0,
            })
        enhanced_data.append(row)
    return enhanced_data

# Main application
def main():
    st.title("Water & Environment Project Planning")

    # Load sections and engineers
    hr_sections = load_data(HR_FILE)
    if hr_sections is None:
        st.stop()

    # Step 1: Choose Action
    st.subheader("Step 1: Choose Action")
    action = st.radio("What would you like to do?", ["Create New Project", "Update Existing Project"])

    # Step 2: Select Section
    st.subheader("Step 2: Select Section")
    sections = load_data(HR_FILE)
    selected_section = st.selectbox("Select a Section", hr_sections)

    # Filter engineers by section
    engineers_data = load_data(HR_FILE, sheet_name=selected_section)
    if engineers_data is None or "Name" not in engineers_data.columns:
        st.error(f"Section {selected_section} doesn't have a valid 'Name' column")
        st.stop()

    engineer_names = engineers_data["Name"].dropna().tolist()

    # Load existing projects
    try:
        project_data = load_data(EXCEL_FILE)
    except FileNotFoundError:
        project_data = pd.DataFrame(columns=["Project Name", "Personnel", "Week", "Budgeted Hrs", "Spent Hrs"])

    if action == "Create New Project":
        st.subheader("Step 3: Create New Project")
        project_name = st.text_input("Project Name")
        project_id = st.text_input("Project ID")
        year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
        month = st.selectbox("Month", list(month_name)[1:])
        month_index = list(month_name).index(month)

        weeks = generate_weeks_for_month(year, month_index)

        st.subheader("Step 4: Assign Weekly Budgeted Hours")
        if "new_project_allocations" not in st.session_state:
            st.session_state.new_project_allocations = []

        for engineer in engineer_names:
            st.markdown(f"**{engineer}**")
            for week_label, _ in weeks:
                hours = st.number_input(f"{week_label} Hours", min_value=0, step=1, key=f"{engineer}_{week_label}")
                if hours > 0:
                    st.session_state.new_project_allocations.append(
                        {"Project Name": project_name, "Personnel": engineer, "Week": week_label, "Budgeted Hrs": hours}
                    )

        if st.button("Submit Project"):
            new_data = pd.DataFrame(st.session_state.new_project_allocations)
            project_data = pd.concat([project_data, new_data], ignore_index=True)
            save_to_excel(project_data, EXCEL_FILE)
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

            if st.button("Save Updates"):
                save_to_excel(project_data, EXCEL_FILE)
                st.success(f"Project '{selected_project}' updated successfully!")

    # Show Summary
    st.subheader("Summary of Projects")
    if not project_data.empty:
        st.dataframe(project_data)
        st.write(f"**Total Budgeted Hours:** {project_data['Budgeted Hrs'].sum()}")
        st.write(f"**Total Spent Hours:** {project_data['Spent Hrs'].sum()}")

if __name__ == "__main__":
    main()
