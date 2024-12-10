import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
from io import BytesIO
import requests

# URL to the Human Resources Excel file (replace with your link)
HR_FILE_URL = "https://khatibandalami-my.sharepoint.com/:x:/g/personal/ahmedsayed_soliman_khatibalami_com/EXQjPzZs9h5Nly5JKGQQmCEBIZP_IqkYo-ClHsSLplfSdQ?e=qXUI84"  # Replace with your direct link
PROJECTS_FILE = "projects_data_weekly.xlsx"  # Local file for project data

# Function to read Excel data directly from a URL
def load_excel_from_url(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            # Specify the engine for .xlsx files
            return pd.read_excel(BytesIO(response.content), engine="openpyxl")
        else:
            st.error(f"Failed to fetch file: HTTP {response.status_code}")
            return pd.DataFrame()
    except ValueError as e:
        st.error(f"Error reading data from URL: {e}")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Unexpected error: {e}")
        return pd.DataFrame()

# Load or create project data
def load_or_create_excel(file_path):
    try:
        return pd.read_excel(file_path)
    except FileNotFoundError:
        # Create an empty DataFrame if file does not exist
        return pd.DataFrame(columns=["Project ID", "Project Name", "Personnel", "Week", "Budgeted Hrs", "Spent Hrs"])

# Save project data to an Excel file
def save_to_excel(data, file_path):
    try:
        data.to_excel(file_path, index=False)
        st.success("Data saved successfully!")
    except Exception as e:
        st.error(f"Failed to save data: {e}")

# Generate weekly labels for a specific year and month
def generate_weeks_for_month(year, month):
    start_date = datetime(year, month, 1)
    end_date = (start_date + timedelta(days=31)).replace(day=1)
    weeks = []
    while start_date < end_date:
        week_label = f"Week {start_date.isocalendar()[1]} - {year}"
        weeks.append((week_label, start_date))
        start_date += timedelta(days=7)
    return weeks

# Main Streamlit Application
def main():
    st.title("Water & Environment Project Planning")

    # Step 1: Load Human Resources Data from URL
    st.subheader("Step 1: Load Human Resources Data")
    hr_data = load_excel_from_url(HR_FILE_URL)
    if hr_data.empty:
        st.error("Could not load Human Resources data. Please check the URL and try again.")
        return

    # Display loaded data for verification
    st.write("Loaded Human Resources Data:")
    st.dataframe(hr_data)

    # Extract sections and engineer names from HR data
    if "Section" not in hr_data.columns or "Name" not in hr_data.columns:
        st.error("The Human Resources data does not have 'Section' or 'Name' columns.")
        return
    sections = hr_data["Section"].dropna().unique().tolist()

    # Step 2: Load or Create Project Data
    project_data = load_or_create_excel(PROJECTS_FILE)

    # Step 3: Choose Action
    st.subheader("Step 2: Choose Action")
    action = st.radio("What would you like to do?", ["Create New Project", "Update Existing Project"])

    # Step 4: Select Section and Filter Engineers
    st.subheader("Step 3: Select Section")
    selected_section = st.selectbox("Select a Section", sections)
    filtered_engineers = hr_data[hr_data["Section"] == selected_section]["Name"].dropna().tolist()

    if action == "Create New Project":
        st.subheader("Create a New Project")

        # Input project details
        project_id = st.text_input("Project ID")
        project_name = st.text_input("Project Name")
        year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
        month = st.selectbox("Month", list(month_name)[1:])
        month_index = list(month_name).index(month)

        weeks = generate_weeks_for_month(year, month_index)

        # Assign weekly hours for each engineer
        st.subheader("Assign Weekly Budgeted Hours")
        if "new_project_allocations" not in st.session_state:
            st.session_state.new_project_allocations = []

        for engineer in filtered_engineers:
            st.markdown(f"**{engineer}**")
            for week_label, _ in weeks:
                hours = st.number_input(f"{week_label} Hours for {engineer}", min_value=0, step=1,
                                        key=f"{engineer}_{week_label}")
                if hours > 0:
                    st.session_state.new_project_allocations.append(
                        {"Project ID": project_id, "Project Name": project_name, "Personnel": engineer,
                         "Week": week_label, "Year": year, "Month": month, "Budgeted Hrs": hours}
                    )

        # Save new project data
        if st.button("Submit Project"):
            new_data = pd.DataFrame(st.session_state.new_project_allocations)
            project_data = pd.concat([project_data, new_data], ignore_index=True)
            save_to_excel(project_data, PROJECTS_FILE)
            st.success(f"Project '{project_name}' created successfully!")
            st.session_state.new_project_allocations = []

    elif action == "Update Existing Project":
        st.subheader("Update an Existing Project")
        project_names = project_data["Project Name"].unique().tolist()
        selected_project = st.selectbox("Select Project", project_names)

        if selected_project:
            filtered_data = project_data[project_data["Project Name"] == selected_project]
            for index, row in filtered_data.iterrows():
                col1, col2 = st.columns(2)
                with col1:
                    budgeted_hours = st.number_input(
                        f"Budgeted Hours for {row['Personnel']} ({row['Week']})",
                        min_value=0, value=row["Budgeted Hrs"], step=1, key=f"budget_{index}"
                    )
                with col2:
                    spent_hours = st.number_input(
                        f"Spent Hours for {row['Personnel']} ({row['Week']})",
                        min_value=0, value=row["Spent Hrs"], step=1, key=f"spent_{index}"
                    )
                project_data.loc[index, "Budgeted Hrs"] = budgeted_hours
                project_data.loc[index, "Spent Hrs"] = spent_hours

            if st.button("Save Updates"):
                save_to_excel(project_data, PROJECTS_FILE)
                st.success(f"Project '{selected_project}' updated successfully!")

    # Display Summary
    st.subheader("Project Summary")
    st.dataframe(project_data)
    st.write(f"**Total Budgeted Hours:** {project_data['Budgeted Hrs'].sum()}")
    st.write(f"**Total Spent Hours:** {project_data['Spent Hrs'].sum()}")

if __name__ == "__main__":
    main()
