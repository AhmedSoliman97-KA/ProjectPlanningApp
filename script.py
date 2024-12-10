import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name

# Path to the Excel file
EXCEL_FILE = "projects_data_weekly.xlsx"

# Function to load data
def load_data(file_path):
    try:
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

# Main application
def main():
    st.title("Water & Environment Project Planning")

    # Load project data
    data = load_data(EXCEL_FILE)
    if data is None:
        st.stop()

    # Step 1: User Selection (New or Update Project)
    action = st.radio("What would you like to do?", ["Create New Project", "Update Existing Project"])

    if action == "Create New Project":
        st.subheader("Create a New Project")

        # Input project details
        project_name = st.text_input("Project Name")
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)
        selected_month_index = list(month_name).index(selected_month)

        # Generate weeks
        weeks = generate_weeks_for_month(selected_year, selected_month_index)

        # Input weekly hours
        st.subheader("Assign Weekly Budgeted Hours")
        weekly_data = []
        for week_label, _ in weeks:
            budgeted_hours = st.number_input(f"Budgeted Hours for {week_label}", min_value=0, step=1)
            weekly_data.append({"Week": week_label, "Budgeted Hours": budgeted_hours})

        # Submit the project
        if st.button("Submit Project"):
            new_data = pd.DataFrame(weekly_data)
            new_data["Project Name"] = project_name
            new_data["Year"] = selected_year
            new_data["Month"] = selected_month

            if data is not None:
                data = pd.concat([data, new_data], ignore_index=True)
            else:
                data = new_data

            save_to_excel(data, EXCEL_FILE)
            st.success("Project created successfully!")

    elif action == "Update Existing Project":
        st.subheader("Update an Existing Project")

        # Select project to update
        if "Project Name" not in data.columns:
            st.error("No project data found in the file.")
            st.stop()

        project_names = data["Project Name"].unique().tolist()
        selected_project = st.selectbox("Select Project", project_names)

        if selected_project:
            project_data = data[data["Project Name"] == selected_project]
            updated_data = []

            for index, row in project_data.iterrows():
                col1, col2 = st.columns(2)
                with col1:
                    budgeted_hours = st.number_input(
                        f"Budgeted Hours ({row['Week']})", 
                        min_value=0, 
                        value=row["Budgeted Hours"],
                        step=1,
                        key=f"budget_{index}"
                    )
                with col2:
                    spent_hours = st.number_input(
                        f"Spent Hours ({row['Week']})",
                        min_value=0,
                        value=row.get("Spent Hours", 0),
                        step=1,
                        key=f"spent_{index}"
                    )
                updated_data.append({
                    "Project Name": row["Project Name"],
                    "Week": row["Week"],
                    "Year": row["Year"],
                    "Month": row["Month"],
                    "Budgeted Hours": budgeted_hours,
                    "Spent Hours": spent_hours
                })

            # Submit updates
            if st.button("Save Updates"):
                updated_df = pd.DataFrame(updated_data)
                remaining_data = data[data["Project Name"] != selected_project]
                final_data = pd.concat([remaining_data, updated_df], ignore_index=True)
                save_to_excel(final_data, EXCEL_FILE)
                st.success("Project updated successfully!")

if __name__ == "__main__":
    main()
