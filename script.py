import streamlit as st
import pandas as pd
import requests
import msal
from io import BytesIO
from datetime import datetime, timedelta
from calendar import month_name

# OneDrive API credentials
CLIENT_ID = "3686715d-f3f7-41d9-ae6b-bd722174bc6b"
TENANT_ID = "0fa087f9-be01-4a3e-874d-03fd3b33f1b6"
CLIENT_SECRET = "53de388a-fc0b-4471-99b7-e4222bca80fd"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# Authenticate and get access token
def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Authentication failed: " + str(result))

# Download a file from OneDrive
def download_file_from_onedrive(access_token, file_path):
    url = f"https://engsohagedu-my.sharepoint.com/:x:/g/personal/ahmed2016018_eng_sohag_edu_eg/Ea1YnybszbRGlMDOuKuYfj0BxX-E7PDl0HctF6SB3KNEyw?e=rM0p69"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return pd.read_excel(BytesIO(response.content), engine="openpyxl")
    else:
        raise Exception(f"Failed to download file: {response.status_code} - {response.text}")

# Upload a file to OneDrive
def upload_file_to_onedrive(access_token, file_path, data_frame):
    url = f"https://engsohagedu-my.sharepoint.com/:f:/g/personal/ahmed2016018_eng_sohag_edu_eg/EqWXv4SllaNOtcCL8Mq7n2AB-dy39lJn6kWhkxFl8a8_rQ?e=IcW4Gu"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    with BytesIO() as output:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            data_frame.to_excel(writer, index=False)
        output.seek(0)
        response = requests.put(url, headers=headers, data=output)
    if response.status_code in [200, 201]:
        print(f"File uploaded successfully to {file_path}.")
    else:
        raise Exception(f"Failed to upload file: {response.status_code} - {response.text}")

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

    # Authenticate and get access token
    try:
        access_token = get_access_token()
    except Exception as e:
        st.error(f"Authentication failed: {e}")
        return

    # Load Human Resources data
    try:
        hr_data = download_file_from_onedrive(access_token, "Human Resources.xlsx")
        st.write("Human Resources Data:")
        st.dataframe(hr_data)
    except Exception as e:
        st.error(f"Failed to load Human Resources data: {e}")
        return

    # Load or initialize project data
    try:
        project_data = download_file_from_onedrive(access_token, "projects_data_weekly.xlsx")
    except Exception:
        st.warning("No existing project data found. Initializing a new file.")
        project_data = pd.DataFrame(columns=["Project ID", "Project Name", "Personnel", "Week", "Budgeted Hrs", "Spent Hrs"])

    # Step 1: Choose Action
    st.subheader("Step 1: Choose Action")
    action = st.radio("What would you like to do?", ["Create New Project", "Update Existing Project"])

    # Step 2: Add New Project
    if action == "Create New Project":
        st.subheader("Create a New Project")

        # Input project details
        project_id = st.text_input("Project ID")
        project_name = st.text_input("Project Name")
        year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
        month = st.selectbox("Month", list(month_name)[1:])
        month_index = list(month_name).index(month)

        weeks = generate_weeks_for_month(year, month_index)

        # Assign weekly hours for each personnel
        if "new_project_allocations" not in st.session_state:
            st.session_state.new_project_allocations = []

        personnel_list = hr_data["Name"].dropna().unique().tolist()
        for person in personnel_list:
            st.markdown(f"**{person}**")
            for week_label, _ in weeks:
                hours = st.number_input(f"{week_label} Hours for {person}", min_value=0, step=1,
                                        key=f"{person}_{week_label}")
                if hours > 0:
                    st.session_state.new_project_allocations.append(
                        {"Project ID": project_id, "Project Name": project_name, "Personnel": person,
                         "Week": week_label, "Year": year, "Month": month, "Budgeted Hrs": hours}
                    )

        # Save new project data
        if st.button("Submit Project"):
            new_data = pd.DataFrame(st.session_state.new_project_allocations)
            project_data = pd.concat([project_data, new_data], ignore_index=True)
            try:
                upload_file_to_onedrive(access_token, "projects_data_weekly.xlsx", project_data)
                st.success(f"Project '{project_name}' created successfully!")
                st.session_state.new_project_allocations = []
            except Exception as e:
                st.error(f"Failed to save project data: {e}")

    # Step 3: Update Existing Project
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
                try:
                    upload_file_to_onedrive(access_token, "projects_data_weekly.xlsx", project_data)
                    st.success(f"Project '{selected_project}' updated successfully!")
                except Exception as e:
                    st.error(f"Failed to save updates: {e}")

    # Display Summary
    st.subheader("Project Summary")
    st.dataframe(project_data)
    st.write(f"**Total Budgeted Hours:** {project_data['Budgeted Hrs'].sum()}")
    st.write(f"**Total Spent Hours:** {project_data['Spent Hrs'].sum()}")

if __name__ == "__main__":
    main()
