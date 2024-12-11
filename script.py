import streamlit as st
import pandas as pd
import requests
import msal
from io import BytesIO
from datetime import datetime, timedelta

# OneDrive API credentials
CLIENT_ID = "a7a2fb29-66bd-4e16-a8eb-65d6b7f7b3f1"  # Application (client) ID
TENANT_ID = "c6344517-088e-43b9-970a-f93b99bb4fde"  # Directory (tenant) ID
CLIENT_SECRET = "adn8Q~Jo1shdXUYouk41Z6D9DVokrZdsEfXLTcN6"  # Secret Value

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

# List files in OneDrive root directory
def list_onedrive_files(access_token):
    url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        files = response.json()["value"]
        file_list = [{"name": file["name"], "id": file["id"]} for file in files]
        return file_list
    else:
        raise Exception(f"Failed to list files: {response.status_code} - {response.text}")

# Download a file from OneDrive
def download_file_from_onedrive(access_token, file_id):
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return pd.read_excel(BytesIO(response.content), engine="openpyxl")
    else:
        raise Exception(f"Failed to download file: {response.status_code} - {response.text}")

# Upload a file to OneDrive
def upload_file_to_onedrive(access_token, file_name, data_frame):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_name}:/content"
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
        print(f"File uploaded successfully to {file_name}.")
    else:
        raise Exception(f"Failed to upload file: {response.status_code} - {response.text}")

# Main function for the app
def main():
    st.title("Water & Environment Project Planning")

    # Authenticate and get access token
    try:
        access_token = get_access_token()
        st.success("Successfully authenticated with OneDrive.")
    except Exception as e:
        st.error(f"Authentication failed: {e}")
        return

    # List files in OneDrive and select the required files
    try:
        files = list_onedrive_files(access_token)
        file_names = [file["name"] for file in files]
        file_dict = {file["name"]: file["id"] for file in files}

        st.write("Available files in OneDrive:")
        selected_hr_file = st.selectbox("Select Human Resources file", file_names)
        hr_file_id = file_dict[selected_hr_file]
        hr_data = download_file_from_onedrive(access_token, hr_file_id)
        st.write("Human Resources Data:")
        st.dataframe(hr_data)
    except Exception as e:
        st.error(f"Failed to load Human Resources data: {e}")
        return

    # Load or initialize project data
    try:
        selected_project_file = st.selectbox("Select Projects file", file_names)
        project_file_id = file_dict[selected_project_file]
        project_data = download_file_from_onedrive(access_token, project_file_id)
    except Exception:
        st.warning("No existing project data found. Initializing a new file.")
        project_data = pd.DataFrame(columns=["Project ID", "Project Name", "Personnel", "Week", "Budgeted Hrs", "Spent Hrs"])

    # Select whether to create or update a project
    action = st.radio("Choose an action", ["Create New Project", "Update Existing Project"])

    if action == "Create New Project":
        st.subheader("Create a New Project")
        project_id = st.text_input("Project ID")
        project_name = st.text_input("Project Name")
        year = st.number_input("Year", min_value=2000, max_value=2100, value=datetime.now().year)
        month = st.number_input("Month", min_value=1, max_value=12, value=datetime.now().month)
        weeks = [f"Week {i}" for i in range(1, 5)]  # Example weekly labels

        if "new_project_allocations" not in st.session_state:
            st.session_state.new_project_allocations = []

        personnel_list = hr_data["Name"].dropna().unique().tolist()
        for person in personnel_list:
            for week in weeks:
                hours = st.number_input(f"{week} Hours for {person}", min_value=0, step=1, key=f"{person}_{week}")
                if hours > 0:
                    st.session_state.new_project_allocations.append(
                        {"Project ID": project_id, "Project Name": project_name, "Personnel": person, "Week": week, "Budgeted Hrs": hours}
                    )

        if st.button("Submit Project"):
            new_data = pd.DataFrame(st.session_state.new_project_allocations)
            project_data = pd.concat([project_data, new_data], ignore_index=True)
            try:
                upload_file_to_onedrive(access_token, selected_project_file, project_data)
                st.success("Project data saved successfully!")
            except Exception as e:
                st.error(f"Failed to save project data: {e}")

    elif action == "Update Existing Project":
        st.subheader("Update an Existing Project")
        project_names = project_data["Project Name"].unique()
        selected_project = st.selectbox("Select Project", project_names)
        filtered_data = project_data[project_data["Project Name"] == selected_project]

        for index, row in filtered_data.iterrows():
            budgeted = st.number_input(f"Budgeted Hours for {row['Personnel']} - {row['Week']}", value=row["Budgeted Hrs"])
            spent = st.number_input(f"Spent Hours for {row['Personnel']} - {row['Week']}", value=row["Spent Hrs"])
            project_data.loc[index, "Budgeted Hrs"] = budgeted
            project_data.loc[index, "Spent Hrs"] = spent

        if st.button("Save Updates"):
            try:
                upload_file_to_onedrive(access_token, selected_project_file, project_data)
                st.success("Updates saved successfully!")
            except Exception as e:
                st.error(f"Failed to save updates: {e}")

    # Display summary
    st.subheader("Project Summary")
    st.dataframe(project_data)

if __name__ == "__main__":
    main()
