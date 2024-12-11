import streamlit as st
import pandas as pd
import requests
import msal
from io import BytesIO

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

# Download a file from OneDrive
def download_file_from_onedrive(access_token, file_path):
    url = f"https://1drv.ms/x/c/dc92d57d369f0c6e/Ee1_xPk8fv9HihGaGmX2fx0BNEiTAKFIPfhsneI7O9425g?download=1"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return pd.read_excel(BytesIO(response.content), engine="openpyxl")
    else:
        raise Exception(f"Failed to download file: {response.status_code} - {response.text}")

# Upload a file to OneDrive
def upload_file_to_onedrive(access_token, file_path, data_frame):
    url = f"https://1drv.ms/f/c/dc92d57d369f0c6e/Eq1SL0aIeV9PjMWWGQH5-OMBu6yVnvzHoyaYu7HQFD8Ujw?e=Zt6da7"
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

def main():
    st.title("KA Planning - Project Management")

    # Authenticate and get access token
    try:
        access_token = get_access_token()
        st.success("Successfully authenticated with OneDrive.")
    except Exception as e:
        st.error(f"Authentication failed: {e}")
        return

    # Example usage: Load Human Resources data
    try:
        hr_data = download_file_from_onedrive(access_token, "Human Resources.xlsx")
        st.write("Human Resources Data:")
        st.dataframe(hr_data)
    except Exception as e:
        st.error(f"Failed to load Human Resources data: {e}")

    # Example usage: Load and update project data
    try:
        project_data = download_file_from_onedrive(access_token, "projects_data.xlsx")
    except Exception:
        st.warning("No existing project data found. Initializing a new file.")
        project_data = pd.DataFrame(columns=["Project ID", "Project Name", "Personnel", "Week", "Budgeted Hrs", "Spent Hrs"])

    if st.button("Save Updates"):
        try:
            upload_file_to_onedrive(access_token, "projects_data.xlsx", project_data)
            st.success("Project data saved successfully!")
        except Exception as e:
            st.error(f"Failed to save project data: {e}")

if __name__ == "__main__":
    main()
