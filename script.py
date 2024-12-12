import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
from st_aggrid import AgGrid, GridOptionsBuilder

# Dropbox App Credentials
APP_KEY = "w6nrz3ghlfskn5i"  # Replace with your Dropbox App Key
APP_SECRET = "uq94leubvg2xc23"  # Replace with your Dropbox App Secret
REFRESH_TOKEN = "sl.CCcuic-FxPSYJgwmzuuBvhTvAMmm4RppC__vkceXVBQh0GvMT-Hm9ttKXQnRpC3f5IbX5oFQOHgDSUF6wJO5lBGQTHLcfU5d0PxMVguWhAgAzpp3f2zpQ8L-MKwrnX9W1R_5D7_Cb0kC-Jp5bbLuLUg"  # Replace with your Dropbox Refresh Token
PROJECTS_FILE_PATH = "/Project_Data/projects_data_weekly.xlsx"
HR_FILE_PATH = "/Project_Data/Human Resources.xlsx"

# Function to Get Access Token
def get_access_token():
    """Fetch a new access token using the refresh token."""
    url = "https://api.dropboxapi.com/oauth2/token"
    data = {
        "grant_type": "refresh_token",
        "refresh_token": REFRESH_TOKEN,
    }
    auth = (APP_KEY, APP_SECRET)
    response = requests.post(url, data=data, auth=auth)
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        st.error(f"Error refreshing Dropbox token: {response.json()}")
        raise Exception(f"Token refresh failed: {response.text}")
        
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

        # Engineer selection with dropdown filter
        st.subheader("Filter Engineers by Section")
        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers_data = hr_excel.parse(sheet_name=selected_section)
        engineers = engineers_data["Name"].dropna().tolist()
        filtered_engineers = st.multiselect("Select Engineers to Include", options=engineers)

        # Tabular input for allocations
        if filtered_engineers:
            st.subheader("Allocate Weekly Hours")
            allocations_template = pd.DataFrame(columns=[
                "Month", "Week", "Engineer", "Budgeted Hours"
            ])
            allocations_template = pd.concat(
                [allocations_template] +
                [pd.DataFrame({"Engineer": [engineer]}) for engineer in filtered_engineers],
                ignore_index=True
            )

            gb = GridOptionsBuilder.from_dataframe(allocations_template)
            gb.configure_default_column(editable=True)
            grid_options = gb.build()
            response = AgGrid(allocations_template, gridOptions=grid_options, update_mode='MANUAL')
            updated_data = pd.DataFrame(response['data'])

            # Calculate total allocated cost
            updated_data["Cost/Hour"] = updated_data["Engineer"].map(
                lambda eng: engineers_data.loc[engineers_data["Name"] == eng, "Cost/Hour"].values[0]
                if not engineers_data[engineers_data["Name"] == eng].empty else 0
            )
            updated_data["Budgeted Cost"] = updated_data["Budgeted Hours"] * updated_data["Cost/Hour"]
            total_allocated_cost = updated_data["Budgeted Cost"].sum()

            # Display allocation summary
            st.subheader("Summary of Allocations")
            st.dataframe(updated_data)

            # Display totals
            st.metric("Total Allocated Budget (in $)", f"${total_allocated_cost:,.2f}")
            st.metric("Approved Total Budget (in $)", f"${approved_budget:,.2f}")
            st.metric("Remaining Budget (in $)", f"${approved_budget - total_allocated_cost:,.2f}")

            # Submit button
            if st.button("Submit Project"):
                # Save to projects data and upload to Dropbox
                updated_data["Composite Key"] = (
                    project_id + "_" + project_name + "_" + updated_data["Engineer"] + "_" + updated_data["Week"]
                )
                projects_data["Composite Key"] = (
                    projects_data["Project ID"] + "_" +
                    projects_data["Project Name"] + "_" +
                    projects_data["Personnel"] + "_" +
                    projects_data["Week"]
                )
                updated_projects = projects_data[~projects_data["Composite Key"].isin(updated_data["Composite Key"])]
                final_data = pd.concat([updated_projects, updated_data], ignore_index=True)
                final_data.drop(columns=["Composite Key"], inplace=True)
                upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
                st.success("Project submitted successfully!")

if __name__ == "__main__":
    main()

