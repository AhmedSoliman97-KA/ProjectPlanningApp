import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
import requests

# Dropbox App Credentials
APP_KEY = "sx01gc363t4x0ki"
APP_SECRET = "tmaowc3yrmzr4vb"
REFRESH_TOKEN = "sl.u.AFYqauRh9brxevk2GcyZ9aObwvEkAPUBwfhqTjyB71Nf_VAYFmmHfNL1hC8S6HkseyOHC_YeaWv6QEj88Q-xV8A3a2aki0YfpRTyXs9TS3UGzPplWF0_32ejMiDDRhDNOPoJ1-NMcomGC5VRcfppl5V96jmo9qKDrLTNe90jj7FJFrQyPa_CYoWJL8H8883KbeukNFisHTxDx-vVK9C4MV_jqDtvUYKirsFs4Uzfh3yKfZysOEyMZ5sHvfFx6_sguvQneXeqYkhMNsuVOydyr74jLqbfeS_BrQrw5jquNINnBJ3m5UsYo10cDscetXHWHWDY_5uxUdYS4dBS_A4v3TA3d-lxDSDqeDjXreAtCVKIbToDXux-7hVXFVSVDEFayUQlXE7c7F164D8DC6M3WRiK5h-C4gRito6k-YmBJcUMImzFLf8aTuT5PGV0ayLdntbFOCY5D7edPWkBYcXnTBJZtdxFi4aq5ZuMp1PAqSzdZR4FlaHgtaYzlY859i3cn4lQak8TNDo7KTMF9UnGzcQ2D4EFZF9e3-ha9fOYvgrZ75c12SZ7CyezYjt4mhIabfFvl00FSydfEs3LRV-wqFtOglX2MqhmAR0np5niqv57R5c1fzoo0sH6blBM9DRCg2ojbmbg8nQdvxoKyHyQ_utX1pSnl0WF6pLYEXByDtaEHrj3smrNPGjh6rZ9RXyRislxaeGjr2qBt-pH5s3Snq4vh7m_KU7-J8dfxAzfSlmAajqN9ryUQl1JDVIqfotnf7juVskoXHKjsHblIz0aRKhqyYYll5e2sczvxa2xtM2I1y8ZhGe0kyBG6AeeP9wQoxc7dsSEtE9wijyKqaLucgQli58DRzdJMRch6DYny91wtg4-k4ysN0CTe-446VH1-D9S7HV1NZeBo5GSBAMobuzP-pngRi1rxmuce4uZBasgLUKMEWIbZT2eLTQD1JWta9EYu01lCJ_AgSbLx9lDLBlkCadlRCKrszjggwG8Lv6jnavt76lggtccIILMA2SJGVGYCWRBWDw9RnNGJoc0pcjHq8ek3oWPEyicVQ3RQWv5T9prvgq7IuxwOomaR6AC5hl74Ew5Jhx49Cc4Da4lQwPXeI3r_EojVv7vmVHvcTzvLnvmIvvZ-ExypVaMxEvMaNvxotDSJwx7BbdvNag3ADqXsZ8PqYzV-39WWQD-ZKt5ArTxPMkkPx0II5oM94OXuS9NAfQvWRwI57IjpJV3iLDSTa-iLPeL1SpnbcxyEjL-GnGZLQWQBMYtT7ZRcMx_zoAPeQLxQN8bu_nvmuWY7c5M0EkepaLSE_lIyTF6UvVNFERqfcefHOCINa1sEeyAMuFL5yX4hDYsYUl_hwEp8kitpbPbgA-7ZxsLsrQxiM8CEm9Fbea8c1AlnaBy4GZPvdCcPWrpJUg99BoCd_1TL3-mCJXxB_qsvTOVjwTlJD6Z_2wOyOwX2oaP9hPL23zYRRk"

# Dropbox File Paths
PROJECTS_FILE_PATH = "/Project_Data/projects_data_weekly.xlsx"
HR_FILE_LOCAL = "Human Resources.xlsx"

# Function to Refresh Access Token
def refresh_access_token():
    """Refresh Dropbox access token."""
    url = "https://api.dropbox.com/oauth2/token"
    data = {
        "grant_type": "refresh_token",
        "refresh_token": REFRESH_TOKEN,
        "client_id": APP_KEY,
        "client_secret": APP_SECRET,
    }
    response = requests.post(url, data=data)
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        st.error("Failed to refresh Dropbox access token.")
        raise Exception(f"Token Refresh Error: {response.json()}")

# Dropbox Functions with Token Refresh
def download_from_dropbox(file_path):
    """Download a file from Dropbox with token refresh logic."""
    try:
        access_token = refresh_access_token()
        dbx = dropbox.Dropbox(access_token)
        metadata, res = dbx.files_download(file_path)
        return pd.ExcelFile(res.content)
    except dropbox.exceptions.ApiError as e:
        if e.error.is_path() and e.error.get_path().is_not_found():
            return None
        else:
            st.error(f"Error downloading file: {e}")
            raise

def upload_to_dropbox(df, dropbox_path):
    """Upload a DataFrame to Dropbox as an Excel file with token refresh logic."""
    try:
        access_token = refresh_access_token()
        dbx = dropbox.Dropbox(access_token)
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

    # Load Human Resources File
    try:
        hr_excel = pd.ExcelFile(HR_FILE_LOCAL)
    except Exception as e:
        st.error(f"Error loading Human Resources file: {e}")
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

    if action == "Create New Project":
        st.subheader("Create a New Project")
        project_id = st.text_input("Project ID", help="Enter a unique ID for the project.")
        project_name = st.text_input("Project Name", help="Enter the name of the project.")
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)

        # Engineer Selection from Dropdown
        st.subheader("Select Engineers for Allocation")
        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers_data = hr_excel.parse(sheet_name=selected_section)
        engineers = engineers_data["Name"].dropna().tolist()
        selected_engineers = st.multiselect("Choose Engineers", options=engineers, help="Select engineers to allocate hours.")

        allocations = []

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

        # Display Summary Allocation
        if allocations:
            st.subheader("Summary of Allocations")
            allocation_df = pd.DataFrame(allocations)
            st.dataframe(allocation_df)
            total_budgeted = allocation_df["Budgeted Hrs"].sum()
            st.metric("Total Budgeted Hours", total_budgeted)

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

if __name__ == "__main__":
    main()
