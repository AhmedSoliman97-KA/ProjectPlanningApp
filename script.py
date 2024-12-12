import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
from st_aggrid import AgGrid, GridOptionsBuilder

# Dropbox App Credentials
APP_KEY = "w6nrz3ghlfskn5i"  # Replace with your Dropbox App Key
APP_SECRET = "uq94leubvg2xc23"  # Replace with your Dropbox App Secret
REFRESH_TOKEN = "sl.u.AFbLKe_s1qjfmL48cPaEOhqwX5FK-veyXBJJaxejYMPZgqgj7BXq_7kXOV0IqasmG37Ud3oNXZT11hWbzGkp2x1pHcRj1pP5b3LMwpS1YXLWMjbvoBwOUyvH3I7lVHwfvpgcA9t2LG91yEmK_4qc7PXs1nab4lyGLdGTURmpUSkSPmSUYLd5cR-lVxOn6bTVNQ5RAobQoEfdLFZnVCu4f-CBgDSnkS7fFkgg5rVhH0uPU4zv5ojO_MV2ar_UM2NEl8-HnNML-L7omJtuRYtlUE1hZtlT6wEUC7MniHnNPJaizMw5XSs7Xz1iTKvd-y1L1rhBvPSUZHV-u_Na8VkiRDZQo5FPYLfzhsrGckp1P3l2ZezBPVvJ8i9hQddEtF5j1_nnyMjk4FfYfr5gUX1GGnz1sb7JoY-p_1lhpwKPbYP45dPsEbrLt2bPEturl0oIpJB3sQ0Rg7b4FfeWvTE1pMv8oX88hrJJoh_OqO4vS_jj15DtmTHyJNE4XPc0uFmIAUNf5ZU_i6CEnW3MMzSUhSqCRS1UNDQKHAYFHnTvsdKRQURGQ1Lu0g-YoZl5Xpy9RbivObTi0Q6EWoOfoBG8aWtlsJJUyxSLPukgLeFZlTCzLiSiwd7EX01VWJldK6xR07XCYV4qncI9wknSct2i1FT21pzXdkjie5j4qC8T27IXSM39HG-5e5vH26N7cBxVLqARSMcNsBbMxNPiWGSPUuBaTxX-88oUsi4q8IMpebcaCVEBF6Qs-nYBr87K4g5jP961HXv1uocjD8uraq_Kj2xfiJsYm4rk5rTrgp8EdtszhSxUbkdp0kqiru6U68AisgGsXwTCxQCk1a3arJofjXGM5leYx3xOwaWwmLmDnFgxTGo5goj-RUOShgjmWjicxLjNRkhSNcPe4IzyAvghXZL9njDYuGuyIR7_RUQ5clkSTogg2UQ0_MbIGztcbL_339dk_wHVwvqAyuHW0ckwoH28K6A4Lqsrg-Cd-_G_zpbD9ADVp3KZt1yg_DkKri4XKt3SFIikGlwSszyc7V2CQ3JKIC3vaDVhtjkHdfu0xX9gRZL8zPvNM_0CqebcptV8A03pzAVx5wqdYAuYDdRnI2vALZ5N_sKlOAGze7CrjfXy8Nmzl4Ej54U4O8CHadth-_j3rqHBCKvJGM6MpPw86xowNIen4vK9LDDPdCHftmeNiDJBv8uk78maGKEtq99SRoJF5xPx4dGthhpL7R8ofCaZ1Arql5attyydWAG7eToQH4E155UXbPSLXwLSZSu6HppIHngu_cHNqL-cWyJ8Tka1n33pkWENXhw7XbEqDWEBvFXzN1_BFSlm0pyo8cu3yA6Hjv20B3b2skNa5gJIQmiwvGuC8TKKjT1E69cOigDwAWj5D6qW5mv-QcbKHTkUR3-fRyOVDF84QIKH7wNvMiTWgkb2C7RhWJLl8ozdHM--F9lKg_Rk0P8uHOJsHEqWvFk"  # Replace with your Dropbox Refresh Token
PROJECTS_FILE_PATH = "/Project_Data/projects_data_weekly.xlsx"
HR_FILE_PATH = "/Project_Data/Human Resources.xlsx"


# Function to fetch Dropbox access token dynamically
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
        access_token = get_access_token()
        dbx = dropbox.Dropbox(access_token)
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
        access_token = get_access_token()
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
    try:
        existing_file = download_from_dropbox(file_path)
        if existing_file is None:
            st.warning(f"{file_path} not found in Dropbox. Creating a new file...")
            empty_df = pd.DataFrame(columns=[
                "Project ID", "Project Name", "Personnel", "Week", "Year", "Month",
                "Budgeted Hrs", "Spent Hrs", "Remaining Hrs", "Cost/Hour", "Budgeted Cost",
                "Remaining Cost", "Section", "Category"
            ])
            upload_to_dropbox(empty_df, file_path)
    except Exception as e:
        st.error(f"Error ensuring file exists: {e}")

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

        # Year and month selection
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)

        # Engineer selection
        st.subheader("Select Engineers for Allocation")
        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers_data = hr_excel.parse(sheet_name=selected_section)
        engineers = engineers_data["Name"].dropna().tolist()
        selected_engineers = st.multiselect("Choose Engineers", options=engineers, help="Select engineers to allocate hours.")

        allocations = []
        total_allocated_budget = 0

        if selected_engineers:
            st.subheader("Allocate Weekly Hours")

            # Generate Weeks
            weeks = generate_weeks(selected_year, list(month_name).index(selected_month))

            for engineer in selected_engineers:
                engineer_details = engineers_data[engineers_data["Name"] == engineer].iloc[0]
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

                        total_allocated_budget += budgeted_cost

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
                            "Section": selected_section,
                            "Category": engineer_details.get("Category", "N/A")
                        })

        if allocations:
            st.subheader("Summary of Allocations")
            allocation_df = pd.DataFrame(allocations)
            st.dataframe(allocation_df)
            st.metric("Total Allocated Budget (in $)", f"${total_allocated_budget:,.2f}")

        if st.button("Submit Project"):
            if not project_id.strip() or not project_name.strip():
                st.error("Project ID and Name cannot be empty.")
            elif not allocations:
                st.error("No allocations made.")
            else:
                new_data = pd.DataFrame(allocations)
                upload_to_dropbox(new_data, PROJECTS_FILE_PATH)
                st.success("Project submitted successfully!")

if __name__ == "__main__":
    main()

