import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# Dropbox Settings
ACCESS_TOKEN = "sl.u.AFY76MVKUHTIGBfqYRqzrPnV19H8Jly5Q-tt5U5r4S1GA9eEq4L_Kn32TGZIEbQjrgEzWWVdAsXyRdEfhOge_Z_K6_M3Z3_Sgr8a_PdXLM82FcXlGTEZoTPH4l9w5vyiIhgIkW0Tq2sE1_6_ajKdhlTeNir06bi_r3XcSR_w9A9AB_4i9CEjwmxgRf_w1zc_rN_d1VBRzhbUJn9yjAsut_J5r0I2RovfSWAFJo6RQcyInPygr2RJLbB0CJqqpQ4z09FHQdwtxtaSzhv68KiQWI8YaNMrcW642Pnu6iQm-oNRXzfXLaVjx6uca_H9Uxmcgx1v-J3_BXNJmjVRHCILAr1aRqtlkZlZd6Qz48924Q5i1vt_MY10gfejguPXKsy_omhQ-4W3FeYLXLNwWcv2cZpBtN0TRRhaLXmxpgfEACcGmpoZNWcrbgjF0wSfWag1vBl-p4NmcXneDhi93PLlLvDNPCN0rAazCiCfK710dwSE1Cz7KcdPW-frmu3CJQewcwL6-znetdFTVdbspL4pNKKktsU9DL0rAY-XVymvyTqpBRTD6Db3oCn3MVtIqxUQQPiNTjQ-T1BB5bls3DdS-ToWaYpQyDslzbYg5BBTcG2oVcCa7W4TbY9yc6KdbiiDyU3pxzuXXDE8mgNGwxauJQrdJt992Pb5stjMxKiY5BnHh6va-yleyZuxbkaldXiAqrwUkhJcwOxf7TE6iA-NIQAzsRleD4OIB8SDBDQfMeBtjXAeuG6I91XhyNgdGOYG8cuwBMSKXDrpqT8p1BIIY9zay0rbBEAOzPJBS4NLJxJL7VR12JpVugHLax3r6TzoQhSPuhvS6EMLQ9PWoKaWaN_tYlSj2wpZDO4jRj5nsIIu1KogQtvKwmVyEd9aIVYmnL95w9iOKKoe5XItbQOtLUwele8S_vmEBheVBSYI-ad9fEZBwCq4CJQ9wOyljRYshassdeBFnK7mx37C1Pr8Nc1sZTQzlEU2OGuP3R2ed3G1y9TL2jNm_K0uxAu4lNIzUg3NgF8Os8FfwFeEQ1DhBjzfeN8Qj4ANOx9VAHYHqFbINQNueg12G0g6rZUSGpCFkHS1g5qwfkJ_UXW32KCkDB0Zbw5VppSulisrwMkbz0HEy7c1cIcoJGG8ac-9cwNePtFc_q8vLdbbDyIr0ryusyYPfRHyQmPTyd6-WODowFqoivxC_1sEz87Xt4oEYRIemtdqEWrznsd_EiTvvWEWtqpmIgjK-2MAJl_tFh_x-6MoXYdplrbJFRVPvlWS_32R_4GOoEIncxJ8IYF3c1VtC7F-iOvSBKTDQPKlhLya_xa54bb6PvyqU4R1drmYWrGzpR0JZR39v-PIlOPlvMm0mM7Fn1ZvkVBOl9c0LXFyZP94fn8NYGE1hfubQHnIG0t_fudUCRvnhVBmyOqiy4HkG-EWnKyFvwoN5GArvtvsmn9wHfG6QjynggOHjJx_YS3h9FI"
PROJECTS_FILE_PATH = "/Project_Data/projects_data_weekly.xlsx"
HR_FILE_PATH = "/Project_Data/Human Resources.xlsx"

# Dropbox Functions
def download_from_dropbox(file_path):
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

        # Year and month selection
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)

        # Engineer selection
        st.subheader("Select Engineers for Allocation")
        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers_data = hr_excel.parse(sheet_name=selected_section)
        engineers = engineers_data["Name"].dropna().tolist()
        selected_engineers = st.multiselect("Choose Engineers", options=engineers, help="Select engineers to allocate hours.")

        # Initialize allocations template
        allocations_template = pd.DataFrame(columns=["Month", "Week", "Engineer", "Budgeted Hours"])

        # Generate Tabular Input
        if selected_engineers:
            st.subheader("Allocate Weekly Hours")
            for engineer in selected_engineers:
                allocations_template = allocations_template.append({"Engineer": engineer}, ignore_index=True)

            gb = GridOptionsBuilder.from_dataframe(allocations_template)
            gb.configure_default_column(editable=True)
            gb.configure_column("Month", editable=True)
            gb.configure_column("Week", editable=True)
            gb.configure_column("Budgeted Hours", editable=True)
            grid_options = gb.build()

            response = AgGrid(
                allocations_template,
                gridOptions=grid_options,
                update_mode=GridUpdateMode.MANUAL
            )

            # Process grid response
            allocations = response["data"]

            # Calculate Totals
            total_allocated_hours = allocations["Budgeted Hours"].sum()
            total_allocated_cost = 0

            # Add Engineer Details and Costs
            for _, row in allocations.iterrows():
                engineer_details = engineers_data[engineers_data["Name"] == row["Engineer"]].iloc[0]
                cost_per_hour = pd.to_numeric(engineer_details.get("Cost/Hour", 0), errors='coerce')
                total_allocated_cost += row["Budgeted Hours"] * cost_per_hour

            st.metric("Total Allocated Hours", f"{total_allocated_hours} hrs")
            st.metric("Total Allocated Cost (in $)", f"${total_allocated_cost:,.2f}")

        # Submit Button
        if st.button("Submit Project"):
            new_data = allocations.copy()
            new_data["Project ID"] = project_id
            new_data["Project Name"] = project_name
            new_data["Year"] = selected_year
            new_data["Month"] = selected_month
            new_data["Section"] = selected_section

            projects_data = pd.concat([projects_data, new_data], ignore_index=True)
            upload_to_dropbox(projects_data, PROJECTS_FILE_PATH)
            st.success("Project submitted successfully!")

if __name__ == "__main__":
    main()

