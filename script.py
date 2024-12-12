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

        # Year and month selection
        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)

        # Engineer selection
        st.subheader("Select Engineers for Allocation")
        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers_data = hr_excel.parse(sheet_name=selected_section)
        engineers = engineers_data["Name"].dropna().tolist()

        # Tabular Input for Allocations
        allocations_template = pd.DataFrame(columns=[
            "Engineer", "Month", "Week", "Budgeted Hours"
        ])
        gb = GridOptionsBuilder.from_dataframe(allocations_template)
        gb.configure_default_column(editable=True)
        gb.configure_column("Engineer", editable=True, cellEditor='agSelectCellEditor', cellEditorParams={"values": engineers})
        grid_options = gb.build()

        response = AgGrid(allocations_template, gridOptions=grid_options, update_mode=GridUpdateMode.MANUAL)
        updated_data = pd.DataFrame(response['data'])

        # Add Section and Category
        updated_data = updated_data.merge(
            engineers_data[["Name", "Category", "Cost/Hour"]].rename(columns={"Name": "Engineer"}),
            on="Engineer", how="left"
        )

        # Calculate Costs and Remaining Hours
        updated_data["Budgeted Cost"] = updated_data["Budgeted Hours"] * updated_data["Cost/Hour"]
        updated_data["Remaining Cost"] = updated_data["Budgeted Cost"]

        # Display Summary
        st.subheader("Summary of Allocations")
        st.dataframe(updated_data)
        total_allocated_budget = updated_data["Budgeted Cost"].sum()
        st.metric("Total Allocated Budget (in $)", f"${total_allocated_budget:,.2f}")
        st.metric("Approved Total Budget (in $)", f"${approved_budget:,.2f}")
        st.metric("Difference (Remaining/Over-Allocated)", f"${approved_budget - total_allocated_budget:,.2f}")

        # Submit Button
        if st.button("Submit Project"):
            if not project_id.strip():
                st.error("Project ID cannot be empty. Please enter a valid Project ID.")
            elif not project_name.strip():
                st.error("Project Name cannot be empty. Please enter a valid Project Name.")
            elif updated_data.empty:
                st.error("No allocations have been made. Please allocate hours before submitting.")
            else:
                updated_data["Project ID"] = project_id
                updated_data["Project Name"] = project_name
                updated_data["Year"] = selected_year
                updated_data["Section"] = selected_section

                # Composite Key for Deduplication
                updated_data["Composite Key"] = (
                    updated_data["Project ID"] + "_" +
                    updated_data["Project Name"] + "_" +
                    updated_data["Engineer"] + "_" +
                    updated_data["Week"]
                )
                projects_data["Composite Key"] = (
                    projects_data["Project ID"] + "_" +
                    projects_data["Project Name"] + "_" +
                    projects_data["Personnel"] + "_" +
                    projects_data["Week"]
                )

                # Merge and Save Data
                updated_projects = projects_data[~projects_data["Composite Key"].isin(updated_data["Composite Key"])]
                final_data = pd.concat([updated_projects, updated_data], ignore_index=True)
                final_data.drop(columns=["Composite Key"], inplace=True)
                upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
                st.success("Project submitted successfully!")

    # Update Existing Project
    if action == "Update Existing Project":
        st.subheader("Update an Existing Project")

        # Step 1: Section Selection
        selected_section = st.selectbox("Choose a Section", hr_sections)

        # Filter Projects by Selected Section
        filtered_projects = projects_data[projects_data["Section"] == selected_section]

        if filtered_projects.empty:
            st.warning(f"No projects found for the section: {selected_section}.")
            st.stop()

        # Step 2: Select Project
        selected_project = st.selectbox("Choose a Project", filtered_projects["Project Name"].unique())
        project_details = filtered_projects[filtered_projects["Project Name"] == selected_project]

        # Display Current Allocations
        st.subheader("Current Allocations for Selected Project")
        st.dataframe(project_details)

        # Tabular Input for Updating Allocations
        gb = GridOptionsBuilder.from_dataframe(project_details)
        gb.configure_default_column(editable=True)
        grid_options = gb.build()
        response = AgGrid(project_details, gridOptions=grid_options, update_mode=GridUpdateMode.MANUAL)
        updated_project_data = pd.DataFrame(response['data'])

        # Submit Updates
        if st.button("Save Updates"):
            updated_project_data["Section"] = selected_section
            updated_project_data = updated_project_data.merge(
                engineers_data[["Name", "Category", "Cost/Hour"]].rename(columns={"Name": "Personnel"}),
                on="Personnel", how="left"
            )
            updated_project_data["Category"] = updated_project_data["Category"].fillna("N/A")

            updated_project_data["Composite Key"] = (
                updated_project_data["Project ID"] + "_" +
                updated_project_data["Project Name"] + "_" +
                updated_project_data["Personnel"] + "_" +
                updated_project_data["Week"]
            )
            projects_data["Composite Key"] = (
                projects_data["Project ID"] + "_" +
                projects_data["Project Name"] + "_" +
                projects_data["Personnel"] + "_" +
                projects_data["Week"]
            )

            # Merge Updated Data with Existing Data
            remaining_data = projects_data[~projects_data["Composite Key"].isin(updated_project_data["Composite Key"])]
            final_data = pd.concat([remaining_data, updated_project_data], ignore_index=True)
            final_data.drop(columns=["Composite Key"], inplace=True)

            # Upload the updated data to Dropbox
            upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
            st.success(f"Updates to '{selected_project}' saved successfully!")

    # End of Main Application Logic
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Utility Options")
    if st.sidebar.button("Download Project Data"):
        if projects_data.empty:
            st.warning("No project data is available to download.")
        else:
            with pd.ExcelWriter("projects_data_weekly.xlsx", engine="openpyxl") as writer:
                projects_data.to_excel(writer, index=False)
            with open("projects_data_weekly.xlsx", "rb") as file:
                st.sidebar.download_button(
                    label="Download Project Data",
                    data=file,
                    file_name="projects_data_weekly.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()

