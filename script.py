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

# Generate Weeks for Dropdown
def generate_weeks(year, month):
    start_date = datetime(year, month, 1)
    end_date = (start_date + timedelta(days=31)).replace(day=1)
    weeks = []
    while start_date < end_date:
        week_label = start_date.strftime("%Y-%m-%d")
        weeks.append(week_label)
        start_date += timedelta(days=7)
    return weeks

# Main Application
def main():
    st.image("image.png", use_container_width=True)
    st.title("Water & Environment Project Planning")

    ensure_dropbox_projects_file_exists(PROJECTS_FILE_PATH)

    hr_excel = download_from_dropbox(HR_FILE_PATH)
    if hr_excel is None:
        st.error(f"Human Resources file not found in Dropbox at {HR_FILE_PATH}.")
        st.stop()

    hr_sections = hr_excel.sheet_names

    projects_excel = download_from_dropbox(PROJECTS_FILE_PATH)
    if projects_excel is None:
        projects_data = pd.DataFrame(columns=[
            "Project ID", "Project Name", "Personnel", "Week", "Year", "Month",
            "Budgeted Hrs", "Spent Hrs", "Remaining Hrs", "Cost/Hour", "Budgeted Cost",
            "Remaining Cost", "Section", "Category"
        ])
    else:
        projects_data = pd.read_excel(projects_excel)

    action = st.sidebar.radio("Choose an Action", ["Create New Project", "Update Existing Project"])

    if action == "Create New Project":
        st.subheader("Create a New Project")

        project_id = st.text_input("Project ID", help="Enter a unique ID for the project.")
        project_name = st.text_input("Project Name", help="Enter the name of the project.")
        approved_budget = st.number_input("Approved Total Budget (in $)", min_value=0, step=1)

        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers_data = hr_excel.parse(sheet_name=selected_section)
        engineers = engineers_data["Name"].dropna().tolist()

        selected_engineer = st.selectbox("Choose Engineer to Allocate", options=engineers)

        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)

        weeks = generate_weeks(selected_year, list(month_name).index(selected_month))

        allocations_template = pd.DataFrame({
            "Month": [selected_month] * len(weeks),
            "Week": weeks,
            "Budgeted Hours": [0] * len(weeks)
        })

        gb = GridOptionsBuilder.from_dataframe(allocations_template)
        gb.configure_default_column(editable=True)
        grid_options = gb.build()

        response = AgGrid(allocations_template, gridOptions=grid_options, update_mode=GridUpdateMode.MANUAL)
        edited_data = pd.DataFrame(response["data"])

        if st.button("Submit Project"):
            engineer_details = engineers_data[engineers_data["Name"] == selected_engineer].iloc[0]
            section = engineer_details.get("Section", "Unknown")
            category = engineer_details.get("Category", "N/A")
            cost_per_hour = pd.to_numeric(engineer_details.get("Cost/Hour", 0), errors='coerce')

            for _, row in edited_data.iterrows():
                week = row["Week"]
                budgeted_hours = row["Budgeted Hours"]

                if budgeted_hours > 0:
                    remaining_hours = budgeted_hours
                    budgeted_cost = budgeted_hours * cost_per_hour

                    new_entry = {
                        "Project ID": project_id,
                        "Project Name": project_name,
                        "Personnel": selected_engineer,
                        "Week": week,
                        "Year": selected_year,
                        "Month": selected_month,
                        "Budgeted Hrs": budgeted_hours,
                        "Spent Hrs": 0,
                        "Remaining Hrs": remaining_hours,
                        "Cost/Hour": cost_per_hour,
                        "Budgeted Cost": budgeted_cost,
                        "Remaining Cost": budgeted_cost,
                        "Section": section,
                        "Category": category
                    }
                    projects_data = projects_data.append(new_entry, ignore_index=True)

            upload_to_dropbox(projects_data, PROJECTS_FILE_PATH)
            st.success("Project submitted successfully!")

    if action == "Update Existing Project":
        st.subheader("Update an Existing Project")

        selected_section = st.selectbox("Choose a Section", hr_sections)
        filtered_projects = projects_data[projects_data["Section"] == selected_section]

        if filtered_projects.empty:
            st.warning(f"No projects found for the section: {selected_section}.")
            st.stop()

        selected_project = st.selectbox("Choose a Project", filtered_projects["Project Name"].unique())
        project_details = filtered_projects[filtered_projects["Project Name"] == selected_project]

        st.subheader("Current Allocations for Selected Project")
        st.dataframe(project_details)

        current_budgeted_hours = project_details["Budgeted Hrs"].sum()
        current_budgeted_cost = project_details["Budgeted Cost"].sum()

        st.metric("Current Budgeted Hours", f"{current_budgeted_hours} hrs")
        st.metric("Current Budgeted Cost", f"${current_budgeted_cost:,.2f}")

        updated_rows = []
        for _, row in project_details.iterrows():
            updated_budgeted = st.number_input(
                f"Updated Budgeted Hours ({row['Week']})",
                min_value=0,
                value=int(row["Budgeted Hrs"]),
                step=1,
                key=f"update_budgeted_{row['Personnel']}_{row['Week']}"
            )

            cost_per_hour = row["Cost/Hour"]
            budgeted_cost = updated_budgeted * cost_per_hour

            updated_rows.append({
                "Project ID": row["Project ID"],
                "Project Name": row["Project Name"],
                "Personnel": row["Personnel"],
                "Week": row["Week"],
                "Year": row["Year"],
                "Month": row["Month"],
                "Budgeted Hrs": updated_budgeted,
                "Spent Hrs": row["Spent Hrs"],
                "Remaining Hrs": updated_budgeted - row["Spent Hrs"],
                "Cost/Hour": cost_per_hour,
                "Budgeted Cost": budgeted_cost,
                "Remaining Cost": budgeted_cost - (row["Spent Hrs"] * cost_per_hour),
                "Section": row["Section"],
                "Category": row["Category"]
            })

        updated_data = pd.DataFrame(updated_rows)
        st.subheader("Updated Allocations for Selected Project")
        st.dataframe(updated_data)

        updated_budgeted_hours = updated_data["Budgeted Hrs"].sum()
        updated_budgeted_cost = updated_data["Budgeted Cost"].sum()

        st.metric("Updated Budgeted Hours", f"{updated_budgeted_hours} hrs")
        st.metric("Updated Budgeted Cost", f"${updated_budgeted_cost:,.2f}")

        if st.button("Save Updates"):
            updated_data["Composite Key"] = (
                updated_data["Project ID"] + "_" +
                updated_data["Project Name"] + "_" +
                updated_data["Personnel"] + "_" +
                updated_data["Week"]
            )
            projects_data["Composite Key"] = (
                projects_data["Project ID"] + "_" +
                projects_data["Project Name"] + "_" +
                projects_data["Personnel"] + "_" +
                projects_data["Week"]
            )
            remaining_data = projects_data[~projects_data["Composite Key"].isin(updated_data["Composite Key"])]
            final_data = pd.concat([remaining_data, updated_data], ignore_index=True)
            final_data.drop(columns=["Composite Key"], inplace=True)

            upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
            st.success(f"Updates to '{selected_project}' saved successfully!")

if __name__ == "__main__":
    main()

