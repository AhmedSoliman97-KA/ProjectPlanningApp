import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# Dropbox Settings
ACCESS_TOKEN = "sl.u.AFZ6CrsoS_Cy_oTTTdpz-o0rfwNEOpKp3jaGTL6CvN4y_OUMQHaVNo8S_1iPpMUvw9Jkksn4uvPpBdDgnayOEBPieYZ1jMjzARiggLDSYwdCO9Jn41-GiRS5d3kN61j5SJrntwTITFUsR-CKNUrlz8NlvCq0E26MK-SGYVmcrLGlvnpm_c6j3AvIR0Ps6QfavYwgorKEdmRw0FXqpXPnfRJP42RIYX4lK6NZxj_8WvNf9tR2amiFKWgOAg5JqtNR0ro593A41bD-xQJBw0laaPxTRNYJLWMaHcR_xp9R6Tq1x71y5gctX-eNzIfOLuzrggo7HS-ZtWJLfL2T12z2jOZP6-G3u2KyMPXkkk8yxfv0ySFtdnQuUj3i9DbyzyB3inuzAXurR12IZY7hvAiAlvXkNrTAQ899FWqHf060mOQxXQVdk8xFh2QsZb2FKc7sAF5G_1cN0nAx2oIqlqmN97IUr98dF7YRFxAns-VJvw2XW2PzF3kswDdCyN873X1Hl-vfwV7UkWq45p408hMDAyMeZq0EYpsP1tiT9-OqK7BzW5OveBmnAu3OMXWhg_r8j78zlmquACVu9NDk9hFqSTj04IA4VMZcCMoByspqzlL0cE6czAXnaNAXs2luCcEyVbnFR_csihS4TRYw_x6Td8Kc0FWCOkwbFy6WUc6iV2rb3r8aucojfHWNR5bG6C1NS4VK98MxVhsBCr7p2aGN3hHDDFmH8jV2Npc3PsEfDDXuJgxN8ARefuXFcCTDS54TK3rZmk0tRxrKIe2SPxLVijLv3Fx0CvszH0tjUTAdf1yMn6j1aCBRVA1a4VAD7QvqdhtBG9o2P7keCbmN-Y1U8W-CTJuAgQv6vfloGDVNQTKU2gG0JZx2wDhkavUhFyDIt37WyS2qxhSMe0RlEWTnWfNOosOvLJUMojcbcBEvzDhj95z1TUGjJyVU8JhldogBSzKp5FM0Af60NtVJ4DrmuS3tkledd4lKq0YCpWArUqEzFjjblq8iobP98bZzu_s9BEDLIvo00s4yKBjqeP68ic3lFrIcE9dlMkbkV_KQIMEuI3n7lA1vm1oJwBo6cCfG6xFsKXOKrA7VixD9fDmEFdQsbHC3tQaWI0I95UqQImbvyFXuAa2QPbvqoCOfvHSL-u6oh9_Hf2ya7CNPpNG5wy8z3P7Blygd6PtDNtgB3fViBIR7X1STbr__DPGdZX8Ow9WP9DHP9G3yndR52nOEe1vGe94q8vyVNJbqM9cMjry0yaxUmIzPmobg_zAzbkJnQx_Ly4joVg8rbsINtPl_eBaJOtiTyQ718O7k_J8UppxoezxCTt7Tt6NgWKhL1rpj7da4orJSpgo-DBRBOFNKgWVO1OagGFbUMGwvX87PIkIwrGD1TWnUjO4y_o9IOtxyAG4o5o8v9M7n9fDl2c-qCTlcdWsb-b3Nvx5pP_7hD8ar6f2zv1-q7tjwFNn06Wl_0cw"
PROJECTS_FILE_PATH = "/Project_Data/projects_data_weekly.xlsx"
HR_FILE_PATH = "/Project_Data/Human Resources.xlsx"

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

# Generate Weeks for a Given Month
def generate_weeks(year, months):
    weeks = []
    for month in months:
        start_date = datetime(year, list(month_name).index(month), 1)
        end_date = (start_date + timedelta(days=31)).replace(day=1)
        while start_date < end_date:
            week_label = start_date.strftime("%Y-%m-%d")
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
        selected_months = st.multiselect("Months", list(month_name)[1:], default=[month_name[datetime.now().month]])

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
            weeks = generate_weeks(selected_year, selected_months)

            for engineer in selected_engineers:
                # Fetch engineer details from Human Resources file
                engineer_details = engineers_data[engineers_data["Name"] == engineer].iloc[0]
                section = engineer_details.get("Section", "Unknown")
                category = engineer_details.get("Category", "N/A")
                cost_per_hour = pd.to_numeric(engineer_details.get("Cost/Hour", 0), errors='coerce')
                cost_per_hour = cost_per_hour if not pd.isna(cost_per_hour) else 0

                st.markdown(f"### Engineer: {engineer}")

                col_list_budgeted = st.columns(len(weeks))

                budgeted_hours_inputs = {}

                for idx, (week_label, col) in enumerate(zip(weeks, col_list_budgeted)):
                    with col:
                        budgeted_hours_inputs[week_label] = st.number_input(
                            f"Budgeted Hours ({week_label})", min_value=0, step=1, key=f"budgeted_{engineer}_{week_label}"
                        )

                    budgeted_hours = budgeted_hours_inputs.get(week_label, 0)
                    budgeted_cost = budgeted_hours * cost_per_hour

                    allocations.append({
                        "Project ID": project_id,
                        "Project Name": project_name,
                        "Personnel": engineer,
                        "Week": week_label,
                        "Year": selected_year,
                        "Month": ", ".join(selected_months),
                        "Budgeted Hrs": budgeted_hours,
                        "Remaining Hrs": budgeted_hours,
                        "Cost/Hour": cost_per_hour,
                        "Budgeted Cost": budgeted_cost,
                        "Remaining Cost": budgeted_cost,
                        "Section": section,
                        "Category": category
                    })

        # Display Allocation Summary
        if allocations:
            st.subheader("Summary of Allocations")
            allocation_df = pd.DataFrame(allocations)
            st.dataframe(allocation_df)

            st.metric("Total Allocated Budget (in $)", f"${total_allocated_budget:,.2f}")

            if st.button("Submit Project"):
                if not project_id.strip() or not project_name.strip():
                    st.error("Project ID and Name cannot be empty.")
                else:
                    new_data = pd.DataFrame(allocations)
                    upload_to_dropbox(new_data, PROJECTS_FILE_PATH)
                    st.success("Project submitted successfully!")

    if action == "Update Existing Project":
        st.subheader("Update an Existing Project")

        # Step 1: Section Selection
        st.subheader("Select Section")
        selected_section = st.selectbox("Choose a Section", hr_sections)

        # Filter Projects by Selected Section
        filtered_projects = projects_data[projects_data["Section"] == selected_section]

        if filtered_projects.empty:
            st.warning(f"No projects found for the section: {selected_section}.")
            st.stop()

        # Step 2: Select Project
        st.subheader("Select a Project")
        selected_project = st.selectbox("Choose a Project", filtered_projects["Project Name"].unique())
        project_details = filtered_projects[filtered_projects["Project Name"] == selected_project]

        # Step 3: Display Current Allocations for the Selected Project
        st.subheader("Current Allocations for Selected Project")
        st.dataframe(project_details)

        # Step 4: Select Engineers
        st.subheader("Select Engineers to Update Allocations")
        engineer_options = list(project_details["Personnel"].unique()) + ["Add New Engineer"]
        selected_engineers = st.multiselect("Choose Engineers", engineer_options, default=engine_options[:-1])

        updated_rows = []
        # Add New Engineer Logic
        if "Add New Engineer" in selected_engineers:
            new_engineers = st.multiselect("Choose New Engineers", options=hr_excel.parse(sheet_name=selected_section)["Name"].dropna().tolist())
                        for engineer in new_engineers:
                if engineer not in project_details["Personnel"].values:
                    project_details = pd.concat([project_details, pd.DataFrame([{
                        "Project ID": selected_project,
                        "Project Name": project_details.iloc[0]["Project Name"],
                        "Personnel": engineer,
                        "Week": None,
                        "Year": None,
                        "Month": None,
                        "Budgeted Hrs": 0,
                        "Spent Hrs": 0,
                        "Remaining Hrs": 0,
                        "Cost/Hour": 0,
                        "Budgeted Cost": 0,
                        "Remaining Cost": 0,
                        "Section": selected_section,
                        "Category": "N/A"
                    }])])

        # Step 5: Update Allocations for Each Selected Engineer
        st.subheader("Update Allocations")

        for engineer in selected_engineers:
            engineer_details = project_details[project_details["Personnel"] == engineer]
            st.markdown(f"## Engineer: {engineer}")

            weeks = engineer_details["Week"].unique()
            if len(weeks) == 0:
                st.warning(f"No weeks available for engineer {engineer}.")
                continue

            col_list_budgeted = st.columns(len(weeks))
            col_list_spent = st.columns(len(weeks))

            updated_budgeted_inputs = {}
            updated_spent_inputs = {}

            st.markdown("### Budgeted Hours")
            for idx, (week, col) in enumerate(zip(weeks, col_list_budgeted)):
                existing_allocation = engineer_details[engineer_details["Week"] == week].iloc[0]
                with col:
                    widget_key_budgeted = f"updated_budgeted_{selected_section}_{engineer}_{week}"
                    updated_budgeted_inputs[week] = st.number_input(
                        f"Budgeted ({week})",
                        min_value=0,
                        value=int(existing_allocation.get("Budgeted Hrs", 0)),
                        step=1,
                        key=widget_key_budgeted
                    )

            st.markdown("### Spent Hours")
            for idx, (week, col) in enumerate(zip(weeks, col_list_spent)):
                existing_allocation = engineer_details[engineer_details["Week"] == week].iloc[0]
                with col:
                    widget_key_spent = f"updated_spent_{selected_section}_{engineer}_{week}"
                    updated_spent_inputs[week] = st.number_input(
                        f"Spent ({week})",
                        min_value=0,
                        value=int(existing_allocation.get("Spent Hrs", 0)),
                        step=1,
                        key=widget_key_spent
                    )

            for week in weeks:
                updated_budgeted = updated_budgeted_inputs.get(week, 0)
                updated_spent = updated_spent_inputs.get(week, 0)
                existing_allocation = engineer_details[engineer_details["Week"] == week].iloc[0] if not engineer_details.empty else {}
                cost_per_hour = existing_allocation.get("Cost/Hour", 0)
                budgeted_cost = updated_budgeted * cost_per_hour
                spent_cost = updated_spent * cost_per_hour
                remaining_cost = budgeted_cost - spent_cost

                updated_rows.append({
                    "Project ID": existing_allocation.get("Project ID", selected_project),
                    "Project Name": existing_allocation.get("Project Name", ""),
                    "Personnel": engineer,
                    "Week": week,
                    "Year": existing_allocation.get("Year", selected_year),
                    "Month": existing_allocation.get("Month", ", ".join(selected_months)),
                    "Budgeted Hrs": updated_budgeted,
                    "Spent Hrs": updated_spent,
                    "Remaining Hrs": updated_budgeted - updated_spent,
                    "Cost/Hour": cost_per_hour,
                    "Budgeted Cost": budgeted_cost,
                    "Remaining Cost": remaining_cost,
                    "Section": selected_section,
                    "Category": existing_allocation.get("Category", "N/A")
                })

        st.subheader("Updated Allocations")
        updated_df = pd.DataFrame(updated_rows)
        st.dataframe(updated_df)

        if st.button("Save Updates"):
            updated_rows_df = pd.DataFrame(updated_rows)
            updated_rows_df["Composite Key"] = (
                updated_rows_df["Project ID"] + "_" +
                updated_rows_df["Project Name"] + "_" +
                updated_rows_df["Personnel"] + "_" +
                updated_rows_df["Week"]
            )
            projects_data["Composite Key"] = (
                projects_data["Project ID"] + "_" +
                projects_data["Project Name"] + "_" +
                projects_data["Personnel"] + "_" +
                projects_data["Week"]
            )
            remaining_data = projects_data[~projects_data["Composite Key"].isin(updated_rows_df["Composite Key"])]
            final_data = pd.concat([remaining_data, updated_rows_df], ignore_index=True)
            final_data.drop(columns=["Composite Key"], inplace=True)

            upload_to_dropbox(final_data, PROJECTS_FILE_PATH)
            st.success(f"Updates to '{selected_project}' saved successfully!")

if __name__ == "__main__":
    main()


