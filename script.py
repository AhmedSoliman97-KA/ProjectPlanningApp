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
    weeks = {}
    for month in months:
        start_date = datetime(year, list(month_name).index(month), 1)
        end_date = (start_date + timedelta(days=31)).replace(day=1)
        while start_date < end_date:
            week_label = start_date.strftime("%Y-%m-%d")
            weeks[week_label] = start_date
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
                engineer_details = engineers_data[engineers_data["Name"] == engineer].iloc[0]
                cost_per_hour = pd.to_numeric(engineer_details.get("Cost/Hour", 0), errors="coerce")
                st.markdown(f"### Engineer: {engineer}")
            
                # Generate weeks horizontally
                col_list = st.columns(len(weeks))
                week_labels = list(weeks.keys())
                budgeted_hours_inputs = {}
            
                for idx, (week_label, col) in enumerate(zip(week_labels, col_list)):
                    with col:
                        budgeted_hours_inputs[week_label] = st.number_input(
                            f"Budgeted ({week_label})",
                            min_value=0,
                            step=1,
                            key=f"{engineer}_{week_label}_budgeted"
                        )
            
                # Save allocation
                # Generate weeks horizontally with unique keys
                for engineer in selected_engineers:
                    engineer_details = engineers_data[engineers_data["Name"] == engineer].iloc[0]
                    cost_per_hour = pd.to_numeric(engineer_details.get("Cost/Hour", 0), errors="coerce")
                    st.markdown(f"### Engineer: {engineer}")
                
                    # Create columns for weeks
                    col_list = st.columns(len(weeks))
                    budgeted_hours_inputs = {}
                
                    for idx, (week_label, col) in enumerate(zip(weeks.keys(), col_list)):
                        with col:
                            # Unique key using engineer and week
                            widget_key = f"budgeted_{engineer}_{week_label}"
                            budgeted_hours_inputs[week_label] = st.number_input(
                                f"{week_label}",
                                min_value=0,
                                step=1,
                                key=widget_key
                            )
                            budgeted_cost = budgeted_hours_inputs[week_label] * cost_per_hour
                
                            # Append allocation data
                            if budgeted_hours_inputs[week_label] > 0:
                                allocations.append({
                                    "Project ID": project_id,
                                    "Project Name": project_name,
                                    "Personnel": engineer,
                                    "Week": week_label,
                                    "Year": selected_year,
                                    "Month": selected_month,
                                    "Budgeted Hrs": budgeted_hours_inputs[week_label],
                                    "Remaining Hrs": budgeted_hours_inputs[week_label],
                                    "Cost/Hour": cost_per_hour,
                                    "Budgeted Cost": budgeted_cost,
                                    "Remaining Cost": budgeted_cost,
                                    "Section": engineer_details.get("Section", "Unknown"),
                                    "Category": engineer_details.get("Category", "N/A")
                                })

            
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
                            "Remaining Cost": remaining_cost,
                            "Section": engineer_details.get("Section", "Unknown"),
                            "Category": engineer_details.get("Category", "N/A")
                        })


        # Display Allocation Summary and Comparison
        if allocations:
            st.subheader("Summary of Allocations")
            allocation_df = pd.DataFrame(allocations)
            st.dataframe(allocation_df)

            # Display Total Allocated Budget
            st.metric("Total Allocated Budget (in $)", f"${total_allocated_budget:,.2f}")
            st.metric("Approved Total Budget (in $)", f"${approved_budget:,.2f}")
            st.metric("Difference (Remaining/Over-Allocated)", f"${approved_budget - total_allocated_budget:,.2f}")
            st.metric("Total Allocated Cost (in $)", f"${total_allocated_budget:,.2f}")

        # Submit Button
        if st.button("Submit Project"):
            if not project_id.strip():
                st.error("Project ID cannot be empty. Please enter a valid Project ID.")
            elif not project_name.strip():
                st.error("Project Name cannot be empty. Please enter a valid Project Name.")
            elif not allocations:
                st.error("No allocations have been made. Please allocate hours before submitting.")
            else:
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
    
        # Metrics for Current Allocations
        current_budgeted_hours = project_details["Budgeted Hrs"].sum()
        current_spent_hours = project_details["Spent Hrs"].sum()
        current_budgeted_cost = project_details["Budgeted Cost"].sum()
        current_spent_cost = (project_details["Spent Hrs"] * project_details["Cost/Hour"]).sum()
    
        st.metric("Current Budgeted Hours", f"{current_budgeted_hours} hrs")
        st.metric("Current Spent Hours", f"{current_spent_hours} hrs")
        st.metric("Current Budgeted Cost", f"${current_budgeted_cost:,.2f}")
        st.metric("Current Spent Cost", f"${current_spent_cost:,.2f}")
    
        # Step 4: Select Engineers
        st.subheader("Select Engineers to Update Allocations")
        engineer_options = list(project_details["Personnel"].unique()) + ["Add New Engineer"]
        selected_engineers = st.multiselect("Choose Engineers", engineer_options, default=engineer_options[:-1])
    
        # Add New Engineer Logic
        if "Add New Engineer" in selected_engineers:
            st.subheader("Add New Engineer")
            new_engineers = st.multiselect("Choose New Engineers", options=hr_excel.parse(sheet_name=selected_section)["Name"].dropna().tolist())
    
            for new_engineer in new_engineers:
                if new_engineer not in project_details["Personnel"].values:
                    project_details = pd.concat([project_details, pd.DataFrame([{
                        "Project ID": project_details.iloc[0]["Project ID"],
                        "Project Name": project_details.iloc[0]["Project Name"],
                        "Personnel": new_engineer,
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
                        "Category": None
                    }])])
    
        # Step 5: Update Allocations for Each Selected Engineer
        st.subheader("Update Allocations")
        updated_rows = []
    
        for engineer in selected_engineers:
            engineer_details = project_details[project_details["Personnel"] == engineer]
            st.markdown(f"## Engineer: {engineer}")
        
            # Generate weeks horizontally
            weeks = engineer_details["Week"].unique()
            col_list = st.columns(len(weeks))
            updated_budgeted_inputs = {}
        
            col_list = st.columns(len(weeks))
            for idx, (week_label, col) in enumerate(zip(week_labels, col_list)):
                with col:
                    budgeted_hours_inputs[week_label] = st.number_input(
                        f"Budgeted ({week_label})",
                        min_value=0,
                        step=1,
                        key=f"{engineer}_{week_label}_budgeted"
                    )
        
            # Save updated rows
            for week, updated_budgeted in updated_budgeted_inputs.items():
                existing_allocation = engineer_details[engineer_details["Week"] == week].iloc[0]
                cost_per_hour = existing_allocation["Cost/Hour"]
                budgeted_cost = updated_budgeted * cost_per_hour
                remaining_cost = budgeted_cost - existing_allocation["Spent Hrs"] * cost_per_hour
        
                updated_rows.append({
                    "Project ID": existing_allocation["Project ID"],
                    "Project Name": existing_allocation["Project Name"],
                    "Personnel": engineer,
                    "Week": week,
                    "Year": existing_allocation["Year"],
                    "Month": existing_allocation["Month"],
                    "Budgeted Hrs": updated_budgeted,
                    "Spent Hrs": existing_allocation["Spent Hrs"],
                    "Remaining Hrs": updated_budgeted - existing_allocation["Spent Hrs"],
                    "Cost/Hour": cost_per_hour,
                    "Budgeted Cost": budgeted_cost,
                    "Remaining Cost": remaining_cost,
                    "Section": existing_allocation["Section"],
                    "Category": existing_allocation["Category"]
                })


    
        # Metrics for Updated Allocations
        st.subheader("Updated Allocations for Selected Project")
        updated_df = pd.DataFrame(updated_rows)
        st.dataframe(updated_df)
    
        if not updated_df.empty:
            updated_budgeted_hours = updated_df["Budgeted Hrs"].sum()
            updated_spent_hours = updated_df["Spent Hrs"].sum()
            updated_budgeted_cost = updated_df["Budgeted Cost"].sum()
            updated_spent_cost = updated_df["Spent Hrs"].sum() * updated_df["Cost/Hour"].sum()
    
            st.metric("Updated Budgeted Hours", f"{updated_budgeted_hours} hrs")
            st.metric("Updated Spent Hours", f"{updated_spent_hours} hrs")
            st.metric("Updated Budgeted Cost", f"${updated_budgeted_cost:,.2f}")
            st.metric("Updated Spent Cost", f"${updated_spent_cost:,.2f}")
    
        # Save Updates
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
