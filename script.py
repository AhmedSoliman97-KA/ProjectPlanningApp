import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
from st_aggrid import AgGrid, GridOptionsBuilder

# Dropbox Settings
ACCESS_TOKEN = "sl.CCcPGy4iU-0WaE7Q7tiRS_4KQVPf7b19Omn3JT4cHxug9_F8Yx5eE1f58zx295o6v6Be_HkWfFEu__8xuBT7tbGo8gKkbKLYvFv-LVCzkF7PbSWyhOs0Cbf0HiIYOJACMfMBALc6TzYeqVYIg8zfw2A"
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
            st.success(f"File {file_path} successfully created in Dropbox.")
        else:
            st.info(f"File {file_path} already exists in Dropbox.")
    except Exception as e:
        st.error(f"Error ensuring file exists in Dropbox: {e}")
        raise

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

