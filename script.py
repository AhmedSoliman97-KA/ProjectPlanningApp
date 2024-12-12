import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox
from st_aggrid import AgGrid, GridOptionsBuilder

# Dropbox Settings
ACCESS_TOKEN = "ssl.u.AFbHvxxiLExZLhPWHviU_aGF1LCc-DKzbQP9g34VD-AH2Y_eTSu60cOipCQuBd4bDi-qr9WCEOyFLkkT_KZzyHM0rZFzDNSct3tpYb-aDQNKjBVWXL-p2omLSt_FtZaH5qwkxDeEdDLuUSUQczAgxas541XWBSDc3EJATtMPj1WhoJqLhSh1uk4KUzkmjyAi1z_LwiQF2gB9zkGXZP7mYwctt7kYZw7DJnOFmXykWF8Vn0chQN9rGxsLyyP4S7Q-qq3Ay0Om7xY6a40keN7Ssut0sAEiepCtT7wj9xapJtW4ubwlqVY9R-7DR-84IaWFTPwGBN0BeKXK6UMn2USLWXpiydgwe4KzF7wU6wvucdg3cxF4iv5IQXDKxeQxYrgBcFbvR3t8RMOIqEHsoPjnQC2tkVhzch0nw9tXbC8sffzNjKI41iGuXi3Cit7vGNkLytu9M6ZnK46jziBtyeCTLq3WBpFugbyehd_CrL9CsVFTntUSQ2GR7voCY6sOcuxV2VOqVjydC91waOrIV-yxu72yxiB7IOjiXGJIp0qUUQ1qsM1kzRebC3dgjYeCzAOrHyDEBFzPiiKtaKAKFRz9EEnm8yApN3qP9s-tBDCTI3lqkvYi8Nw4IkMJqhpsQ6qYXCij5HM4a7x13DrQbaHASxaJjVzuoLkSzSMgBgXI1LmEXsDdNPTtfb3Co9wm5eqxxhdGKAnBqg1Yg4H5vVpDZcs4XBQVmuDt4rN34bOzqpxC36llQ_maYONUq9oiD0VfQTBPugg1i1N3EHh0iSHMEaKYOgRRtXx8jFmhCOou9PU2TSWS6fvYA4KIs0O4a2uvVa_m_V4sMwKOmrhAueCmp-9On3ZYuXM5G1RbSzqDsRLvMXYx1p8LmJT9KfgcoFta9cZr-7oRvBOVlUYLOOL9Uc6m82TIfWfkjlv4KN0-0gyRu_wWEbtbCVdrIfLeDoba-Qqf2qZsIbu4rLkTT7TrEvznTUfrSroPZBHZhGOqc05tdiC3aHQay5wAbXiiwieD3ZQALK3zdvQ_c6VEQMZdh-8JMm2UEAn60pBgFAGu44cY3isNhU7-Amo1hadp7KL1xAm83aV2AvCpDyTmxBS6q-2tVdqOlp8F3ir-VvpEKt4yr7iceyw4EwumpsUink4ihRbuw9dQMxUalQi2VJsyArYdImzwhba5SET3U8qvX5qrTbEz3cDtCnA73py1mw10KZiGBoVwW_MnetvtweZERbIB1_7qO53Je74mz0ZBFFYtuLlz4-q4c_QyDpLbkD9JRarMfnMjuiJ_zDRfPcgJbpkodn2DUEALlt4U-yYMP5Us2yqY5EKqEEWn0bBSlvs1-w8sx6kjZhEVUxLo5q9-H5ZZ-KenA9bPavYIj41LBRAWCleb-2LwHlUqlqcGAgVPNJspmSSgnpPL79PNz1PsdHIX-iiWEZrF4_HXVW-JilZnCk5_uapgUBekMxGPZzWNfhE"

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

