import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import month_name
import dropbox

# Dropbox Settings
ACCESS_TOKEN = "sl.u.AFYUwp18eTgoFNK7TUsLEW18-UpWE80v1N9bqMD9uWJrv8hHZHQXAFL8oRen_WCym4SuhXCAkfaxmPFyPiuetz5pmbGot-4Q_niLdbuOwcWTUiolWsqLMwsO74vvUNVy0bI3yrALytVBriQ9wpxB5kNrj9L69qPZn1iTGetTDxJD3TTmW3naoDFg4RiiBRXKVyQw5-T6xSix8TfXv36S1opHOZLiCD3O_pIrzNSlrWI6oiARKTLUOM0GlQWUcsfGjltD8vhJu2KZPQ9bbuXj4r8UFqiLr5-Rv3MDUjLX8DFkcBP7Uy8LwXKqDcF2z0rML9pEe-yCdRQ6WvpApenPoszFOEnM9t8mJM4FzMRjGZm6J27TAzE2sM55-Hxu3NRBayy1cG3ouAw9vt_ucKF2uQ3nJIKHdrD8c1Nwl4xTpVYbhxmkVp1lFdOAh5qjXzKxOfxb0O-quOgZQMoYxJ_FDQLXEZQKI6HMmNHC4pjPNGvSvFTk6Z1n7Ltedyouaw9MkC3H4apZG2mhH2P-M6StJ0Uwfrmpn53ExceC7wGjkBcRWeuTYfYcbO4NZPfjfZf8Of4gey3rldtgW-FXRB8yj_MkFNA5yngF0n7eAZMFQboVkPZlWLrMV_XTP9EYQmSNNENsiL1hb311gRiI1lA6c2BPYn-RqJfe4Gt0MBfTACc-QonhzafTJCsmSdL5KWw0kk51K2BIuZSHMSfKPWwxae08N8aGQCHb1_CSLlId0tWaogkjoTQjd4aGS3OuM7AI6KcMCacDzdaBUBzJRiG2LjYi5zNkm3e10e63SKGmDFOKVaQ_pHMXJov3pSDfqs0L4JwO1OonFxBf8ZbyxtQJZwRD6XrTmGw0b9yMrhd8qKhlfz2aOGoxqYPx-0GMLod2yam279Fo5-qmzsWjNPxUqH-n8_63UvLI1qkfVo1VPRdYG_HXTF0C-p71oiMS9-iV8wWd1w0Ya1c2P7wGm2GuWpLogbqW0n4W_hgerMU5yyhSaNlQldYxsSWQrprIdur70DCoc1XF0BI-lorjjaYLqA2W0mzFZKp1Wi3UCXJmqvVFU7t3klA1d28S5CUbu57WzUiWg1cYxPtzumTsmxK3nkkhnUd0E1k-0sM0Z-QVprwJwGadwcc6bBieEeNylJN7aASIwSZSMXRjrwOXgMOS7YQ1s5ro2ZskXXyLritvjBJf6wuAWvdhgIy4i5QZimN56HzaxwOlFNk_fbe24NJLENUaUxP4rkGkEwu8rQ-vcthmGANL36HZ_NoEL46uSiv1kPTZuCN9_XRPTh2HKZ4R7_LDxkc__zL2G7CwDaMTeHSxAuX4o0j2hxIMEwmZlLqzyKPRiaolAuzBdXDsgpO7l9pMT1U_DFixfbFbllhswH2senCfI4T9A2FSj-6NF1wAvQXPxwklxIrER_1qZlNtgadScMJy3sZQheLS9fosdXNMW1BpVyeR_2RTvMLtcY17q9c"
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

def upload_to_dropbox(dfs, dropbox_path):
    """Upload multiple DataFrames to Dropbox as an Excel file."""
    try:
        dbx = dropbox.Dropbox(ACCESS_TOKEN)
        with pd.ExcelWriter("temp.xlsx", engine="openpyxl") as writer:
            for sheet_name, data in dfs.items():
                data.to_excel(writer, index=False, sheet_name=sheet_name)
        with open("temp.xlsx", "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
        st.success(f"Data successfully uploaded to {dropbox_path}.")
    except Exception as e:
        st.error(f"Error uploading data: {e}")
        raise

def ensure_dropbox_projects_file_exists(file_path):
    """Ensure the projects file exists in Dropbox, create or fix if not."""
    existing_file = download_from_dropbox(file_path)
    writer_needed = False

    if existing_file is None:
        st.warning(f"{file_path} not found in Dropbox. Creating a new file...")
        writer_needed = True
        projects_data = pd.DataFrame(columns=[
            "Project ID", "Project Name", "Personnel", "Week", "Year", "Month",
            "Budgeted Hrs", "Spent Hrs", "Remaining Hrs", "Cost/Hour", "Budgeted Cost",
            "Remaining Cost", "Section", "Category"
        ])
        approved_budgets_data = pd.DataFrame(columns=["Project ID", "Project Name", "Approved Total Budget"])
    else:
        # Load existing sheets
        sheets = existing_file.sheet_names

        # Validate "Projects" sheet
        if "Projects" not in sheets:
            st.warning("'Projects' sheet missing. Reinitializing...")
            writer_needed = True
            projects_data = pd.DataFrame(columns=[
                "Project ID", "Project Name", "Personnel", "Week", "Year", "Month",
                "Budgeted Hrs", "Spent Hrs", "Remaining Hrs", "Cost/Hour", "Budgeted Cost",
                "Remaining Cost", "Section", "Category"
            ])
        else:
            projects_data = pd.read_excel(existing_file, sheet_name="Projects")

        # Validate "Approved Budgets" sheet
        if "Approved Budgets" not in sheets:
            st.warning("'Approved Budgets' sheet missing. Reinitializing...")
            writer_needed = True
            approved_budgets_data = pd.DataFrame(columns=["Project ID", "Project Name", "Approved Total Budget"])
        else:
            approved_budgets_data = pd.read_excel(existing_file, sheet_name="Approved Budgets")

    # Save back if changes were made
    if writer_needed:
        with pd.ExcelWriter("temp.xlsx", engine="openpyxl") as writer:
            projects_data.to_excel(writer, index=False, sheet_name="Projects")
            approved_budgets_data.to_excel(writer, index=False, sheet_name="Approved Budgets")
        with open("temp.xlsx", "rb") as f:
            dbx = dropbox.Dropbox(ACCESS_TOKEN)
            dbx.files_upload(f.read(), file_path, mode=dropbox.files.WriteMode("overwrite"))

    return projects_data, approved_budgets_data

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
    projects_data, approved_budgets_data = ensure_dropbox_projects_file_exists(PROJECTS_FILE_PATH)

    # Load Human Resources File from Dropbox
    hr_excel = download_from_dropbox(HR_FILE_PATH)
    if hr_excel is None:
        st.error(f"Human Resources file not found in Dropbox at {HR_FILE_PATH}.")
        st.stop()

    hr_sections = hr_excel.sheet_names

    # Action Selection
    st.sidebar.subheader("Actions")
    action = st.sidebar.radio("Choose an Action", ["Create New Project", "Update Existing Project"])

    # Create New Project
    if action == "Create New Project":
        st.subheader("Create a New Project")

        project_id = st.text_input("Project ID", help="Enter a unique ID for the project.")
        project_name = st.text_input("Project Name", help="Enter the name of the project.")
        approved_budget = st.number_input("Approved Total Budget (in $)", min_value=0, step=1)

        selected_year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6), index=5)
        selected_month = st.selectbox("Month", list(month_name)[1:], index=datetime.now().month - 1)

        st.subheader("Select Engineers for Allocation")
        selected_section = st.selectbox("Choose a Section", hr_sections)
        engineers_data = hr_excel.parse(sheet_name=selected_section)
        engineers = engineers_data["Name"].dropna().tolist()
        selected_engineers = st.multiselect("Choose Engineers", options=engineers)

        allocations = []
        total_allocated_budget = 0

        if selected_engineers:
            st.subheader("Allocate Weekly Hours")
            weeks = generate_weeks(selected_year, list(month_name).index(selected_month))

            for engineer in selected_engineers:
                engineer_details = engineers_data[engineers_data["Name"] == engineer].iloc[0]
                section = engineer_details.get("Section", "Unknown")
                category = engineer_details.get("Category", "N/A")
                cost_per_hour = pd.to_numeric(engineer_details.get("Cost/Hour", 0), errors="coerce")
                cost_per_hour = cost_per_hour if not pd.isna(cost_per_hour) else 0

                st.markdown(f"### Engineer: {engineer}")
                for week_label, _ in weeks:
                    budgeted_hours = st.number_input(
                        f"Budgeted Hours ({week_label}) for {engineer}",
                        min_value=0, step=1, key=f"{engineer}_{week_label}"
                    )
                    if budgeted_hours > 0:
                        budgeted_cost = budgeted_hours * cost_per_hour
                        total_allocated_budget += budgeted_cost
                        allocations.append({
                            "Project ID": project_id,
                            "Project Name": project_name,
                            "Personnel": engineer,
                            "Week": week_label,
                            "Year": selected_year,
                            "Month": selected_month,
                            "Budgeted Hrs": budgeted_hours,
                            "Cost/Hour": cost_per_hour,
                            "Budgeted Cost": budgeted_cost,
                            "Section": section,
                            "Category": category
                        })

        if allocations:
            st.subheader("Summary of Allocations")
            allocation_df = pd.DataFrame(allocations)
            st.dataframe(allocation_df)

            st.metric("Total Allocated Budget (in $)", f"${total_allocated_budget:,.2f}")
            st.metric("Approved Total Budget (in $)", f"${approved_budget:,.2f}")
            st.metric("Difference (Remaining/Over-Allocated)", f"${approved_budget - total_allocated_budget:,.2f}")

        if st.button("Submit Project"):
            new_data = pd.DataFrame(allocations)
            approved_budgets_data = approved_budgets_data.append(
                {"Project ID": project_id, "Project Name": project_name, "Approved Total Budget": approved_budget},
                ignore_index=True
            )
            upload_to_dropbox(
                {"Projects": pd.concat([projects_data, new_data], ignore_index=True),
                 "Approved Budgets": approved_budgets_data},
                PROJECTS_FILE_PATH
            )
            st.success("Project submitted successfully!")

    # Update Existing Project
    if action == "Update Existing Project":
        st.subheader("Update an Existing Project")
        project_id = st.selectbox("Select Project ID", approved_budgets_data["Project ID"].unique())
        selected_budget = approved_budgets_data[approved_budgets_data["Project ID"] == project_id]["Approved Total Budget"].values[0]

        st.metric("Approved Total Budget (in $)", f"${selected_budget:,.2f}")

        project_details = projects_data[projects_data["Project ID"] == project_id]
        st.subheader("Current Allocations for Selected Project")
        st.dataframe(project_details)

        updated_rows = []
        for _, row in project_details.iterrows():
            updated_budgeted = st.number_input(
                f"Budgeted Hours ({row['Week']})",
                min_value=0,
                value=int(row["Budgeted Hrs"]) if not pd.isna(row["Budgeted Hrs"]) else 0,
                step=1,
                key=f"update_budgeted_{row['Personnel']}_{row['Week']}"
            )
            updated_spent = st.number_input(
                f"Spent Hours ({row['Week']})",
                min_value=0,
                value=int(row["Spent Hrs"]) if not pd.isna(row["Spent Hrs"]) else 0,
                step=1,
                key=f"update_spent_{row['Personnel']}_{row['Week']}"
            )
            updated_rows.append({
                **row.to_dict(),
                "Budgeted Hrs": updated_budgeted,
                "Spent Hrs": updated_spent,
                "Remaining Hrs": updated_budgeted - updated_spent
            })

        if st.button("Save Updates"):
            updated_rows_df = pd.DataFrame(updated_rows)
            updated_data = pd.concat(
                [projects_data[projects_data["Project ID"] != project_id], updated_rows_df],
                ignore_index=True
            )
            upload_to_dropbox({"Projects": updated_data, "Approved Budgets": approved_budgets_data}, PROJECTS_FILE_PATH)
            st.success(f"Updates to project '{project_id}' saved successfully!")

if __name__ == "__main__":
    main()

