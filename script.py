import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# OneDrive file download link (replace with your link)
ONEDRIVE_LINK = "https://khatibandalami-my.sharepoint.com/:f:/g/personal/ahmedsayed_soliman_khatibalami_com/EpxIyzwnrYdNqoneD-qC9ioBRyd4FBcvlpL9HrZhPfI2vg?e=6scK2w"

# Load data from OneDrive
def load_excel_from_onedrive(link):
    try:
        response = requests.get(link)
        if response.status_code == 200:
            return pd.read_excel(BytesIO(response.content))
        else:
            st.error(f"Failed to download file: {response.status_code}")
            return pd.DataFrame(columns=["Project ID", "Project Name", "Personnel", "Week", "Budgeted Hrs", "Spent Hrs"])
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame(columns=["Project ID", "Project Name", "Personnel", "Week", "Budgeted Hrs", "Spent Hrs"])

# Save data to OneDrive
def save_excel_to_onedrive(data, local_file_path):
    try:
        with pd.ExcelWriter(local_file_path, engine="openpyxl") as writer:
            data.to_excel(writer, index=False)
        st.success("Data saved successfully! Upload the updated file to OneDrive manually.")
    except Exception as e:
        st.error(f"Failed to save data: {e}")

# Main Streamlit App
def main():
    st.title("Water & Environment Project Planning")

    # Load existing data
    st.subheader("Existing Projects")
    project_data = load_excel_from_onedrive(ONEDRIVE_LINK)
    st.dataframe(project_data)

    # Add new project data
    st.subheader("Add New Project")
    project_id = st.text_input("Project ID")
    project_name = st.text_input("Project Name")
    engineer = st.text_input("Engineer Name")
    week = st.text_input("Week")
    budgeted_hours = st.number_input("Budgeted Hours", min_value=0)

    if st.button("Submit"):
        new_row = {
            "Project ID": project_id,
            "Project Name": project_name,
            "Personnel": engineer,
            "Week": week,
            "Budgeted Hrs": budgeted_hours,
            "Spent Hrs": 0
        }
        project_data = pd.concat([project_data, pd.DataFrame([new_row])], ignore_index=True)
        save_excel_to_onedrive(project_data, "updated_projects_data.xlsx")

    # Update existing data
    st.subheader("Update Existing Project")
    if not project_data.empty:
        project_names = project_data["Project Name"].unique().tolist()
        selected_project = st.selectbox("Select a Project to Update", project_names)

        if selected_project:
            filtered_data = project_data[project_data["Project Name"] == selected_project]
            for index, row in filtered_data.iterrows():
                col1, col2 = st.columns(2)
                with col1:
                    budgeted_hours = st.number_input(
                        f"Budgeted Hours for {row['Personnel']} ({row['Week']})",
                        min_value=0,
                        value=row["Budgeted Hrs"],
                        step=1,
                        key=f"budget_{index}"
                    )
                with col2:
                    spent_hours = st.number_input(
                        f"Spent Hours for {row['Personnel']} ({row['Week']})",
                        min_value=0,
                        value=row["Spent Hrs"],
                        step=1,
                        key=f"spent_{index}"
                    )
                project_data.loc[index, "Budgeted Hrs"] = budgeted_hours
                project_data.loc[index, "Spent Hrs"] = spent_hours

            if st.button("Save Updates"):
                save_excel_to_onedrive(project_data, "updated_projects_data.xlsx")
                st.success(f"Project '{selected_project}' updated successfully!")

    # Summary of projects
    st.subheader("Summary of All Projects")
    if not project_data.empty:
        st.dataframe(project_data)
        total_budgeted_hours = project_data["Budgeted Hrs"].sum()
        total_spent_hours = project_data["Spent Hrs"].sum()
        st.write(f"**Total Budgeted Hours:** {total_budgeted_hours}")
        st.write(f"**Total Spent Hours:** {total_spent_hours}")

if __name__ == "__main__":
    main()
