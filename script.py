        st.stop()

    engineer_names = engineers_data["Name"].dropna().tolist()

    # Load existing projects
    try:
        project_data = load_data(EXCEL_FILE)
    except FileNotFoundError:
        project_data = pd.DataFrame(columns=["Project Name", "Personnel", "Week", "Budgeted Hrs", "Spent Hrs"])

    if action == "Create New Project":
        st.subheader("Step 3: Create New Project")
        project_name = st.text_input("Project Name")
        project_id = st.text_input("Project ID")
        year = st.selectbox("Year", range(datetime.now().year - 5, datetime.now().year + 6))
        month = st.selectbox("Month", list(month_name)[1:])
        month_index = list(month_name).index(month)

        weeks = generate_weeks_for_month(year, month_index)

        st.subheader("Step 4: Assign Weekly Budgeted Hours")
        if "new_project_allocations" not in st.session_state:
            st.session_state.new_project_allocations = []

        for engineer in engineer_names:
            st.markdown(f"**{engineer}**")
            for week_label, _ in weeks:
                hours = st.number_input(f"{week_label} Hours", min_value=0, step=1, key=f"{engineer}_{week_label}")
                if hours > 0:
                    st.session_state.new_project_allocations.append(
                        {"Project Name": project_name, "Personnel": engineer, "Week": week_label, "Budgeted Hrs": hours}
                    )

        if st.button("Submit Project"):
            new_data = pd.DataFrame(st.session_state.new_project_allocations)
            project_data = pd.concat([project_data, new_data], ignore_index=True)
            save_to_excel(project_data, EXCEL_FILE)
            st.success(f"Project '{project_name}' created successfully!")
            st.session_state.new_project_allocations = []

    elif action == "Update Existing Project":
        st.subheader("Step 3: Update Existing Project")
        project_names = project_data["Project Name"].unique()
        selected_project = st.selectbox("Select Project", project_names)

        if selected_project:
            project_allocations = project_data[project_data["Project Name"] == selected_project]

            for i, row in project_allocations.iterrows():
                st.markdown(f"**{row['Personnel']} ({row['Week']})**")
                col1, col2 = st.columns(2)
                with col1:
                    budgeted_hours = st.number_input(
                        "Budgeted Hours", min_value=0, value=row["Budgeted Hrs"], step=1, key=f"budget_{i}"
                    )
                with col2:
                    spent_hours = st.number_input(
                        "Spent Hours", min_value=0, value=row.get("Spent Hrs", 0), step=1, key=f"spent_{i}"
                    )

                project_data.loc[i, "Budgeted Hrs"] = budgeted_hours
                project_data.loc[i, "Spent Hrs"] = spent_hours

            if st.button("Save Updates"):
                save_to_excel(project_data, EXCEL_FILE)
                st.success(f"Project '{selected_project}' updated successfully!")

    # Show Summary
    st.subheader("Summary of Projects")
    if not project_data.empty:
        st.dataframe(project_data)
        st.write(f"**Total Budgeted Hours:** {project_data['Budgeted Hrs'].sum()}")
        st.write(f"**Total Spent Hours:** {project_data['Spent Hrs'].sum()}")

if __name__ == "__main__":
    main()
