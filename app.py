from logic import generate_excel_from_csv
import streamlit as st

st.title("Beach Assignment Generator")

uploaded_file = st.file_uploader("Upload the Google form data/CSV file", type=["csv"])

use_previous_schedule = st.checkbox("Do you want to upload a previous schedule to avoid repeat assignments?")
previous_schedule_file = None
if use_previous_schedule:
    previous_schedule_file = st.file_uploader('📅 Upload previous schedule (Excel)", type=["xlsx"]')

if uploaded_file:
    with st.spinner("Generating assignments..."):
        excel_file = generate_excel_from_csv(uploaded_file, previous_schedule_file)
    
    st.success("Assignments generated!")

    st.download_button(
        label="📥 Download Excel File",
        data=excel_file,
        file_name="beach_assignments.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )