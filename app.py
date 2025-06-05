from logic import generate_excel_from_csv
import streamlit as st

st.title("Beach Assignment Generator")

uploaded_file = st.file_uploader("Upload the Google form data/CSV file", type=["csv"])

if uploaded_file:
    with st.spinner("Generating assignments..."):
        excel_file = generate_excel_from_csv(uploaded_file)
    
    st.success("Assignments generated!")

    st.download_button(
        label="📥 Download Excel File",
        data=excel_file,
        file_name="beach_assignments.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )