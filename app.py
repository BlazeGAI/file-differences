import streamlit as st
import docx
from io import BytesIO
from difflib import Differ
import pandas as pd

def compare_word_documents(master_file, student_file):
    try:
        master_doc = docx.Document(master_file)
        student_doc = docx.Document(student_file)

        master_text = "\n".join([p.text for p in master_doc.paragraphs])
        student_text = "\n".join([p.text for p in student_doc.paragraphs])

        differ = Differ()
        diff = list(differ.compare(master_text.splitlines(), student_text.splitlines()))

        diff_data = []
        for line in diff:
            status = ""
            text = line[2:]  # Remove the "+", "-", or "?"
            if line.startswith("+ "):
                status = "Added in Student"
            elif line.startswith("- "):
                status = "Removed from Student"
            elif line.startswith("? "):
                status = "Changed in Student" # Less common, but included
            else:
                status = "Same"
            if status != "Same": # We only care about the differences.
                diff_data.append({"Status": status, "Text": text})
        return diff_data

    except docx.opc.exceptions.PackageNotFoundError:
        return ["Error: One or both files not found or invalid Word documents."]
    except Exception as e:
        return [f"Error processing documents: {e}"]


st.title("Word Document Comparison (Master vs. Student)")

uploaded_master_file = st.file_uploader("Upload Master Document", type=["docx"])
uploaded_student_file = st.file_uploader("Upload Student Submission", type=["docx"])

if uploaded_master_file and uploaded_student_file:
    diff_result = compare_word_documents(uploaded_master_file, uploaded_student_file)

    if "Error:" in diff_result[0]:
        st.error(diff_result[0])
    else:
        st.subheader("Comparison Results (Master vs. Student)")

        if diff_result:  # Check if there are any differences to display
            df = pd.DataFrame(diff_result)
            st.dataframe(df)  # Display the DataFrame as a table
        else:
            st.info("No differences found between the documents.") # Inform if no differences.


# Optional: Download Diff (now downloads a CSV)
if uploaded_master_file and uploaded_student_file and "Error:" not in diff_result[0]:
    if diff_result:
        csv_data = pd.DataFrame(diff_result).to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download Diff (CSV)",
            data=csv_data,
            file_name="diff.csv",
            mime="text/csv",
        )
