import streamlit as st
import docx
from io import BytesIO
from difflib import Differ
import pandas as pd

def compare_word_documents(master_file_bytes, student_file_bytes):
    try:
        master_doc = docx.Document(BytesIO(master_file_bytes))
        student_doc = docx.Document(BytesIO(student_file_bytes))

        master_text = "\n".join([p.text for p in master_doc.paragraphs])
        student_text = "\n".join([p.text for p in student_doc.paragraphs])

        differ = Differ()
        diff = list(differ.compare(master_text.splitlines(), student_text.splitlines()))

        diff_data = []
        for line in diff:
            status = ""
            text = line[2:].strip()  # Remove "+", "-", or "?", and strip whitespace
            if line.startswith("+ "):
                status = "Added in Student"
            elif line.startswith("- "):
                status = "Removed from Student"
            elif line.startswith("? "):
                status = "Changed in Student"
            else:
                continue  # Skip "Same" lines to keep the table concise

            diff_data.append({"Status": status, "Text": text})

        if not diff_data:
            return ["No differences found."]

        return diff_data

    except docx.opc.exceptions.PackageNotFoundError:
        return ["Error: One or both files not found or invalid Word documents."]
    except Exception as e:
        return [f"Error processing documents: {e}"]


st.title("Word Document Comparison (Master vs. Student)")

uploaded_master_file = st.file_uploader("Upload Master Document", type=["docx"])
uploaded_student_file = st.file_uploader("Upload Student Submission", type=["docx"])

if uploaded_master_file and uploaded_student_file:
    master_file_bytes = uploaded_master_file.getvalue()
    student_file_bytes = uploaded_student_file.getvalue()

    diff_result = compare_word_documents(master_file_bytes, student_file_bytes)

    if isinstance(diff_result, list) and diff_result and "Error:" in diff_result[0]:
        st.error(diff_result[0])
    elif isinstance(diff_result, list) and diff_result and "No differences found." in diff_result[0]:
        st.info(diff_result[0])
    elif isinstance(diff_result, list) and diff_result:
        df = pd.DataFrame(diff_result)
        st.dataframe(df)

        csv_data = df.to_csv(index=False).encode('utf-8')  # Create CSV data here
        st.download_button(
            label="Download Diff (CSV)",
            data=csv_data,
            file_name="diff.csv",
            mime="text/csv",
        )
    else:
        st.error("An unexpected error occurred during comparison.")
