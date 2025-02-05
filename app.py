import streamlit as st
import docx
from io import BytesIO
from difflib import Differ
import pandas as pd
from itertools import zip_longest
from docx.api import Paragraph, Run # Import correctly


def compare_word_documents(master_file_bytes, student_file_bytes):
    try:
        master_doc = docx.Document(BytesIO(master_file_bytes))
        student_doc = docx.Document(BytesIO(student_file_bytes))

        diff_data = []

        for i, (mp, sp) in enumerate(zip_longest(master_doc.paragraphs, student_doc.paragraphs, fillvalue=Paragraph(""))):
            for j, (mr, sr) in enumerate(zip_longest(mp.runs, sp.runs, fillvalue=Run(mp))): # Use correct Run
                if mr.text != sr.text or get_run_format(mr) != get_run_format(sr):
                    diff_data.append({
                        "Paragraph": i + 1,
                        "Run": j + 1,
                        "Master Text": mr.text,
                        "Student Text": sr.text,
                        "Master Format": get_run_format(mr),
                        "Student Format": get_run_format(sr)
                    })

        if not diff_data:
            return ["No differences found."]

        return diff_data

    except docx.opc.exceptions.PackageNotFoundError:
        return ["Error: One or both files not found or invalid Word documents."]
    except Exception as e:
        return [f"Error processing documents: {e}"]

def get_run_format(run):
    format_info = {}
    format_info["bold"] = run.bold
    format_info["italic"] = run.italic
    format_info["underline"] = run.underline
    format_info["font_size"] = run.font.size and run.font.size.pt
    format_info["color"] = run.font.color.rgb if run.font.color.rgb else None
    return format_info


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

        csv_data = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download Diff (CSV)",
            data=csv_data,
            file_name="diff.csv",
            mime="text/csv",
        )
    else:
        st.error("An unexpected error occurred during comparison.")

