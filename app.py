import streamlit as st
import docx
from io import BytesIO
from difflib import Differ
import pandas as pd

def compare_word_documents(master_file, student_file):
    try:
        # ... (rest of the document comparison logic)

        if not diff_data:  # Check if diff_data is empty
            return ["No differences found."] # Return a list with a message.

        return diff_data

    except docx.opc.exceptions.PackageNotFoundError:
        return ["Error: One or both files not found or invalid Word documents."]
    except Exception as e:
        return [f"Error processing documents: {e}"]


st.title("Word Document Comparison (Master vs. Student)")

# ... (file uploaders)

if uploaded_master_file and uploaded_student_file:
    diff_result = compare_word_documents(uploaded_master_file, uploaded_student_file)

    if isinstance(diff_result, list) and diff_result and "Error:" in diff_result[0]: # Check if diff_result is a list, is not empty, and contains "Error:"
        st.error(diff_result[0])
    elif isinstance(diff_result, list) and diff_result and "No differences found." in diff_result[0]:
        st.info(diff_result[0]) # Display the "no differences" message
    elif isinstance(diff_result, list) and diff_result: # Check if diff_result is a non-empty list of differences
        df = pd.DataFrame(diff_result)
        st.dataframe(df)
        # ... (download logic)
    else: # Handle unexpected return values
        st.error("An unexpected error occurred during comparison.")
