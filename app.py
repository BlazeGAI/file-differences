import streamlit as st
import docx
from io import BytesIO
from difflib import Differ

def compare_word_documents(master_file, student_file):
    try:
        master_doc = docx.Document(master_file)
        student_doc = docx.Document(student_file)

        master_text = "\n".join([p.text for p in master_doc.paragraphs])
        student_text = "\n".join([p.text for p in student_doc.paragraphs])

        differ = Differ()
        diff = list(differ.compare(master_text.splitlines(), student_text.splitlines()))

        return diff

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

        for line in diff_result:
            if line.startswith("+ "):
                st.markdown(f"<span style='color:green;'>{line.lstrip('+ ')} (Added in Student)</span>", unsafe_allow_html=True)
            elif line.startswith("- "):
                st.markdown(f"<span style='color:red;'>{line.lstrip('- ')} (Removed from Student)</span>", unsafe_allow_html=True)
            elif line.startswith("? "):  # For highlighting changes, less common in this context
                st.markdown(f"<span style='color:orange;'>{line}</span>", unsafe_allow_html=True)
            else:
                st.write(line)


# Optional: Download Diff
if uploaded_master_file and uploaded_student_file and "Error:" not in diff_result[0]:
    diff_text = "\n".join(diff_result)
    b = BytesIO()
    b.write(diff_text.encode())
    st.download_button(
        label="Download Diff",
        data=b,
        file_name="diff.txt",
        mime="text/plain",
    )
