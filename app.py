import streamlit as st
import docx
from io import BytesIO
from difflib import Differ

def compare_word_documents(file1, file2):
    try:
        doc1 = docx.Document(file1)
        doc2 = docx.Document(file2)

        text1 = "\n".join([p.text for p in doc1.paragraphs])
        text2 = "\n".join([p.text for p in doc2.paragraphs])

        differ = Differ()
        diff = list(differ.compare(text1.splitlines(), text2.splitlines()))

        return diff

    except docx.opc.exceptions.PackageNotFoundError:
        return ["Error: One or both files not found or invalid Word documents."]
    except Exception as e:  # Catching a broader exception for other potential docx errors
        return [f"Error processing documents: {e}"]


st.title("Word Document Comparison")

uploaded_file1 = st.file_uploader("Upload Document 1", type=["docx"])
uploaded_file2 = st.file_uploader("Upload Document 2", type=["docx"])

if uploaded_file1 and uploaded_file2:
    diff_result = compare_word_documents(uploaded_file1, uploaded_file2)

    if "Error:" in diff_result[0]: # Check if there was an error
        st.error(diff_result[0]) # Display the error to the user
    else:
        st.subheader("Comparison Results")

        for line in diff_result:
            if line.startswith("+ "):
                st.markdown(f"<span style='color:green;'>{line}</span>", unsafe_allow_html=True)  # Added
            elif line.startswith("- "):
                st.markdown(f"<span style='color:red;'>{line}</span>", unsafe_allow_html=True)  # Added
            elif line.startswith("? "): # Added for highlighting changes
                st.markdown(f"<span style='color:orange;'>{line}</span>", unsafe_allow_html=True)  # Added
            else:
                st.write(line)


# Optional: Add a download button for the diff results
if uploaded_file1 and uploaded_file2 and "Error:" not in diff_result[0]:
    diff_text = "\n".join(diff_result)
    b = BytesIO()
    b.write(diff_text.encode())  # Encode to bytes
    st.download_button(
        label="Download Diff",
        data=b,
        file_name="diff.txt",
        mime="text/plain",
    )
