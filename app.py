import streamlit as st
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from deepdiff import DeepDiff
import webcolors
import os

st.title("Assignment Checker 📄")

# Section: Word Document Comparison
st.header("Word Document Comparison")
word_file_master = st.file_uploader("Upload Master Document (Correct Version)", type=["docx"], key="master_doc")
word_file_student = st.file_uploader("Upload Student Document", type=["docx"], key="student_doc")

def closest_color(hex_code):
    """Convert hex color codes to the closest known color name."""
    hex_code = f"#{hex_code}" if not hex_code.startswith("#") else hex_code
    try:
        return webcolors.hex_to_name(hex_code)
    except ValueError:
        return hex_code  # Return hex code if no name found

def extract_text_with_styles(doc):
    """Extracts text, font styles, colors, font sizes, and heading styles."""
    content = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        styles = {
            "text": text,
            "font_color": None,
            "background_color": None,
            "font_size": None,
            "bold": any(run.bold for run in para.runs),
            "italic": any(run.italic for run in para.runs),
            "underline": any(run.underline for run in para.runs),
            "heading": para.style.name if para.style.name.startswith("Heading") else None,
        }

        for run in para.runs:
            rpr = run._element.find("w:rPr", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
            if rpr is not None:
                sz = rpr.find("w:sz", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                if sz is not None:
                    styles["font_size"] = int(sz.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"]) / 2

                color = rpr.find("w:color", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                if color is not None:
                    styles["font_color"] = closest_color(color.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"])

                highlight = rpr.find("w:highlight", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                if highlight is not None:
                    styles["background_color"] = closest_color(highlight.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"])

        content.append(styles)

    return content

def compare_word_documents(master_doc, student_doc):
    """Compares two Word documents for text and formatting differences."""
    master_text = extract_text_with_styles(master_doc)
    student_text = extract_text_with_styles(student_doc)

    differences = DeepDiff(master_text, student_text, ignore_order=False, report_repetition=True)

    results = []
    for key, diff in differences.items():
        if key == "values_changed":
            for path, change in diff.items():
                index = int(path.split("[")[1].split("]")[0])  # Extract index
                master_entry = master_text[index]
                student_entry = student_text[index]

                results.append({
                    "Category": "Text Change",
                    "Student Version": student_entry["text"],
                    "Master Version": master_entry["text"],
                })

        elif key == "iterable_item_removed":
            for path, removed in diff.items():
                results.append({
                    "Category": "Removed Text",
                    "Student Version": "(Missing)",
                    "Master Version": removed["text"],
                })

        elif key == "iterable_item_added":
            for path, added in diff.items():
                results.append({
                    "Category": "Extra Text",
                    "Student Version": added["text"],
                    "Master Version": "(Not in Master)",
                })

    return results

if word_file_master and word_file_student:
    with open("master.docx", "wb") as f1, open("student.docx", "wb") as f2:
        f1.write(word_file_master.getbuffer())
        f2.write(word_file_student.getbuffer())

    master_doc = Document("master.docx")
    student_doc = Document("student.docx")

    st.subheader("Comparison Results")
    differences = compare_word_documents(master_doc, student_doc)

    if differences:
        st.write("### Differences Found:")
        st.table(differences)
    else:
        st.write("✅ No differences found. The student document matches the master.")

    os.remove("master.docx")
    os.remove("student.docx")

# Section: Excel Assignments
st.header("Excel Assignments")
excel_file = st.file_uploader("Upload Excel Assignment", type=["xlsx"], key="excel")

if excel_file:
    try:
        workbook = load_workbook(excel_file)
        st.write("✅ Excel file uploaded successfully.")
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")

# Section: PowerPoint Assignments
st.header("PowerPoint Assignments")
ppt_file = st.file_uploader("Upload PowerPoint Assignment", type=["pptx"], key="ppt")

if ppt_file:
    try:
        prs = Presentation(ppt_file)
        st.write("✅ PowerPoint file uploaded successfully.")
    except Exception as e:
        st.error(f"Error loading PowerPoint file: {str(e)}")
