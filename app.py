import streamlit as st
from docx import Document
from deepdiff import DeepDiff
import pandas as pd
import webcolors
import os

st.title("Word Document Checker 📄")

# Upload Master and Student Documents
st.header("Upload Documents for Comparison")
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
    """Extracts text, font styles, colors, font sizes, headings, and table structures."""
    content = []
    tables = []

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

    # Extract tables
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                cell_styles = {
                    "text": cell_text,
                    "background_color": None,
                    "border_bottom": "Unknown",  # Default to unknown
                }

                tc = cell._tc
                tc_pr = tc.find("w:tcPr", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                if tc_pr is not None:
                    shd = tc_pr.find("w:shd", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                    if shd is not None:
                        cell_styles["background_color"] = closest_color(shd.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill", ""))

                    borders = tc_pr.find("w:tcBorders", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                    if borders is not None:
                        bottom_border = borders.find("w:bottom", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                        cell_styles["border_bottom"] = "Present" if bottom_border is not None else "Missing"

                row_data.append(cell_styles)
            table_data.append(row_data)
        tables.append(table_data)

    return content, tables

def compare_word_documents(master_doc, student_doc):
    """Compares two Word documents for text, formatting, and table differences."""
    master_text, master_tables = extract_text_with_styles(master_doc)
    student_text, student_tables = extract_text_with_styles(student_doc)

    differences = []
    
    for i, (master_entry, student_entry) in enumerate(zip(master_text, student_text)):
        if master_entry["text"] != student_entry["text"]:
            differences.append(["Text Change", student_entry["text"], master_entry["text"]])

        if master_entry["bold"] != student_entry["bold"]:
            differences.append(["Bold Difference", "Bold" if student_entry["bold"] else "Not Bold", "Bold" if master_entry["bold"] else "Not Bold"])

        if master_entry["font_color"] != student_entry["font_color"]:
            differences.append(["Font Color", student_entry["font_color"], master_entry["font_color"]])

    for i, (master_table, student_table) in enumerate(zip(master_tables, student_tables)):
        for j, (master_row, student_row) in enumerate(zip(master_table, student_table)):
            for k, (master_cell, student_cell) in enumerate(zip(master_row, student_row)):
                if master_cell["background_color"] != student_cell["background_color"]:
                    differences.append(["Cell Background Color", student_cell["background_color"], master_cell["background_color"]])

                if master_cell["border_bottom"] != student_cell["border_bottom"]:
                    differences.append(["Bottom Border", student_cell["border_bottom"], master_cell["border_bottom"]])

    return pd.DataFrame(differences, columns=["Category", "Student Version", "Master Version"])

if word_file_master and word_file_student:
    with open("master.docx", "wb") as f1, open("student.docx", "wb") as f2:
        f1.write(word_file_master.getbuffer())
        f2.write(word_file_student.getbuffer())

    master_doc = Document("master.docx")
    student_doc = Document("student.docx")

    st.subheader("Comparison Results")
    differences_df = compare_word_documents(master_doc, student_doc)

    if not differences_df.empty:
        st.write("### Differences Found:")
        st.dataframe(differences_df)
    else:
        st.write("✅ No differences found. The student document matches the master.")

    os.remove("master.docx")
    os.remove("student.docx")
