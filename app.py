import streamlit as st
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from deepdiff import DeepDiff
import webcolors
import os
import xml.etree.ElementTree as ET

st.title("Assignment Checker 📄")

# Section: Word Document Comparison
st.header("Word Document Comparison")
word_file_1 = st.file_uploader("Upload First Word Document", type=["docx"], key="word_1")
word_file_2 = st.file_uploader("Upload Second Word Document", type=["docx"], key="word_2")

def closest_color(hex_code):
    """Convert hex color codes to the closest known color name."""
    hex_code = f"#{hex_code}" if not hex_code.startswith("#") else hex_code
    try:
        return webcolors.hex_to_name(hex_code)
    except ValueError:
        min_diff = float("inf")
        closest_name = None
        for hex_value, name in webcolors.CSS3_HEX_TO_NAMES.items():
            r1, g1, b1 = webcolors.hex_to_rgb(hex_code)
            r2, g2, b2 = webcolors.hex_to_rgb(hex_value)
            diff = (r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2
            if diff < min_diff:
                min_diff = diff
                closest_name = name
        return closest_name

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

    # Extract tables
    tables = []
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                row_data.append(cell_text)
            table_data.append(row_data)
        tables.append(table_data)

    return content, tables

def compare_word_documents(doc1, doc2):
    """Compares two Word documents for text and formatting differences."""
    text1, tables1 = extract_text_with_styles(doc1)
    text2, tables2 = extract_text_with_styles(doc2)

    differences = {
        "Text & Formatting Differences": DeepDiff(text1, text2, ignore_order=False, report_repetition=True),
        "Table Differences": DeepDiff(tables1, tables2, ignore_order=False, report_repetition=True),
    }

    return differences

if word_file_1 and word_file_2:
    with open("temp1.docx", "wb") as f1, open("temp2.docx", "wb") as f2:
        f1.write(word_file_1.getbuffer())
        f2.write(word_file_2.getbuffer())

    doc1 = Document("temp1.docx")
    doc2 = Document("temp2.docx")

    st.subheader("Comparison Results")
    differences = compare_word_documents(doc1, doc2)

    if differences["Text & Formatting Differences"] or differences["Table Differences"]:
        st.json(differences)
    else:
        st.write("✅ No differences found. The documents are identical.")

    os.remove("temp1.docx")
    os.remove("temp2.docx")

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
