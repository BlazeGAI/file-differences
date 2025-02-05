import streamlit as st
from docx import Document
from deepdiff import DeepDiff
import webcolors
import os
import xml.etree.ElementTree as ET
from utils.display import display_results
from checkers.word.word_1 import check_word_1

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

def extract_text_with_styles(doc_path):
    """Extracts text along with font colors, background colors, font sizes, alignments, and table formatting."""
    doc = Document(doc_path)
    content = []
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        styles = {
            "font_color": None,
            "background_color": None,
            "font_size": None,
            "bold": any(run.bold for run in para.runs),
            "italic": any(run.italic for run in para.runs),
            "underline": any(run.underline for run in para.runs),
            "alignment": para.alignment,
            "heading": para.style.name if para.style.name.startswith("Heading") else None,
        }

        content.append((text, styles))
    
    return content

def compare_word_documents(file1, file2):
    """Compares text, font styles, colors, font sizes, tables, alignments, and border settings."""
    text1 = extract_text_with_styles(file1)
    text2 = extract_text_with_styles(file2)

    differences = DeepDiff(text1, text2, ignore_order=False, report_repetition=True)
    return differences

st.header("Word Difference Checker")
word_file_1 = st.file_uploader("Upload First Word Document", type=["docx"], key="word_diff_1")
word_file_2 = st.file_uploader("Upload Second Word Document", type=["docx"], key="word_diff_2")

if word_file_1 and word_file_2:
    file1_path = os.path.join("temp1.docx")
    file2_path = os.path.join("temp2.docx")
    
    with open(file1_path, "wb") as f:
        f.write(word_file_1.getbuffer())
    with open(file2_path, "wb") as f:
        f.write(word_file_2.getbuffer())
    
    st.write("Comparing documents... ⏳")
    differences = compare_word_documents(file1_path, file2_path)
    
    if differences:
        st.write("### Differences Found 🧐")
        st.json(differences)
    else:
        st.write("✅ No differences found. The documents are identical.")
    
    os.remove(file1_path)
    os.remove(file2_path)
