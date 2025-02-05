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


# ... (rest of the Streamlit code - same as before)
