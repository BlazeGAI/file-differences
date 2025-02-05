import streamlit as st
from docx import Document
import pandas as pd
import os
from collections import defaultdict

st.title("Word Document Format Checker 📄")

# Upload Documents
st.header("Upload Documents for Comparison")
word_file_student = st.file_uploader("Upload Student Document", type=["docx"], key="student_doc")

def check_table_above_date(doc):
    """Check if there's a 2x1 table above the date paragraph (Criteria 2)"""
    found_date = False
    for i in range(len(doc.paragraphs) - 1, -1, -1):
        if "date" in doc.paragraphs[i].text.lower():
            found_date = True
            # Check if there's a table before this paragraph
            tables_before = [table for table in doc.tables 
                           if doc.element.body.index(table._element) < 
                           doc.element.body.index(doc.paragraphs[i]._element)]
            if tables_before and tables_before[-1].rows[0].cells[0].text.strip():
                last_table = tables_before[-1]
                return (len(last_table.rows) == 1 and len(last_table.rows[0].cells) == 2,
                        "Found 2x1 table above date" if len(last_table.rows) == 1 and 
                        len(last_table.rows[0].cells) == 2 else "Table dimensions incorrect")
    return (False, "Could not find table above date paragraph")

def check_address_in_table(doc):
    """Check if the address is correctly entered in the table (Criteria 3)"""
    for table in doc.tables:
        if len(table.rows) == 1 and len(table.rows[0].cells) == 2:
            cell_1_text = table.rows[0].cells[0].text.strip()
            return ("Classic Cars Club" in cell_1_text and 
                    "PO Box 6987" in table.rows[0].cells[1].text.strip(),
                    "Address found in correct format" if "Classic Cars Club" in cell_1_text else 
                    "Address not found or incorrectly formatted")
    return (False, "Could not find appropriate table")

def check_formatting(doc):
    """Check text formatting in the first table (Criteria 4)"""
    for table in doc.tables:
        if len(table.rows) == 1 and len(table.rows[0].cells) == 2:
            cell = table.rows[0].cells[0]
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if "Classic Cars Club" in run.text:
                        is_correct = run.bold and run.font.size and run.font.size.pt == 14
                        return (is_correct, 
                                "Correct formatting applied" if is_correct else 
                                "Formatting incorrect or missing")
    return (False, "Could not find text to check formatting")

def check_bullet_points(doc):
    """Check for properly formatted bullet points (Criteria 8)"""
    required_bullets = {
        "Free entry to local and regional shows",
        "A 30% entry discount on the national show",
        "A 25% discount on merchandise purchases",
        "A free Classic Cars Club plaque",
        "A free Classic Cars Club license plate frame"
    }
    
    found_bullets = set()
    for paragraph in doc.paragraphs:
        if paragraph.style.name and "List" in paragraph.style.name:
            found_bullets.add(paragraph.text.strip())
    
    all_found = required_bullets.issubset(found_bullets)
    return (all_found, 
            "All required bullet points found" if all_found else 
            "Missing or incorrect bullet points")

def check_main_table(doc):
    """Check the main three-column table formatting (Criteria 9-17)"""
    results = []
    
    for table in doc.tables:
        if len(table.columns) == 3:  # Found our target table
            # Check column widths (Criteria 9)
            widths = [cell._tc.tcPr.tcW.w for cell in table.rows[0].cells]
            correct_widths = (abs(widths[0] - 1.5 * 1440) < 100 and
                            abs(widths[1] - 2.25 * 1440) < 100 and
                            abs(widths[2] - 1.0 * 1440) < 100)
            results.append(("Column widths correct", correct_widths))
            
            # Check for header row formatting (Criteria 15-17)
            if len(table.rows) > 1:
                header_row = table.rows[1]  # Second row should be header
                
                # Check shading
                has_shading = False
                for cell in header_row.cells:
                    tc_pr = cell._tc.get_or_add_tcPr()
                    shading = tc_pr.find("w:shd", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                    if shading is not None and shading.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill") == "D9D9D9":
                        has_shading = True
                results.append(("Header row shading", has_shading))
                
                # Check bold formatting
                all_bold = True
                for cell in header_row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if not run.bold:
                                all_bold = False
                results.append(("Header row bold formatting", all_bold))
                
            # Check merged header cell (Criteria 12-13)
            if len(table.rows) > 0:
                first_row = table.rows[0]
                is_merged = len(first_row.cells) == 1
                if is_merged:
                    cell_text = first_row.cells[0].text.strip()
                    has_correct_text = cell_text == "Available Partner Discounts"
                    results.append(("Merged header cell", is_merged and has_correct_text))
            
    return results

def generate_report(doc):
    """Generate a comprehensive report of all criteria checks"""
    results = []
    
    # Check criteria 2
    table_result, table_msg = check_table_above_date(doc)
    results.append(("2. Table above date", table_result, table_msg))
    
    # Check criteria 3
    address_result, address_msg = check_address_in_table(doc)
    results.append(("3. Address in table", address_result, address_msg))
    
    # Check criteria 4
    format_result, format_msg = check_formatting(doc)
    results.append(("4. Text formatting", format_result, format_msg))
    
    # Check criteria 8
    bullets_result, bullets_msg = check_bullet_points(doc)
    results.append(("8. Bullet points", bullets_result, bullets_msg))
    
    # Check criteria 9-17
    main_table_results = check_main_table(doc)
    for idx, (check_name, result) in enumerate(main_table_results):
        results.append((f"{idx + 9}. {check_name}", result, 
                       "Requirement met" if result else "Requirement not met"))
    
    return pd.DataFrame(results, columns=["Criteria", "Passed", "Details"])

if word_file_student:
    try:
        with open("student.docx", "wb") as f:
            f.write(word_file_student.getbuffer())
        
        doc = Document("student.docx")
        
        st.subheader("Format Check Results")
        results_df = generate_report(doc)
        
        # Display results with color coding
        st.dataframe(
            results_df,
            column_config={
                "Criteria": "Requirement",
                "Passed": st.column_config.CheckboxColumn(
                    "Status",
                    help="Whether the requirement was met",
                    default=False,
                ),
                "Details": "Additional Information"
            },
            hide_index=True
        )
        
        # Calculate total score
        total_points = len(results_df) * 10
        earned_points = sum(results_df["Passed"]) * 10
        
        st.write(f"### Total Score: {earned_points}/{total_points}")
        
        if earned_points < total_points:
            st.warning("Some requirements were not met. Please review the details above.")
        else:
            st.success("All requirements met! Great job!")
            
    except Exception as e:
        st.error(f"An error occurred while checking the document: {str(e)}")
    finally:
        if os.path.exists("student.docx"):
            os.remove("student.docx")
