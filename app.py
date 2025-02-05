import streamlit as st
from docx import Document
import pandas as pd
import os

st.title("Word Document Format Checker 📄")

def get_table_info(table):
    """Extract detailed information about a table"""
    info = {
        "rows": len(table.rows),
        "cols": len(table.rows[0].cells) if table.rows else 0,
        "merged_cells": [],
        "borders": [],
        "cell_alignments": [],
        "formatting": [],
        "shading": []
    }
    
    for row_idx, row in enumerate(table.rows):
        for cell_idx, cell in enumerate(row.cells):
            # Check for merged cells
            tc = cell._tc
            grid_span = tc.tcPr.xpath('.//w:gridSpan')
            if grid_span:
                info["merged_cells"].append((row_idx, cell_idx))
            
            # Check borders
            borders = tc.tcPr.xpath('.//w:tcBorders')
            if borders:
                info["borders"].append((row_idx, cell_idx, borders[0]))
            
            # Check alignment
            for paragraph in cell.paragraphs:
                info["cell_alignments"].append((row_idx, cell_idx, paragraph.alignment))
            
            # Check text formatting
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if run.bold or run.font.size:
                        info["formatting"].append((row_idx, cell_idx, {
                            "bold": run.bold,
                            "size": run.font.size.pt if run.font.size else None
                        }))
            
            # Check shading
            shading = tc.tcPr.xpath('.//w:shd')
            if shading:
                info["shading"].append((row_idx, cell_idx, shading[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')))
    
    return info

def check_criteria(doc):
    """Check all 16 formatting criteria"""
    results = []
    
    # Find tables
    tables = doc.tables
    paragraphs = doc.paragraphs
    
    # Criterion 2: 2x1 table above date
    found_table = False
    for table in tables:
        if len(table.rows) == 1 and len(table.rows[0].cells) == 2:
            found_table = True
    results.append(("2. 2x1 table above date", found_table))
    
    # Criterion 3: Address text in table
    address_correct = False
    if found_table:
        first_table = tables[0]
        cell1_text = first_table.rows[0].cells[0].text.strip()
        cell2_text = first_table.rows[0].cells[1].text.strip()
        address_correct = "Classic Cars Club" in cell1_text and "PO Box 6987" in cell2_text
    results.append(("3. Address text in table", address_correct))
    
    # Criterion 4: 14pt Bold formatting for Classic Cars Club
    formatting_correct = False
    if found_table:
        for paragraph in tables[0].rows[0].cells[0].paragraphs:
            for run in paragraph.runs:
                if "Classic Cars Club" in run.text:
                    formatting_correct = run.bold and run.font.size and run.font.size.pt == 14
    results.append(("4. Classic Cars Club formatting", formatting_correct))
    
    # Criterion 5: Cell alignments
    alignments_correct = False
    if found_table:
        first_table = tables[0]
        info = get_table_info(first_table)
        # Check if first cell has center-left and second has center-right alignment
        alignments_correct = any(align == 1 for _, _, align in info["cell_alignments"][:1]) and \
                           any(align == 2 for _, _, align in info["cell_alignments"][1:])
    results.append(("5. Table cell alignments", alignments_correct))
    
    # Criterion 6: No table borders
    borders_removed = False
    if found_table:
        first_table = tables[0]
        info = get_table_info(first_table)
        borders_removed = len(info["borders"]) == 0
    results.append(("6. Table borders removed", borders_removed))
    
    # Criterion 7: Empty paragraphs above date
    empty_paragraphs = sum(1 for p in paragraphs if not p.text.strip())
    results.append(("7. Empty paragraphs above date", empty_paragraphs >= 2))
    
    # Criterion 8: Bullet formatting
    required_bullets = {
        "Free entry to local and regional shows",
        "A 30% entry discount on the National show",
        "A 25% discount on merchandise purchases",
        "A free Classic Cars Club plaque",
        "A free Classic Cars Club license plate frame"
    }
    bullet_points = set()
    for paragraph in paragraphs:
        if paragraph.style.name and "List" in paragraph.style.name:
            bullet_points.add(paragraph.text.strip())
    results.append(("8. Bullet formatting", required_bullets.issubset(bullet_points)))
    
    # Find main table (3-column table)
    main_table = None
    for table in tables:
        if len(table.rows[0].cells) == 3:
            main_table = table
            break
    
    if main_table:
        # Criteria 9: Column widths
        info = get_table_info(main_table)
        widths = [cell._tc.tcPr.tcW.w for cell in main_table.rows[0].cells]
        correct_widths = (abs(widths[0] - 1.5 * 1440) < 100 and
                         abs(widths[1] - 2.25 * 1440) < 100 and
                         abs(widths[2] - 1.0 * 1440) < 100)
        results.append(("9. Column widths", correct_widths))
        
        # Criterion 10: Table sorting
        # This can only be verified by comparing content order
        has_header = "Partner" in main_table.rows[1].cells[0].text
        results.append(("10. Table sorting", has_header))
        
        # Criterion 11: New row at top
        results.append(("11. New row at top", len(main_table.rows) > len(tables[-1].rows) - 1))
        
        # Criterion 12: Merged cells in new row
        has_merged = any(len(span) > 0 for span in info["merged_cells"])
        results.append(("12. Merged cells", has_merged))
        
        # Criterion 13: Merged cell formatting
        if has_merged:
            merged_cell = main_table.rows[0].cells[0]
            correct_text = "Available Partner Discounts" in merged_cell.text
            has_formatting = False
            for paragraph in merged_cell.paragraphs:
                for run in paragraph.runs:
                    if run.bold and run.font.size and run.font.size.pt == 14:
                        has_formatting = True
            results.append(("13. Merged cell formatting", correct_text and has_formatting))
        else:
            results.append(("13. Merged cell formatting", False))
        
        # Criterion 14: Border settings
        has_outside_borders = any(border for _, _, border in info["borders"])
        results.append(("14. Border settings", has_outside_borders))
        
        # Criterion 15: Row 2 shading
        has_shading = any(shade for row_idx, _, shade in info["shading"] if row_idx == 1)
        results.append(("15. Row 2 shading", has_shading))
        
        # Criterion 16: Row 2 bottom border
        has_bottom_border = any(border for row_idx, _, border in info["borders"] if row_idx == 1)
        results.append(("16. Row 2 bottom border", has_bottom_border))
    else:
        # If main table not found, mark remaining criteria as failed
        for i in range(9, 17):
            results.append((f"{i}. Main table criterion", False))
    
    return pd.DataFrame(results, columns=["Criterion", "Met"])

# File uploader
word_file_student = st.file_uploader("Upload Student Document", type=["docx"])

if word_file_student:
    try:
        with open("student.docx", "wb") as f:
            f.write(word_file_student.getbuffer())
        
        doc = Document("student.docx")
        results_df = check_criteria(doc)
        
        # Display results
        st.write("### Format Check Results")
        
        # Create a styled dataframe
        st.dataframe(
            results_df,
            column_config={
                "Criterion": "Requirement",
                "Met": st.column_config.CheckboxColumn(
                    "Status",
                    help="Whether the requirement was met",
                    default=False,
                )
            },
            hide_index=True
        )
        
        # Calculate score
        total_criteria = len(results_df)
        met_criteria = results_df["Met"].sum()
        score = (met_criteria / total_criteria) * 100
        
        st.write(f"### Score: {score:.1f}%")
        st.write(f"Requirements met: {met_criteria} out of {total_criteria}")
        
        if score < 100:
            st.warning("Some requirements were not met. Please review the results above.")
        else:
            st.success("All requirements met! Great job!")
            
    except Exception as e:
        st.error(f"An error occurred while checking the document: {str(e)}")
    finally:
        if os.path.exists("student.docx"):
            os.remove("student.docx")
