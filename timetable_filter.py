import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import streamlit as st
import pandas as pd
import os
from openpyxl.styles import PatternFill
from all_courses import TERM_4_COURSES as courses

def check_condition(cell_value: str, condition: dict) -> bool:
    return condition["course_code"] in cell_value and condition["class"] in cell_value

def apply_condition_set(cell_value: str, conditions: list) -> bool:
    """Return True if the cell should be cleared."""
    if not cell_value:
        return False
    results = [check_condition(cell_value, c) for c in conditions]
    return any(results)


def build_merge_map(ws):
    """Returns a dict mapping (row, col) → (master_row, master_col) for every merged cell."""
    merge_map = {}
    for merge_range in ws.merged_cells.ranges:
        master = (merge_range.min_row, merge_range.min_col)
        for row in range(merge_range.min_row, merge_range.max_row + 1):
            for col in range(merge_range.min_col, merge_range.max_col + 1):
                merge_map[(row, col)] = master
    return merge_map

def process_excel(conditions:  list):
    wb = load_workbook("TERM 4 MBA TT.xlsx")
    ws = wb["TERM 4 MBA"]

    row_start: int = 9
    row_end: int = 53
    col_start: int = 3
    col_end: int = 10


    cleared = 0
    merge_map = build_merge_map(ws)
    already_cleared = set()  # avoid hitting the same master cell twice

    merge_map     = build_merge_map(ws)
    already_seen  = set()   # master coordinates already evaluated
    cleared       = 0

    colors = ["yellow", "blue", "green", "red", "orange", "pink", "purple"]

    # Map color names to hex codes
    color_hex_map = {
        "yellow": "FFFF00",
        "blue":   "0070C0",
        "green":  "00B050",
        "red":    "FF0000",
        "orange": "FFA500",
        "pink":   "FFC0CB",
        "purple": "800080",
    }

    set_color = 0
    subject_color_map = {}
    for condition in conditions:
        subject_color_map[condition['course_code']] = colors[set_color]
        set_color += 1
    
    
    # Build a PatternFill for each color name
    fill_map = {
        name: PatternFill(start_color=hex_code, end_color=hex_code, fill_type="solid")
        for name, hex_code in color_hex_map.items()
    }
    

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=col)

            if isinstance(cell, MergedCell):
                master_coords = merge_map.get((row, col))
                if not master_coords or master_coords in already_seen:
                    continue
                already_seen.add(master_coords)
                master = ws.cell(*master_coords)
            else:
                if (row, col) in already_seen:
                    continue
                already_seen.add((row, col))
                master = cell

            if not apply_condition_set(cell.value, conditions):
                print(f"  Clearing {cell.coordinate}  ← was: {repr(cell.value)}")
                master.value = None
                cleared += 1
            else:
                cell_content = master.value
                course_code = cell_content[:8]
                course_name = course_name_from_code(course_code)
                master.value = master.value.replace(course_code, course_name + f"({course_code[:3]})")
                master.fill = fill_map[subject_color_map[course_code]]



    wb.save("TIMETABLE.xlsx")

    print(f"\n✓ Done — {cleared} cell(s) cleared. Saved to 'TIMETABLE.xlsx'.")


def course_name_from_code(course_code: str) -> str:
    return courses[course_code]["Course Name"]


def get_subject_codes() -> list:
    st.subheader("Select Classes")
    major = st.radio("Pick your major", ["Marketing - MKT", "Finance - FIN", "Information Technology - ITS", "Accounting - ANT"], horizontal = True)
    major_courses =  [(code, courses[code]["Course Name"]) for code in courses.keys() if major[-3:] in code]

    major_selected = st.pills("Pick major subjects", [course_name for code, course_name in major_courses], selection_mode = "multi")
    major_codes = [code for code, course_name in major_courses if course_name in major_selected]
        
    minor_courses = [(code, courses[code]["Course Name"]) for code in courses.keys() if major[-3:] not in code]
    minor_selected = st.pills("Pick minor subjects", [course_name for code, course_name in minor_courses], selection_mode = "multi")
    minor_codes = [code for code, course_name in minor_courses if course_name in minor_selected]

    codes_list = [code for code in courses.keys() if code in major_codes or code in minor_codes]
    return codes_list

def get_subject_classes(codes_list: list) -> list:

    selected_subjects = []

    for code in codes_list:
        subject_classes = []
        for faculty in courses[code]["Faculty"]:
            for _class in faculty["Sections"]:
                if _class[:2] not in subject_classes:
                    subject_classes.append(_class[:2])
        subject_class = st.radio(f"Pick Class for {courses[code]["Course Name"]} ({code})", subject_classes, horizontal = True, key = code)
        selected_subjects.append({"course_code": code, "class": subject_class})

    return selected_subjects



def filter_page():

    codes_list = get_subject_codes()
    selected_subjects = get_subject_classes(codes_list)
    

    if len(selected_subjects) > 0:
        subject_data = dict([(code, courses[code]) for code in courses.keys() if code in codes_list])
        data = pd.DataFrame(subject_data)
        st.table(data.loc[["Course Name", "Credits", "Instructor"]])

        if st.button("Filter Timetable"):
            process_excel(selected_subjects)
            file_path = "TIMETABLE.xlsx"
            if os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    st.download_button(
                        label="Download Timetable",
                        data=f,
                        file_name="TIMETABLE.xlsx"
                    )
                

