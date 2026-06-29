import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import streamlit as st
import pandas as pd
import os
from openpyxl.styles import PatternFill

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
                master.fill = yellow_fill



    wb.save("TIMETABLE.xlsx")

    print(f"\n✓ Done — {cleared} cell(s) cleared. Saved to 'TIMETABLE.xlsx'.")


def course_name_from_code(course_code: str) -> str:
    return courses[course_code]["Course Name"]

courses = {
    "FIN 6002": {
        "Instructor": "PRS",
        "Course Name": "Commercial Banking",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. Pradeepta Sethi", "Sections": ["SA", "SB"]}
        ],
        "Students": 141,
        "Venue": "LG 18"
    },
    "FIN 6004": {
        "Instructor": "KPD",
        "Course Name": "Financial Derivatives",
        "Credits": 2,
        "Faculty": [
            {"Name": "Prof. Krishna Prasad", "Sections": ["SA", "SB"]}
        ],
        "Students": 124,
        "Venue": "LG 11"
    },
    "FIN 6022": {
        "Instructor": "VNG",
        "Course Name": "Financial Statement Analysis",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. Vinay Goyal", "Sections": ["SA", "SB"]}
        ],
        "Students": 141,
        "Venue": "LG 18"
    },
    "FIN 6006": {
        "Instructor": "SJK",
        "Course Name": "Fixed Income Securities",
        "Credits": 2,
        "Faculty": [
            {"Name": "Prof. Sanjeev Kumar", "Sections": ["SA", "SB"]}
        ],
        "Students": 82,
        "Venue": "GF 22"
    },
    "FIN 6005": {
        "Instructor": "DPD",
        "Course Name": "Security Valuation",
        "Credits": 2,
        "Faculty": [
            {"Name": "Prof. M. Durga Prasad", "Sections": ["SA", "SB"]}
        ],
        "Students": 133,
        "Venue": "LG 11"
    },
    "ITS 6009": {
        "Instructor": "BSR",
        "Course Name": "Emerging Technologies for Managers",
        "Credits": 2,
        "Faculty": [
            {"Name": "Prof. B Sreevatsa", "Sections": ["SA", "SB"]}
        ],
        "Students": 104,
        "Venue": "GF 22"
    },
    "ITS 6002": {
        "Instructor": "GRN",
        "Course Name": "IT Risk Management and Cyber Security",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. Gurudutt S Nayak", "Sections": ["S1", "S2", "S3", "S4"]}
        ],
        "Students": 274,
        "Venue": "LG 18"
    },
    "ITS 6001": {
        "Instructor": "AVK",
        "Course Name": "Technology Consulting & Business Analysis",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. Avant Kumar", "Sections": ["S1", "S2", "S3"]}
        ],
        "Students": 209,
        "Venue": "LG 11"
    },
    "ITS 6007": {
        "Instructor": "AMK",
        "Course Name": "Digital Transformation",
        "Credits": 2,
        "Faculty": [
            {"Name": "Prof. Amit Kumar", "Sections": []}
        ],
        "Students": 35,
        "Venue": "GF 22"
    },
    "ITS 6012": {
        "Instructor": "VVJ",
        "Course Name": "Digital Platform and Technology Ecosystem",
        "Credits": 2,
        "Faculty": [
            {"Name": "Prof. Vivek Kumar Jha", "Sections": []}
        ],
        "Students": 39,
        "Venue": "GF 22"
    },
    "ANT 6009": {
        "Instructor": "KKB",
        "Course Name": "Data Visualization",
        "Credits": 2,
        "Faculty": [
            {"Name": "Prof. Kartikeya Bolar Pramoda", "Sections": ["SA", "SB"]}
        ],
        "Students": 82,
        "Venue": "ANALYTICS LAB"
    },
    "ANT 6010": {
        "Instructor": "KKB",
        "Course Name": "Foundations of Business Analytics",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. Kartikeya Bolar Pramoda", "Sections": ["SA", "SB"]}
        ],
        "Students": 67,
        "Venue": "ANALYTICS LAB"
    },
    "MKT 6006": {
        "Instructor": "SNR / STY",
        "Course Name": "Consumer Behaviour",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. N Srinivasa Reddy", "Sections": ["S1"]},
            {"Name": "Prof. Satyaki Datta", "Sections": ["S2", "S3"]}
        ],
        "Students": 161,
        "Venue": {
            "S1": "LG 18",
            "S2": "LG 11",
            "S3": "LG 11"
        }
    },
    "MKT 6004": {
        "Instructor": "JVN / SLP",
        "Course Name": "Brand Management",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. Jeevan J Arakal", "Sections": ["S1"]},
            {"Name": "Prof. Shilpa Praveen", "Sections": ["S2", "S3"]}
        ],
        "Students": 196,
        "Venue": {
            "S1": "LG 18",
            "S2": "GF 22",
            "S3": "GF 22"
        }
    },
    "MKT 6003": {
        "Instructor": "AKL",
        "Course Name": "Services Marketing",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. Ashish Kalra", "Sections": ["SA", "SB"]}
        ],
        "Students": 130,
        "Venue": "GF 22"
    },
    "MKT 6009": {
        "Instructor": "SLP / JYT / SNR",
        "Course Name": "Sales & Distribution Management",
        "Credits": 3,
        "Faculty": [
            {"Name": "Prof. Shilpa Praveen", "Sections": ["S1", "S2 (sessions 13-24)"]},
            {"Name": "Prof. Jayanthi Thanigan", "Sections": ["S3"]},
            {"Name": "Prof. N Srinivasa Reddy", "Sections": ["S4", "S2 (sessions 1-12)"]}
        ],
        "Students": 221,
        "Venue": {
            "S1": "GF 22",
            "S2": "GF 22",
            "S3": "LG 18",
            "S4": "LG 18"
        }
    },
    "MKT 6005": {
        "Instructor": "SNR",
        "Course Name": "Digital Marketing",
        "Credits": 2,
        "Faculty": [
            {"Name": "Prof. N Srinivasa Reddy", "Sections": ["SA", "SB"]}
        ],
        "Students": 133,
        "Venue": "LG 18"
    },
}

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

def get_subject_classes(codes_list) -> list:

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



def editing_page():

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
                

