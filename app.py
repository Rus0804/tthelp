import streamlit as st
from display_media import display_default
from timetable_filter import filter_classes
from all_courses import get_course_dict
st.title("TAPMI Class Finder")
st.set_page_config(layout="wide")

file = None
with st.sidebar:
    file = st.file_uploader("Upload your timetable", type = '.xlsx')

def show_courses():
    st.write(get_course_dict(file))

default_page = st.Page(lambda: display_default(file), title = "Full Timetable", url_path="full-timetable")

filter_page = st.Page(lambda: filter_classes(file), title = "Filter Classes", url_path="filter-classes")


pg = st.navigation([default_page, filter_page, show_courses])
pg.run()