import streamlit as st
from display_media import display_default
from timetable_filter import filter_page


st.title("TAPMI Class Finder")
st.set_page_config(layout="wide")



default_page = st.Page(display_default, title = "Full Timetable")

pg = st.navigation([default_page, filter_page])
pg.run()