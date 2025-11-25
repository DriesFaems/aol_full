import streamlit as st
from utils import set_branding

st.set_page_config(
    page_title="WHU AoL Tools",
    page_icon="assets/whu-logo-icon.png",
    layout="wide"
)

set_branding()

st.title("Assurance of Learning (AoL) Tools")

st.markdown("""
Welcome to the WHU Assurance of Learning (AoL) Toolkit. This application consolidates three key tools to streamline the evaluation process:

### 1. Evaluation Sheet Creator
Upload a course roster (Excel) to generate a customized evaluation workbook. You can define assessment dimensions, weights, and rubrics.

### 2. Mail Merge Tool
Upload completed evaluation sheets to generate individual feedback documents for students. Supports both combined Word documents and individual files.

### 3. Aggregator
Upload multiple completed evaluation workbooks to aggregate scores and compute competency-level statistics across assignments.

---
**👈 Select a tool from the sidebar to get started.**
""")
