import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import streamlit as st
from pos_analysis import render_retailer_page

st.set_page_config(
    page_title="ハンズ | POS分析",
    page_icon="🔧",
    layout="wide",
)

render_retailer_page("ハンズ")
