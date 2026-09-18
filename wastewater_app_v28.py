import os
import datetime
import io
import re
import base64
import json
import urllib.request
import urllib.error
import pandas as pd
import numpy as np
import streamlit as st
import streamlit.components.v1 as components
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. ページ基本設定 & 明るい黄緑・薄緑ベースデザイン (深緑排他)
# ==========================================
st.set_page_config(
    page_title="水処理点検・分析統合システム (KBL Wastewater Management)",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# カスタムCSSスタイリング (黄緑・薄緑・白・グレーベース)
st.markdown("""
<style>
    /* メインヘッダー */
    .main-header {
        font-size: 26px;
        font-weight: bold;
        color: #2E7D32;
        padding-bottom: 8px;
        border-bottom: 3.5px solid #66BB6A;
        margin-bottom: 15px;
    }
    /* サブヘッダー */
    .sub-header {
        font-size: 18px;
        font-weight: bold;
        color: #388E3C;
        margin-top: 15px;
        margin-bottom: 10px;
        padding-left: 8px;
        border-left: 4px solid #8BC34A;
    }
    /* メトリックカード */
    .metric-card {
        background-color: #F1F8E9;
        border-radius: 8px;
        padding: 12px;
        border-left: 5px solid #66BB6A;
    }
    /* プライマリボタン (明るいフレッシュグリーン) */
    .stButton>button {
        background-color: #43A047;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        border: none;
        padding: 8px 16px;
    }
    .stButton>button:hover {
        background-color: #66BB6A;
        color: white;
    }
    /* タブのデザイン調整 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #F1F8E9;
        border-radius: 6px 6px 0px 0px;
        padding: 8px 18px;
        color: #388E3C;
        font-weight: bold;
    }
    .stTabs [aria-selected="true"] {
        background-color: #43A047 !important;
        color: white !important;
    }
    /* サイドバーの背景トーン */
    [data-testid="stSidebar"] {
        background-color: #F7FAF7;
    }
    /* インフォボックスのトーン */
    .stAlert {
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1.5 簡易パスワード認証保護 (Security - StreamlitSecretNotFoundError安全保護)
# ==========================================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.markdown("<div class='main-header'>🔒 KBL 排水管理システム - ログイン</div>", unsafe_allow_html=True)
    st.markdown("#### 🔑 パスワード認証")
    pwd_input = st.text_input("パスワード", type="password", key="login_pwd_key")
    
    # st.secrets の安全保護 (secrets.toml 未設定時もアプリが白画面エラーにならないよう防御)
    try:
        correct_pwd = st.secrets.get("APP_PASSWORD", "kbl2026")
    except Exception:
        correct_pwd = "kbl2026"

    if pwd_input:
        if pwd_input == correct_pwd:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ パスワードが正しくありません。")
    st.stop()
