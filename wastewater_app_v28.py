import streamlit.components.v1 as components
import base64
import json
import urllib.request
import urllib.error
import os
import datetime
import io
import re
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------
# 1. ページ基本設定 & 明るい黄緑・薄緑ベースデザイン (深緑排他)
# ------------------------------------------
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

# ------------------------------------------
# 1.5 簡易パスワード認証保護 (Security)
# ------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.markdown("<div class='main-header'>🔒 KBL 排水管理システム - ログイン</div>", unsafe_allow_html=True)
    st.markdown("#### 🔑 パスワード認証")
    pwd_input = st.text_input("パスワード", type="password", key="login_pwd_key")
    correct_pwd = st.secrets.get("APP_PASSWORD", "kbl2026")
    if pwd_input:
        if pwd_input == correct_pwd:
            st.session_state["authenticated"] = True
            st.rerun()
        elif pwd_input:
            st.error("❌ パスワードが正しくありません。")
    st.stop()

# ------------------------------------------
# GitHub REST API 自動同期関数
# ------------------------------------------
def sync_to_github_api(file_path):
    """Syncs local file to GitHub repository using GitHub REST API and st.secrets['GITHUB_TOKEN']"""
    try:
        if "GITHUB_TOKEN" not in st.secrets:
            return False, "GITHUB_TOKEN is not set in st.secrets"
        
        token = st.secrets["GITHUB_TOKEN"]
        repo = st.secrets.get("GITHUB_REPO", "kbl-wastewater-app")
        branch = st.secrets.get("GITHUB_BRANCH", "main")
        target_filename = os.path.basename(file_path)
        
        url = f"https://api.github.com/repos/{repo}/contents/{target_filename}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "StreamlitApp"
        }
        
        # 1. Fetch current file SHA
        sha = None
        req_get = urllib.request.Request(f"{url}?ref={branch}", headers=headers, method="GET")
        try:
            with urllib.request.urlopen(req_get) as resp:
                res_data = json.loads(resp.read().decode("utf-8"))
                sha = res_data.get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                return False, f"HTTP Error {e.code} during SHA fetch"
                
        # 2. Base64 encode file content
        with open(file_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode("utf-8")
            
        payload = {
            "message": f"Auto-update wastewater data via App [{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]",
            "content": content_b64,
            "branch": branch
        }
        if sha:
            payload["sha"] = sha
            
        data_json = json.dumps(payload).encode("utf-8")
        req_put = urllib.request.Request(url, data=data_json, headers=headers, method="PUT")
        
        with urllib.request.urlopen(req_put) as resp:
            if resp.status in [200, 201]:
                return True, "Success"
            return False, f"HTTP Status {resp.status}"
    except Exception as ex:
        return False, str(ex)

# ------------------------------------------
# 自動計算関数 (有機割合、無機割合、沈降速度、水面積負荷、返送率、SVI、BOD除去率)
# ------------------------------------------
def calculate_auto_metrics(df_item_all, df_loc_all, selected_site_id, selected_sheet_type, form_data_dict, df_rec_latest, target_rep_id):
    calc_updates = {}
    
    def _parse_num(item_id):
        if item_id in form_data_dict:
            v_str = str(form_data_dict[item_id]).replace(',', '').replace('%', '').replace('<', '').replace('>', '').strip()
            try:
                return float(v_str)
            except ValueError:
                pass
        match = df_rec_latest[(df_rec_latest['Report_ID'] == target_rep_id) & (df_rec_latest['Item_ID'] == item_id)]
        if not match.empty:
            v_str = str(match.iloc[0]['Value']).replace(',', '').replace('%', '').replace('<', '').replace('>', '').strip()
            try:
                return float(v_str)
            except ValueError:
                pass
        return None

    def _get_raw_val(item_id):
        if item_id in form_data_dict:
            return str(form_data_dict[item_id]).strip()
        match = df_rec_latest[(df_rec_latest['Report_ID'] == target_rep_id) & (df_rec_latest['Item_ID'] == item_id)]
        if not match.empty:
            return str(match.iloc[0]['Value']).strip()
        return ""

    def _find_items(df_subset, query_str):
        exact = df_subset[df_subset['Item_Name'] == query_str]
        if not exact.empty:
            return exact
        prefix = df_subset[df_subset['Item_Name'].astype(str).str.startswith(query_str)]
        if not prefix.empty:
            return prefix
        contains = df_subset[df_subset['Item_Name'].astype(str).str.contains(query_str, regex=False)]
        return contains

    # 1. 有機割合 [%] (MLVSS/MLSS*100) & 無機割合 [%] (100 - 有機割合)
    locs_in_site = df_loc_all[df_loc_all['Site_ID'] == selected_site_id]['Loc_ID'].unique()
    for loc_id in locs_in_site:
        loc_items = df_item_all[(df_item_all['Loc_ID'] == loc_id) & (df_item_all['Sheet_Type'] == selected_sheet_type)]
        mlss_item = _find_items(loc_items, 'MLSS')
        if mlss_item.empty:
            mlss_item = _find_items(loc_items, 'MLSS(簡)')
        mlvss_item = _find_items(loc_items, 'MLVSS')
        org_item = _find_items(loc_items, '有機割合')
        inorg_item = _find_items(loc_items, '無機割合')
        
        if not mlss_item.empty and not mlvss_item.empty:
            mlss_v = _parse_num(mlss_item.iloc[0]['Item_ID'])
            mlvss_v = _parse_num(mlvss_item.iloc[0]['Item_ID'])
            if mlss_v and mlss_v > 0 and mlvss_v is not None:
                org_ratio = round((mlvss_v / mlss_v) * 100.0, 1)
                inorg_ratio = round(100.0 - org_ratio, 1)
                if not org_item.empty:
                    calc_updates[org_item.iloc[0]['Item_ID']] = f"{org_ratio}%"
                if not inorg_item.empty:
                    calc_updates[inorg_item.iloc[0]['Item_ID']] = f"{inorg_ratio}%"

    # 2. キャンパック（S004/S005）専用計算 (沈降速度, 水面積負荷, 沈降速度/水面積負荷, 返送率)
    if selected_site_id in ['S004', 'S005'] and selected_sheet_type == '点検管理表':
        site_items = df_item_all.merge(df_loc_all[df_loc_all['Site_ID'] == selected_site_id], on='Loc_ID')
        site_items = site_items[site_items['Sheet_Type'] == '点検管理表']
        
        isou_item = _find_items(site_items, '移送量')
        hensou_item = _find_items(site_items, '返送量')
        
        aeration_end = site_items[site_items['Loc_Name'] == '曝気槽（末端側）']
        mlss_item = _find_items(aeration_end, 'MLSS')
        temp_item = _find_items(aeration_end, '水温')
        sv30_item = _find_items(aeration_end, 'SV30')
        
        isou_v = _parse_num(isou_item.iloc[0]['Item_ID']) if not isou_item.empty else None
        hensou_v = _parse_num(hensou_item.iloc[0]['Item_ID']) if not hensou_item.empty else None
        mlss_v = _parse_num(mlss_item.iloc[0]['Item_ID']) if not mlss_item.empty else None
        temp_v = _parse_num(temp_item.iloc[0]['Item_ID']) if not temp_item.empty else None
        sv30_v = _parse_num(sv30_item.iloc[0]['Item_ID']) if not sv30_item.empty else None
        
        load_item = _find_items(site_items, '水面積負荷')
        return_ratio_item = _find_items(site_items, '返送率')
        settling_item = _find_items(site_items, '沈降速度')
        ratio_item = site_items[site_items['Item_Name'] == '沈降速度/水面積負荷']
        if ratio_item.empty:
            ratio_item = _find_items(site_items, '沈降速度/水面積負荷')
            
        if not settling_item.empty:
            settling_item = settling_item[~settling_item['Item_Name'].astype(str).str.contains('水面積負荷')]
        
        # 水面積負荷 = 移送量 / 143
        surf_load = None
        if isou_v is not None and isou_v > 0:
            surf_load = round(isou_v / 143.0, 2)
            if not load_item.empty:
                calc_updates[load_item.iloc[0]['Item_ID']] = str(surf_load)
                
        # 返送率 [%] = 返送量 / 移送量 * 100
        if isou_v is not None and isou_v > 0 and hensou_v is not None and hensou_v >= 0:
            r_ratio = round((hensou_v / isou_v) * 100.0, 1)
            if not return_ratio_item.empty:
                calc_updates[return_ratio_item.iloc[0]['Item_ID']] = f"{r_ratio}%"
                
        # 沈降速度 [m/h]
        v_settling = None
        if mlss_v is not None and mlss_v > 0 and temp_v is not None and temp_v > 0 and sv30_v is not None and sv30_v > 0:
            try:
                sv_term = (sv30_v * 10000.0) / mlss_v
                v_settling = (1.78 * (10**7) * (mlss_v**(-1.46)) * (temp_v**(0.853)) * (sv_term**(-0.804))) / 24.0
                v_settling = round(v_settling, 2)
                if not settling_item.empty:
                    calc_updates[settling_item.iloc[0]['Item_ID']] = str(v_settling)
            except Exception:
                pass
                
        # 沈降速度 / 水面積負荷
        if v_settling is not None and surf_load is not None and surf_load > 0:
            ratio_val = round(v_settling / surf_load, 2)
            if not ratio_item.empty:
                calc_updates[ratio_item.iloc[0]['Item_ID']] = str(ratio_val)

    # 3. SVI Calculation: SV30 * 10000 / MLSS
    sv30_mlss_map = {
        'S001': ('I0028', 'I0027', 'I0231'),
        'S002': ('I0075', 'I0074', 'I0232'),
        'S004': ('I0137', 'I0136', 'I0233'),
        'S005': ('I0204', 'I0203', 'I0234')
    }
    if selected_sheet_type == '点検管理表' and selected_site_id in sv30_mlss_map:
        sv30_id, mlss_id, svi_id = sv30_mlss_map[selected_site_id]
        sv30_v = _parse_num(sv30_id)
        mlss_v = _parse_num(mlss_id)
        if sv30_v is not None and sv30_v > 0 and mlss_v is not None and mlss_v > 0:
            svi_val = round((sv30_v * 10000.0) / mlss_v, 1)
            calc_updates[svi_id] = str(svi_val)

    # 4. BOD除去率 Calculation: (原水BOD - 処理水BOD) / 原水BOD * 100
    bod_rem_map = {
        'S001': ('I0049', 'I0054', 'I0235'),
        'S002': ('I0090', 'I0108', 'I0236'),
        'S004': ('I0160', 'I0177', 'I0237')
    }
    if selected_sheet_type == '計量証明' and selected_site_id in bod_rem_map:
        raw_bod_id, treated_bod_id, rem_item_id = bod_rem_map[selected_site_id]
        raw_v = _parse_num(raw_bod_id)
        treated_raw_str = _get_raw_val(treated_bod_id)
        
        if raw_v is not None and raw_v > 0 and treated_raw_str and treated_raw_str not in ['', '-', 'nan']:
            is_less = ('未満' in treated_raw_str) or ('<' in treated_raw_str)
            treated_clean = treated_raw_str.replace(',', '').replace('%', '').replace('<', '').replace('>', '').replace('未満', '').strip()
            treated_v = None
            try:
                treated_v = float(treated_clean)
            except ValueError:
                pass
                
            if treated_v is not None:
                if is_less or treated_v < 5.0:
                    rem_rate = round((raw_v - 5.0) / raw_v * 100.0, 1)
                    calc_updates[rem_item_id] = f"{rem_rate}%以上"
                else:
                    rem_rate = round((raw_v - treated_v) / raw_v * 100.0, 1)
                    calc_updates[rem_item_id] = f"{rem_rate}%"

    return calc_updates

# DBファイルパスの取得
def get_db_path():
    for fname in ["wastewater-appsheet-db-v3.xlsx", "wastewater-appsheet-db-v2.xlsx", "wastewater-appsheet-db.xlsx"]:
        if os.path.exists(fname):
            return fname
        abs_p = f"/workspace/artifacts/{fname}"
        if os.path.exists(abs_p):
            return abs_p
    return "wastewater-appsheet-db-v3.xlsx"

db_path = get_db_path()

# ------------------------------------------
# 2. データ読み込み ＆ キャッシュ処理 (パーセント・不等号クレンジング強化)
# ------------------------------------------
@st.cache_data(ttl=1)
def load_all_data(path):
    if not os.path.exists(path):
        st.error(f"データベースファイルが見つかりません: {path}")
        return None, None, None, None, None
    
    xls = pd.ExcelFile(path)
    df_site = pd.read_excel(xls, "Site Master")
    df_loc = pd.read_excel(xls, "Location Master")
    df_item = pd.read_excel(xls, "Item Master")
    df_rep = pd.read_excel(xls, "Daily Report")
    df_rec = pd.read_excel(xls, "Inspection Records")
    
    # 旧DBファイル(北越コーポレーション/S003)が読まれても自動排除
    df_site = df_site[~df_site["Site_Name"].astype(str).str.contains("北越") & (df_site["Site_ID"] != "S003")]
    df_loc = df_loc[~df_loc["Site_ID"].isin(["S003"]) & df_loc["Site_ID"].isin(df_site["Site_ID"])]
    df_item = df_item[df_item["Loc_ID"].isin(df_loc["Loc_ID"])]
    df_rep = df_rep[df_rep["Site_ID"].isin(df_site["Site_ID"])]
    df_rec = df_rec[df_rec["Report_ID"].isin(df_rep["Report_ID"])]

    # 型変換・整形
    df_rep["Date"] = pd.to_datetime(df_rep["Date"]).dt.strftime("%Y/%m/%d")
    
    val_str_series = df_rec["Value"].astype(str)
    has_comma_decimal = val_str_series.str.contains(r'^\d+,\d+$', regex=True)
    val_clean = val_str_series.copy()
    val_clean[has_comma_decimal] = val_clean[has_comma_decimal].str.replace(',', '.', regex=False)
    val_clean = (
        val_clean
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("<", "", regex=False)
        .str.replace(">", "", regex=False)
        .str.strip()
    )
    df_rec["Value_Num"] = pd.to_numeric(val_clean, errors="coerce")
    
    return df_site, df_loc, df_item, df_rep, df_rec

# ------------------------------------------
# 3. アプリメイン画面 & サイドバー設定 (全4現場完全切替)
# ------------------------------------------
df_site, df_loc, df_item, df_rep, df_rec = load_all_data(db_path)

if df_site is None:
    st.stop()

# サイドバー設定 (ヘッダー配色も統一)
st.sidebar.markdown("<h2 style='color:#2E7D32;'>⚙️ KBL 排水管理設定</h2>", unsafe_allow_html=True)

site_options = dict(zip(df_site["Site_ID"], df_site["Site_Name"]))
selected_site_id = st.sidebar.selectbox("① 対象現場（工場）を選択", options=list(site_options.keys()), format_func=lambda x: site_options[x])

sheet_type_options = ["点検管理表", "計量証明"]
selected_sheet_type = st.sidebar.selectbox("② 管理シート種別を選択", options=sheet_type_options)

st.sidebar.markdown("---")
st.sidebar.info(f"📍 選択中現場: **{site_options[selected_site_id]}**\\n📄 種別: **{selected_sheet_type}**")

# メインタイトル
st.markdown(f"<div class='main-header'>🌱 KBL 排水処理統合管理システム - {site_options[selected_site_id]}</div>", unsafe_allow_html=True)

# ------------------------------------------
# 4. タブ構築 (要件①〜⑦の全機能を4タブに整理)
# ------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📝 点検データ入力・異常値チェック",
    "📊 水質比較グラフ",
    "📑 ⑥ エクセル帳票ダウンロード",
    "🗃️ データベース全件閲覧"
])

# ------------------------------------------
# TAB 1: 点検データ入力 (要件①〜④) + 異常値チェック
# ------------------------------------------
with tab1:
    st.markdown("<div class='sub-header'>📝 点検結果の新規入力 ＆ 蓄積 (過去データとの異常値チェック機能付き)</div>", unsafe_allow_html=True)
    
    if "save_success_msg" in st.session_state:
        st.success(st.session_state["save_success_msg"])
        del st.session_state["save_success_msg"]

    col_input1, col_input2 = st.columns(2)

    with col_input1:
        input_date = st.date_input("③ 点検日を選択", datetime.date.today())
        input_date_str = input_date.strftime("%Y/%m/%d")

    with col_input2:
        locs_with_items = df_item[(df_item["Loc_ID"].isin(df_loc[df_loc["Site_ID"] == selected_site_id]["Loc_ID"])) & (df_item["Sheet_Type"] == selected_sheet_type)]["Loc_ID"].unique()
        site_locs = df_loc[(df_loc["Site_ID"] == selected_site_id) & (df_loc["Loc_ID"].isin(locs_with_items))].sort_values("Display_Order")
        if site_locs.empty:
            st.warning(f"⚠️ 【{site_options[selected_site_id]}】には「{selected_sheet_type}」の項目データが登録されていません。")
            selected_loc_id = None
        else:
            loc_options = dict(zip(site_locs["Loc_ID"], site_locs["Loc_Name"]))
            selected_loc_id = st.selectbox("④ 槽（測定箇所）を選択", options=list(loc_options.keys()), format_func=lambda x: loc_options[x], key="input_loc")

    st.markdown("---")
    if selected_loc_id:
        st.markdown(f"#### 📍 【{site_options[selected_site_id]}】 - 『{loc_options[selected_loc_id]}』 ({selected_sheet_type})")

        loc_items = df_item[(df_item["Loc_ID"] == selected_loc_id) & (df_item["Sheet_Type"] == selected_sheet_type)].sort_values("Display_Order")

        if loc_items.empty:
            st.warning("指定された槽・シート種別の点検項目が登録されていません。")
        else:
            existing_rep = df_rep[(df_rep["Site_ID"] == selected_site_id) & (df_rep["Date"] == input_date_str) & (df_rep["Sheet_Type"] == selected_sheet_type)]
            existing_values = {}
            if not existing_rep.empty:
                rep_id_exist = existing_rep.iloc[0]["Report_ID"]
                rec_exist = df_rec[df_rec["Report_ID"] == rep_id_exist]
                existing_values = dict(zip(rec_exist["Item_ID"], rec_exist["Value"]))
                st.info(f"💡 {input_date_str} の既存データが読み込まれました。必要に応じて内容を更新してください。")

            df_hist = df_rec[df_rec["Item_ID"].isin(loc_items["Item_ID"])].dropna(subset=["Value_Num"])
            stats_by_item = {}
            for item_id_key, group in df_hist.groupby("Item_ID"):
                vals = group["Value_Num"].values
                if len(vals) >= 2:
                    mean_v = np.mean(vals)
                    std_v = np.std(vals)
                    min_v = np.min(vals)
                    max_v = np.max(vals)
                    lower_b = max(0, mean_v - 2.5 * std_v) if mean_v >= 0 else mean_v - 2.5 * std_v
                    upper_b = mean_v + 2.5 * std_v
                    stats_by_item[item_id_key] = {
                        "mean": mean_v, "std": std_v, "min": min_v, "max": max_v,
                        "lower": lower_b, "upper": upper_b
                    }
                elif len(vals) == 1:
                    stats_by_item[item_id_key] = {
                        "mean": vals[0], "std": 0, "min": vals[0], "max": vals[0],
                        "lower": vals[0]*0.5, "upper": vals[0]*1.5
                    }

            st.info("💡 **自動計算項目（SVI・BOD除去率・有機割合・無機割合・水面積負荷・沈降速度・返送率など）** は、入力された基礎数値から保存時に**全自動計算されExcelデータベースへ直接格納**されます（入力ボックスは非表示としています）。")
            
            calc_item_names = ['有機割合', '有機割合 (%)', '無機割合', '沈降速度', '水面積負荷', '沈降速度/水面積負荷', '返送率', '返送率 (%)', 'SVI', 'BOD除去率']
            input_items = loc_items[~loc_items['Item_Name'].isin(calc_item_names)].sort_values("Display_Order")

            # スマホ用テンキー（inputmode="decimal"）自動適用 JavaScript (Form外部で安全に実行)
            components.html("""
            <script>
            function setInputMode() {
                const inputs = window.parent.document.querySelectorAll('input[type="text"]');
                inputs.forEach(input => {
                    input.setAttribute('inputmode', 'decimal');
                });
            }
            setTimeout(setInputMode, 300);
            setTimeout(setInputMode, 1000);
            </script>
            """, height=0)

            with st.form("inspection_input_form"):
                form_data = {}
                input_items_list = list(input_items.iterrows())

                for row_start in range(0, len(input_items_list), 3):
                    row_chunk = input_items_list[row_start : row_start + 3]
                    cols = st.columns(3)
                    for c_idx, (_, item) in enumerate(row_chunk):
                        item_id = item["Item_ID"]
                        item_name = item["Item_Name"]
                        vtype = item["Value_Type"]
                        unit = str(item["Unit_or_Options"]) if pd.notna(item["Unit_or_Options"]) else ""

                        default_val = existing_values.get(item_id, "")
                        label_str = item_name if vtype == "Enum" or not unit or unit == "nan" else f"{item_name} ({unit})"
                        col_target = cols[c_idx]

                        with col_target:
                            if vtype == "Enum":
                                opts = ["-", "無", "微少", "少", "中", "多"]
                                curr_idx = opts.index(default_val) if default_val in opts else 0
                                form_data[item_id] = st.selectbox(label_str, options=opts, index=curr_idx)
                            else:
                                init_num = None
                                if str(default_val) != "nan" and str(default_val).strip() not in ["", "-"]:
                                    try:
                                        init_num = float(str(default_val).replace(",", "").replace("%", "").replace("<", "").replace(">", "").strip())
                                    except ValueError:
                                        pass
                                form_data[item_id] = st.number_input(label_str, value=init_num, step=None, format=None, key=f"num_{item_id}")
                                if item_id in stats_by_item:
                                    st_info = stats_by_item[item_id]
                                    st.caption(f"💡 過去平均: {st_info['mean']:.2f} (目安: {st_info['lower']:.1f} 〜 {st_info['upper']:.1f})")

                notes_input = st.text_area("備考 (Notes)", value=existing_rep.iloc[0]["Notes"] if not existing_rep.empty and pd.notna(existing_rep.iloc[0]["Notes"]) else "")

                submit_btn = st.form_submit_button("💾 点検データを保存・蓄積する")

                if submit_btn:
                    anomalies = []
                    for item_id, val in form_data.items():
                        if val is not None:
                            val_str = str(val).strip().replace(",", "").replace("%", "").replace("<", "").replace(">", "")
                            if val_str and val_str not in ["-", "nan", "None"]:
                                try:
                                    num_v = float(val_str)
                                    if item_id in stats_by_item:
                                        st_info = stats_by_item[item_id]
                                        if num_v < st_info['lower'] or num_v > st_info['upper']:
                                            it_name = df_item[df_item["Item_ID"] == item_id].iloc[0]["Item_Name"]
                                            anomalies.append(f"・**{it_name}**: 入力値 {num_v} (過去通常範囲: {st_info['lower']:.1f} 〜 {st_info['upper']:.1f} / 平均: {st_info['mean']:.2f})")
                                except ValueError:
                                    pass

                    if anomalies:
                        st.warning("⚠️ **【異常値・誤入力の可能性あり】**\\n" + "\\n".join(anomalies) + "\\n\\n※数値に問題が無ければ、データは正常に保存されます。")

                    wb = openpyxl.load_workbook(db_path)
                    ws_rep = wb["Daily Report"]
                    ws_rec = wb["Inspection Records"]

                    df_rep_latest = pd.read_excel(db_path, "Daily Report")
                    df_rec_latest = pd.read_excel(db_path, "Inspection Records")

                    exist_rep_match = df_rep_latest[(df_rep_latest["Site_ID"] == selected_site_id) & (df_rep_latest["Date"] == input_date_str) & (df_rep_latest["Sheet_Type"] == selected_sheet_type)]

                    if not exist_rep_match.empty:
                        target_rep_id = exist_rep_match.iloc[0]["Report_ID"]
                        for row in ws_rep.iter_rows(min_row=2):
                            if str(row[0].value) == str(target_rep_id):
                                row[4].value = ""
                                row[5].value = notes_input
                                break
                    else:
                        max_rep_num = 0
                        for rid in df_rep_latest["Report_ID"].dropna():
                            m = re.search(r"\\d+", str(rid))
                            if m:
                                max_rep_num = max(max_rep_num, int(m.group()))
                        target_rep_id = f"REP{max_rep_num+1:05d}"
                        ws_rep.append([target_rep_id, input_date_str, selected_site_id, selected_sheet_type, "", notes_input])

                    existing_item_ids = set(df_rec_latest[df_rec_latest["Report_ID"] == target_rep_id]["Item_ID"])

                    max_rec_num = 0
                    for rcid in df_rec_latest["Rec_ID"].dropna():
                        m = re.search(r"\\d+", str(rcid))
                        if m:
                            max_rec_num = max(max_rec_num, int(m.group()))

                    auto_calcs = calculate_auto_metrics(df_item, df_loc, selected_site_id, selected_sheet_type, form_data, df_rec_latest, target_rep_id)
                    final_save_data = form_data.copy()
                    final_save_data.update(auto_calcs)

                    for item_id, val in final_save_data.items():
                        if val is None or str(val).strip() in ["", "-", "nan", "None"]:
                            continue
                        if item_id in existing_item_ids:
                            for row in ws_rec.iter_rows(min_row=2):
                                if str(row[1].value) == str(target_rep_id) and str(row[2].value) == str(item_id):
                                    row[3].value = str(val).strip()
                                    break
                        else:
                            max_rec_num += 1
                            new_rec_id = f"REC{max_rec_num:06d}"
                            ws_rec.append([new_rec_id, target_rep_id, item_id, str(val).strip()])

                    wb.save(db_path)
                    st.cache_data.clear()
                    
                    sync_ok, sync_msg = sync_to_github_api(db_path)
                    if sync_ok:
                        st.session_state["save_success_msg"] = f"✅ {input_date_str} 『{loc_options[selected_loc_id]}』 の点検データを正常に保存しました（GitHubへの自動同期も成功しました）！"
                    else:
                        st.session_state["save_success_msg"] = f"✅ {input_date_str} 『{loc_options[selected_loc_id]}』 の点検データを正常に保存しました（ローカル更新完了 / GitHub同期: {sync_msg}）"
                    st.rerun()

# ------------------------------------------
# TAB 2: 動的2軸・複数槽比較水質グラフ (要件⑤)
# ------------------------------------------
with tab2:
    st.markdown("<div class='sub-header'>📊 期間・複数槽(最大3箇所)・1軸/2軸の比較水質グラフ</div>", unsafe_allow_html=True)

    g_locs_with_items = df_item[(
        df_item["Loc_ID"].isin(df_loc[df_loc["Site_ID"] == selected_site_id]["Loc_ID"])
    ) & (df_item["Sheet_Type"] == selected_sheet_type)]["Loc_ID"].unique()

    g_site_locs = df_loc[(
        df_loc["Site_ID"] == selected_site_id
    ) & (df_loc["Loc_ID"].isin(g_locs_with_items))].sort_values("Display_Order")

    if g_site_locs.empty:
        st.warning("⚠️ 該当する槽・測定箇所データが存在しません。")
    else:
        graph_loc_options = dict(zip(g_site_locs["Loc_ID"], g_site_locs["Loc_Name"]))
        all_loc_ids = list(graph_loc_options.keys())

        col_g1, col_g2, col_g3 = st.columns([2.5, 2, 2])

        with col_g1:
            selected_graph_locs = st.multiselect(
                "比較対象の槽・測定箇所 (最大3つまで選択可)",
                options=all_loc_ids,
                default=all_loc_ids[:1],
                format_func=lambda x: graph_loc_options[x],
                max_selections=3
            )

        with col_g2:
            dates = pd.to_datetime(df_rep[(df_rep["Site_ID"] == selected_site_id) & (df_rep["Sheet_Type"] == selected_sheet_type)]["Date"]).sort_values().dropna()
            if not dates.empty:
                min_d = dates.min().date()
                max_d = dates.max().date()
                start_d, end_d = st.date_input("表示期間を選択", [min_d, max_d], min_value=min_d, max_value=max_d)
            else:
                start_d = datetime.date.today() - datetime.timedelta(days=30)
                end_d = datetime.date.today()

        with col_g3:
            axis_mode = st.radio(" Y軸モード選択", ["1軸 (同一スケール)", "2軸 (別スケール分析)"], horizontal=True)

        if not selected_graph_locs:
            st.info("比較する槽を選択してください。")
        else:
            loc_avail_items = df_item[(
                df_item["Loc_ID"].isin(selected_graph_locs)
            ) & (df_item["Sheet_Type"] == selected_sheet_type)]["Item_Name"].unique()

            col_y1, col_y2 = st.columns(2)
            with col_y1:
                y1_item_name = st.selectbox("📈 Y軸 1 (左軸) の測定項目", options=loc_avail_items)
            with col_y2:
                if axis_mode == "2軸 (別スケール分析)":
                    y2_item_name = st.selectbox("📉 Y軸 2 (右軸) の測定項目", options=[i for i in loc_avail_items if i != y1_item_name])
                else:
                    y2_item_name = None

            target_reps = df_rep[(
                df_rep["Site_ID"] == selected_site_id
            ) & (
                df_rep["Sheet_Type"] == selected_sheet_type
            )].copy()
            target_reps["Date_dt"] = pd.to_datetime(target_reps["Date"]).dt.date
            target_reps = target_reps[(target_reps["Date_dt"] >= start_d) & (target_reps["Date_dt"] <= end_d)].sort_values("Date_dt")

            if target_reps.empty:
                st.warning("選択された期間内にデータが存在しません。")
            else:
                fig = make_subplots(specs=[[{"secondary_y": True}]]) if axis_mode == "2軸 (別スケール分析)" else go.Figure()

                colors = ["#2E7D32", "#1E88E5", "#FB8C00", "#E53935", "#8E24AA", "#00ACC1"]
                c_idx = 0

                for loc_id in selected_graph_locs:
                    loc_name = graph_loc_options[loc_id]
                    y1_item_row = df_item[(df_item["Loc_ID"] == loc_id) & (df_item["Item_Name"] == y1_item_name) & (df_item["Sheet_Type"] == selected_sheet_type)]

                    if not y1_item_row.empty:
                        y1_item_id = y1_item_row.iloc[0]["Item_ID"]
                        loc_recs = df_rec[df_rec["Item_ID"] == y1_item_id]
                        merged = pd.merge(target_reps, loc_recs, on="Report_ID", how="inner").sort_values("Date_dt")
                        merged = merged.dropna(subset=["Value_Num"])

                        if not merged.empty:
                            if axis_mode == "2軸 (別スケール分析)":
                                fig.add_trace(go.Scatter(
                                    x=merged["Date"], y=merged["Value_Num"],
                                    mode="lines+markers", name=f"{loc_name} - {y1_item_name} (左軸)",
                                    line=dict(color=colors[c_idx % len(colors)], width=2.5)
                                ), secondary_y=False)
                            else:
                                fig.add_trace(go.Scatter(
                                    x=merged["Date"], y=merged["Value_Num"],
                                    mode="lines+markers", name=f"{loc_name} - {y1_item_name}",
                                    line=dict(color=colors[c_idx % len(colors)], width=2.5)
                                ))
                            c_idx += 1

                    if axis_mode == "2軸 (別スケール分析)" and y2_item_name:
                        y2_item_row = df_item[(df_item["Loc_ID"] == loc_id) & (df_item["Item_Name"] == y2_item_name) & (df_item["Sheet_Type"] == selected_sheet_type)]
                        if not y2_item_row.empty:
                            y2_item_id = y2_item_row.iloc[0]["Item_ID"]
                            loc_recs2 = df_rec[df_rec["Item_ID"] == y2_item_id]
                            merged2 = pd.merge(target_reps, loc_recs2, on="Report_ID", how="inner").sort_values("Date_dt")
                            merged2 = merged2.dropna(subset=["Value_Num"])

                            if not merged2.empty:
                                fig.add_trace(go.Scatter(
                                    x=merged2["Date"], y=merged2["Value_Num"],
                                    mode="lines+markers", name=f"{loc_name} - {y2_item_name} (右軸)",
                                    line=dict(color=colors[c_idx % len(colors)], width=2, dash="dash")
                                ), secondary_y=True)
                                c_idx += 1

                fig.update_layout(
                    title=f"📊 {site_options[selected_site_id]} 水質トレンド比較 ({start_d.strftime('%Y/%m/%d')} 〜 {end_d.strftime('%Y/%m/%d')})",
                    xaxis_title="点検日",
                    hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    plot_bgcolor="#F9FBF9",
                    paper_bgcolor="white",
                    height=520
                )
                fig.update_yaxes(title_text=y1_item_name)
                if axis_mode == "2軸 (別スケール分析)" and y2_item_name:
                    fig.update_yaxes(title_text=y2_item_name, secondary_y=True)

                st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------
# TAB 3: エクセル帳票一括ダウンロード (原本仕様2段ヘッダー出力)
# ------------------------------------------
with tab3:
    st.markdown("<div class='sub-header'>📑 全点検データの一括エクセル化 (原本仕様2段ヘッダー帳票出力)</div>", unsafe_allow_html=True)
    st.write("アップロードされた原本帳票と同じ「1行目: 槽名 / 2行目: 項目名」の2段ヘッダーおよび並び順で、一括エクセルファイルをダウンロードできます。")

    site_locs = df_loc[df_loc["Site_ID"] == selected_site_id].sort_values("Display_Order")
    loc_id_to_name = dict(zip(site_locs["Loc_ID"], site_locs["Loc_Name"]))
    items_ordered = df_item[(df_item["Loc_ID"].isin(site_locs["Loc_ID"])) & (df_item["Sheet_Type"] == selected_sheet_type)].sort_values("Display_Order")

    if items_ordered.empty:
        st.warning("該当する点検データが存在しません。")
    else:
        reports = df_rep[(df_rep["Site_ID"] == selected_site_id) & (df_rep["Sheet_Type"] == selected_sheet_type)].copy()
        reports["Date_dt"] = pd.to_datetime(reports["Date"], errors="coerce")
        reports = reports.sort_values("Date_dt").reset_index(drop=True)

        records = df_rec[df_rec["Report_ID"].isin(reports["Report_ID"])].copy()
        rec_dict = records.groupby(["Report_ID", "Item_ID"])["Value"].first().to_dict()

        col_defs = []
        for _, item_row in items_ordered.iterrows():
            l_name = loc_id_to_name.get(item_row["Loc_ID"], "")
            col_defs.append((item_row["Item_ID"], l_name, item_row["Item_Name"]))

        data_rows = []
        for _, rep_row in reports.iterrows():
            rep_id = rep_row["Report_ID"]
            r_date = rep_row["Date"]
            r_note = rep_row["Notes"] if pd.notna(rep_row["Notes"]) else ""
            row_vals = {"Date": r_date, "Notes": r_note}
            for item_id, l_name, i_name in col_defs:
                row_vals[item_id] = rec_dict.get((rep_id, item_id), "")
            data_rows.append(row_vals)

        df_raw = pd.DataFrame(data_rows)

        # 出力対象年の選択ドロップダウン (デフォルト: 全データ)
        available_years = sorted(reports["Date_dt"].dt.year.dropna().unique().astype(int).tolist(), reverse=True)
        year_options = ["全データ"] + [f"{y}年" for y in available_years]

        selected_year_opt = st.selectbox("📅 出力対象年を選択", options=year_options, index=0)

        if selected_year_opt == "全データ":
            df_export = df_raw.copy()
            file_year_str = "全データ"
        else:
            sel_y = int(selected_year_opt.replace("年", ""))
            df_export = df_raw[pd.to_datetime(df_raw["Date"], errors="coerce").dt.year == sel_y].copy()
            file_year_str = f"{sel_y}年"

        mi_cols = [("基本情報", "日付")] + [(loc_name, item_name) for _, loc_name, item_name in col_defs] + [("基本情報", "備考")]
        raw_col_keys = ["Date"] + [item_id for item_id, _, _ in col_defs] + ["Notes"]

        df_preview = df_export[raw_col_keys].copy()
        df_preview.columns = pd.MultiIndex.from_tuples(mi_cols)

        st.markdown(f"##### プレビュー ({selected_year_opt} - 2段ヘッダー構造)")
        st.dataframe(df_preview.tail(10) if len(df_preview) > 10 else df_preview, use_container_width=True)

        def generate_formatted_excel(selected_site_id, selected_sheet_type, site_name, df_data_source):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"{selected_sheet_type}_集計"
            ws.append([f"◆ {site_name} 【{selected_sheet_type}】 {file_year_str} 点検データ一覧表"])
            ws.append([f"出力日時: {datetime.datetime.now().strftime('%Y/%m/%d %H:%M')}"])
            ws.append([])

            header1 = ["日付"] + [loc_name for _, loc_name, _ in col_defs] + ["備考"]
            header2 = ["日付"] + [item_name for _, _, item_name in col_defs] + ["備考"]
            ws.append(header1)
            ws.append(header2)

            for _, r in df_data_source.iterrows():
                row_arr = [r["Date"]] + [r[item_id] for item_id, _, _ in col_defs] + [r["Notes"]]
                ws.append(row_arr)

            total_cols = len(header1)
            ws.merge_cells(start_row=4, start_column=1, end_row=5, end_column=1)
            ws.merge_cells(start_row=4, start_column=total_cols, end_row=5, end_column=total_cols)

            curr_start = 2
            for c_idx in range(2, total_cols):
                loc_curr = header1[c_idx - 1]
                loc_next = header1[c_idx] if c_idx < total_cols - 1 else None
                if loc_curr != loc_next:
                    if curr_start < c_idx:
                        ws.merge_cells(start_row=4, start_column=curr_start, end_row=4, end_column=c_idx)
                    curr_start = c_idx + 1

            fill_header1 = PatternFill(start_color="43A047", end_color="43A047", fill_type="solid")
            fill_header2 = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
            font_header1 = Font(name="メイリオ", size=11, bold=True, color="FFFFFF")
            font_header2 = Font(name="メイリオ", size=10, bold=True, color="2E7D32")
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin_border = Border(
                left=Side(style="thin", color="A5D6A7"),
                right=Side(style="thin", color="A5D6A7"),
                top=Side(style="thin", color="A5D6A7"),
                bottom=Side(style="thin", color="A5D6A7")
            )

            ws.cell(row=1, column=1).font = Font(name="メイリオ", size=14, bold=True, color="2E7D32")
            ws.cell(row=2, column=1).font = Font(name="メイリオ", size=10, italic=True, color="666666")

            for col in range(1, total_cols + 1):
                cell1 = ws.cell(row=4, column=col)
                cell2 = ws.cell(row=5, column=col)
                cell1.fill = fill_header1
                cell1.font = font_header1
                cell1.alignment = align_center
                cell1.border = thin_border
                cell2.fill = fill_header2
                cell2.font = font_header2
                cell2.alignment = align_center
                cell2.border = thin_border

            for row_idx in range(6, ws.max_row + 1):
                for col_idx in range(1, total_cols + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font = Font(name="メイリオ", size=10)
                    cell.border = thin_border
                    if col_idx == 1:
                        cell.alignment = Alignment(horizontal="center")
                    elif col_idx == total_cols:
                        cell.alignment = Alignment(horizontal="left")
                    else:
                        cell.alignment = Alignment(horizontal="right")

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row >= 4 and cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 5, 12)

            wb.save(output)
            output.seek(0)
            return output

        excel_data = generate_formatted_excel(selected_site_id, selected_sheet_type, site_options[selected_site_id], df_export)
        st.download_button(
            label=f"📥 【{site_options[selected_site_id]}】 原本仕様エクセル帳票を一括ダウンロード ({file_year_str}.xlsx)",
            data=excel_data,
            file_name=f"{site_options[selected_site_id]}_{selected_sheet_type}_{file_year_str}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ------------------------------------------
# TAB 4: データベース全件閲覧・簡易フィルタ
# ------------------------------------------
with tab4:
    st.markdown("<div class='sub-header'>🗃️ データベース全件閲覧 ＆ フィルタ検索</div>", unsafe_allow_html=True)
    st.write("システムの裏側で管理・蓄積されている Excel データベースの全テーブルを直接確認・照会できます。")

    db_view_tab1, db_view_tab2, db_view_tab3, db_view_tab4 = st.tabs([
        "📋 点検結果 (Inspection Records)",
        "📄 日報 master (Daily Report)",
        "📍 測定箇所 master (Location Master)",
        "🏷️ 測定項目 master (Item Master)"
    ])

    with db_view_tab1:
        st.markdown("##### 📋 点検結果レコード一覧 (Inspection Records)")
        st.dataframe(df_rec, use_container_width=True)

    with db_view_tab2:
        st.markdown("##### 📄 日報ヘッダー一覧 (Daily Report)")
        st.dataframe(df_rep, use_container_width=True)

    with db_view_tab3:
        st.markdown("##### 📍 測定箇所マスター (Location Master)")
        st.dataframe(df_loc, use_container_width=True)

    with db_view_tab4:
        st.markdown("##### 🏷️ 測定項目マスター (Item Master)")
        st.dataframe(df_item, use_container_width=True)
