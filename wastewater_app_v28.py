import os
import datetime
import io
import re
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
# 1.5 簡易パスワード認証保護 (Security)
# ==========================================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.markdown("<div class='main-header'>🔒 KBL 排水管理システム - ログイン</div>", unsafe_allow_html=True)
    st.markdown("#### 🔑 パスワード認証")
    pwd_input = st.text_input("パスワード", type="password", key="login_pwd_key")
    correct_pwd = st.secrets.get("APP_PASSWORD", "kbl2026")
    if st.button("ログイン"):
        if pwd_input == correct_pwd:
            st.session_state["authenticated"] = True
            st.rerun()
        elif pwd_input:
            st.error("❌ パスワードが正しくありません。")
    st.stop()


# ==========================================
# GitHub REST API による自動コミット・同期関数
# ==========================================
def sync_to_github_api(file_path):
    """Syncs local file to GitHub repository using GitHub REST API and st.secrets["GITHUB_TOKEN"]"""
    try:
        if "GITHUB_TOKEN" not in st.secrets:
            return False, "GITHUB_TOKEN 未設定 (st.secrets)"
        
        token = st.secrets["GITHUB_TOKEN"]
        repo = st.secrets.get("GITHUB_REPO", "kbl-wastewater-app")
        branch = st.secrets.get("GITHUB_BRANCH", "main")
        
        target_path = os.path.basename(file_path)
        if not target_path.endswith('.xlsx'):
            target_path = "wastewater-appsheet-db-v3.xlsx"
            
        import base64
        import json
        import urllib.request
        import urllib.error
        
        url = f"https://api.github.com/repos/{repo}/contents/{target_path}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "StreamlitApp"
        }
        
        # 1. Fetch current SHA
        sha = None
        req_get = urllib.request.Request(f"{url}?ref={branch}", headers=headers, method="GET")
        try:
            with urllib.request.urlopen(req_get) as resp:
                res_data = json.loads(resp.read().decode("utf-8"))
                sha = res_data.get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                return False, f"HTTP Error {e.code} (SHA取得失敗)"
                
        # 2. Read local file and encode base64
        with open(file_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode("utf-8")
            
        now_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        payload = {
            "message": f"Auto-update wastewater data via App [{now_str}]",
            "content": content_b64,
            "branch": branch
        }
        if sha:
            payload["sha"] = sha
            
        data_json = json.dumps(payload).encode("utf-8")
        req_put = urllib.request.Request(url, data=data_json, headers=headers, method="PUT")
        
        with urllib.request.urlopen(req_put) as resp:
            if resp.status in [200, 201]:
                return True, "成功"
            return False, f"HTTP Status {resp.status}"
    except Exception as ex:
        return False, str(ex)


# ==========================================
# 自動計算関数 (有機割合、無機割合、沈降速度、水面積負荷、返送率など)
# ==========================================
def calculate_auto_metrics(df_item_all, df_loc_all, selected_site_id, selected_sheet_type, form_data_dict, df_rec_latest, target_rep_id):
    calc_updates = {}
    
    def _parse_val(item_id):
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

    def _find_items(df_subset, name_query):
        exact = df_subset[df_subset['Item_Name'] == name_query]
        if not exact.empty:
            return exact
        prefix = df_subset[df_subset['Item_Name'].astype(str).str.startswith(name_query)]
        if not prefix.empty:
            return prefix
        contains = df_subset[df_subset['Item_Name'].astype(str).str.contains(name_query, regex=False)]
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
            mlss_v = _parse_val(mlss_item.iloc[0]['Item_ID'])
            mlvss_v = _parse_val(mlvss_item.iloc[0]['Item_ID'])
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
        
        isou_v = _parse_val(isou_item.iloc[0]['Item_ID']) if not isou_item.empty else None
        hensou_v = _parse_val(hensou_item.iloc[0]['Item_ID']) if not hensou_item.empty else None
        mlss_v = _parse_val(mlss_item.iloc[0]['Item_ID']) if not mlss_item.empty else None
        temp_v = _parse_val(temp_item.iloc[0]['Item_ID']) if not temp_item.empty else None
        sv30_v = _parse_val(sv30_item.iloc[0]['Item_ID']) if not sv30_item.empty else None
        
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
                
    return calc_updates


# DBファイルパスの取得
def get_db_path():
    db_v3 = "wastewater-appsheet-db-v3.xlsx"
    db_v2 = "wastewater-appsheet-db-v2.xlsx"
    db_v1 = "wastewater-appsheet-db.xlsx"
    if os.path.exists(db_v3):
        return db_v3
    elif os.path.exists(db_v2):
        return db_v2
    elif os.path.exists(db_v1):
        return db_v1
    else:
        v3_abs = "/workspace/artifacts/wastewater-appsheet-db-v3.xlsx"
        v2_abs = "/workspace/artifacts/wastewater-appsheet-db-v2.xlsx"
        if os.path.exists(v3_abs):
            return v3_abs
        if os.path.exists(v2_abs):
            return v2_abs
        return "/workspace/artifacts/wastewater-appsheet-db.xlsx"

db_path = get_db_path()

# ==========================================
# 2. データ読み込み ＆ キャッシュ処理
# ==========================================
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
    
    df_site = df_site[~df_site["Site_Name"].astype(str).str.contains("北越") & (df_site["Site_ID"] != "S003")]
    df_loc = df_loc[~df_loc["Site_ID"].isin(["S003"]) & df_loc["Site_ID"].isin(df_site["Site_ID"])]
    df_item = df_item[df_item["Loc_ID"].isin(df_loc["Loc_ID"])]
    df_rep = df_rep[df_rep["Site_ID"].isin(df_site["Site_ID"])]
    df_rec = df_rec[df_rec["Report_ID"].isin(df_rep["Report_ID"])]

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

df_site, df_loc, df_item, df_rep, df_rec = load_all_data(db_path)

if df_site is None:
    st.stop()

# 日本時間 (JST) 当日日付取得
def get_jst_today():
    return (datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=9)).date()

# タイトル表示
st.markdown("<div class='main-header'>🌱 排水処理点検・水質データ管理システム (KBL Management App)</div>", unsafe_allow_html=True)

# サイドバー：グローバル選択ヘッダー
st.sidebar.image("https://img.icons8.com/color/96/000000/sprout.png", width=64)
st.sidebar.title("📌 対象設定")

site_options = dict(zip(df_site["Site_ID"], df_site["Site_Name"]))
selected_site_id = st.sidebar.selectbox(
    "① 現場を選択",
    options=list(site_options.keys()),
    format_func=lambda x: site_options[x]
)

available_sheet_types = ["点検管理表", "計量証明"]

selected_sheet_type = st.sidebar.radio(
    "② シート種別を選択",
    options=available_sheet_types,
    horizontal=True
)

st.sidebar.markdown("---")
st.sidebar.info(f"**選択中の現場**: {site_options[selected_site_id]}\n\n**シート**: {selected_sheet_type}")

# ==========================================
# 3. メイン画面タブ構成
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📝 ①〜④ 点検データ入力",
    "📊 ⑤ 動的2軸・複数槽比較水質グラフ",
    "📑 ⑥ エクセル帳票ダウンロード",
    "🗃️ データベース全件閲覧"
])

# ------------------------------------------
# TAB 1: 点検データ入力
# ------------------------------------------
with tab1:
    # 保存成功メッセージの維持表示
    if "save_success_msg" in st.session_state:
        st.success(st.session_state["save_success_msg"])
        del st.session_state["save_success_msg"]
        
    st.markdown("<div class='sub-header'>📝 点検結果の新規入力 ＆ 蓄積 (過去データとの異常値チェック機能付き)</div>", unsafe_allow_html=True)

    col_input1, col_input2 = st.columns(2)

    with col_input1:
        input_date = st.date_input("③ 点検日を選択", value=get_jst_today(), key="input_date_key")
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

            stats_by_item = {}
            rep_site_match = df_rep[(df_rep["Site_ID"] == selected_site_id) & (df_rep["Sheet_Type"] == selected_sheet_type)]
            rec_site_match = df_rec[df_rec["Report_ID"].isin(rep_site_match["Report_ID"])]
            
            for item_id_key in loc_items["Item_ID"].unique():
                item_recs = rec_site_match[rec_site_match["Item_ID"] == item_id_key]["Value_Num"].dropna()
                vals = item_recs.values
                if len(vals) >= 2:
                    mean_v = float(np.mean(vals))
                    std_v = float(np.std(vals))
                    min_v = float(np.min(vals))
                    max_v = float(np.max(vals))
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

            st.info("💡 **自動計算項目（有機割合・無機割合・水面積負荷・沈降速度・返送率など）** は、入力された基礎数値から保存時に**全自動計算されExcelデータベースへ直接格納**されます（入力ボックスは非表示としています）。")
            
            calc_item_names = ['有機割合', '有機割合 (%)', '無機割合', '沈降速度', '水面積負荷', '沈降速度/水面積負荷', '返送率', '返送率 (%)']
            input_items = loc_items[~loc_items['Item_Name'].isin(calc_item_names)].sort_values("Display_Order")

            # スマホ用テンキー（inputmode="decimal"）自動適用 JavaScript
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
                                form_data[item_id] = st.text_input(label_str, value=str(default_val) if str(default_val) != "nan" else "")
                                if item_id in stats_by_item:
                                    st_info = stats_by_item[item_id]
                                    st.caption(f"💡 過去平均: {st_info['mean']:.2f} (目安: {st_info['lower']:.1f} 〜 {st_info['upper']:.1f})")

                notes_input = st.text_area("備考 (Notes)", value=existing_rep.iloc[0]["Notes"] if not existing_rep.empty and pd.notna(existing_rep.iloc[0]["Notes"]) else "")

                submit_btn = st.form_submit_button("💾 点検データを保存・蓄積する")

                if submit_btn:
                    anomalies = []
                    for item_id, val in form_data.items():
                        val_str = str(val).strip().replace(",", "").replace("%", "").replace("<", "").replace(">", "")
                        if val_str and val_str not in ["-", "nan"]:
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
                        st.warning("⚠️ **【異常値・誤入力の可能性あり】**\n" + "\n".join(anomalies) + "\n\n※数値に問題が無ければ、データは正常に保存されます。")

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
                            m = re.search(r"\d+", str(rid))
                            if m:
                                max_rep_num = max(max_rep_num, int(m.group()))
                        target_rep_id = f"REP{max_rep_num+1:05d}"
                        ws_rep.append([target_rep_id, input_date_str, selected_site_id, selected_sheet_type, "", notes_input])

                    existing_item_ids = set(df_rec_latest[df_rec_latest["Report_ID"] == target_rep_id]["Item_ID"])

                    max_rec_num = 0
                    for rcid in df_rec_latest["Rec_ID"].dropna():
                        m = re.search(r"\d+", str(rcid))
                        if m:
                            max_rec_num = max(max_rec_num, int(m.group()))

                    auto_calcs = calculate_auto_metrics(df_item, df_loc, selected_site_id, selected_sheet_type, form_data, df_rec_latest, target_rep_id)
                    final_save_data = form_data.copy()
                    final_save_data.update(auto_calcs)

                    for item_id, val in final_save_data.items():
                        if str(val).strip() in ["", "-", "nan"]:
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

                    # GitHub API 同期処理
                    sync_ok, sync_msg = sync_to_github_api(db_path)
                    if sync_ok:
                        st.session_state["save_success_msg"] = f"✅ {input_date_str} 『{loc_options[selected_loc_id]}』 の点検データを正常に保存しました（GitHubへの自動同期も成功しました）！"
                    else:
                        st.session_state["save_success_msg"] = f"✅ {input_date_str} 『{loc_options[selected_loc_id]}』 の点検データを正常に保存しました（ローカル更新完了 / GitHub同期: {sync_msg}）"

                    st.rerun()

# ------------------------------------------
# TAB 2: 動的2軸・複数槽比較水質グラフ
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
            selected_loc_ids = st.multiselect(
                "① 比較する槽を選択 (最大3箇所)",
                options=all_loc_ids,
                default=[all_loc_ids[0]] if all_loc_ids else [],
                format_func=lambda x: graph_loc_options[x],
                max_selections=3
            )

        with col_g2:
            if selected_loc_ids:
                avail_items_primary = df_item[
                    (df_item["Loc_ID"].isin(selected_loc_ids)) & 
                    (df_item["Sheet_Type"] == selected_sheet_type) &
                    (df_item["Value_Type"] == "Number")
                ]["Item_Name"].unique()
                
                selected_primary_item = st.selectbox("② 主軸 (左Y軸) 項目", options=avail_items_primary if len(avail_items_primary) > 0 else ["データなし"])
            else:
                selected_primary_item = None

        with col_g3:
            if selected_loc_ids:
                avail_items_secondary = ["(なし)"] + list(avail_items_primary) if selected_primary_item else ["(なし)"]
                selected_secondary_item = st.selectbox("③ 副軸 (右Y軸) 項目 [任意]", options=avail_items_secondary)
            else:
                selected_secondary_item = "(なし)"

        rep_site = df_rep[(df_rep["Site_ID"] == selected_site_id) & (df_rep["Sheet_Type"] == selected_sheet_type)].sort_values("Date")
        
        if not rep_site.empty:
            dates = pd.to_datetime(rep_site["Date"]).sort_values()
            min_d = dates.min().date()
            max_d = dates.max().date()
            
            st.markdown("##### 📅 期間絞り込み")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                start_d = st.date_input("開始日", value=min_d, min_value=min_d, max_value=max_d)
            with col_d2:
                end_d = st.date_input("終了日", value=max_d, min_value=min_d, max_value=max_d)
        else:
            start_d, end_d = None, None

        if selected_loc_ids and selected_primary_item and selected_primary_item != "データなし":
            filtered_rep = rep_site[
                (pd.to_datetime(rep_site["Date"]).dt.date >= start_d) & 
                (pd.to_datetime(rep_site["Date"]).dt.date <= end_d)
            ]
            
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            colors = ["#2E7D32", "#1976D2", "#E65100", "#D32F2F", "#7B1FA2"]
            
            for i, loc_id in enumerate(selected_loc_ids):
                loc_name = graph_loc_options[loc_id]
                color = colors[i % len(colors)]
                
                item_p = df_item[(df_item["Loc_ID"] == loc_id) & (df_item["Sheet_Type"] == selected_sheet_type) & (df_item["Item_Name"] == selected_primary_item)]
                if not item_p.empty:
                    item_p_id = item_p.iloc[0]["Item_ID"]
                    rec_p = df_rec[(df_rec["Report_ID"].isin(filtered_rep["Report_ID"])) & (df_rec["Item_ID"] == item_p_id)]
                    merged_p = filtered_rep.merge(rec_p, on="Report_ID").sort_values("Date")
                    
                    fig.add_trace(
                        go.Scatter(
                            x=merged_p["Date"],
                            y=merged_p["Value_Num"],
                            mode="lines+markers",
                            name=f"{loc_name} - {selected_primary_item}",
                            line=dict(color=color, width=2.5),
                            marker=dict(size=6)
                        ),
                        secondary_y=False
                    )
                
                if selected_secondary_item and selected_secondary_item != "(なし)":
                    item_s = df_item[(df_item["Loc_ID"] == loc_id) & (df_item["Sheet_Type"] == selected_sheet_type) & (df_item["Item_Name"] == selected_secondary_item)]
                    if not item_s.empty:
                        item_s_id = item_s.iloc[0]["Item_ID"]
                        rec_s = df_rec[(df_rec["Report_ID"].isin(filtered_rep["Report_ID"])) & (df_rec["Item_ID"] == item_s_id)]
                        merged_s = filtered_rep.merge(rec_s, on="Report_ID").sort_values("Date")
                        
                        fig.add_trace(
                            go.Scatter(
                                x=merged_s["Date"],
                                y=merged_s["Value_Num"],
                                mode="lines+markers",
                                name=f"{loc_name} - {selected_secondary_item}",
                                line=dict(color=color, width=2, dash="dash"),
                                marker=dict(size=5, symbol="diamond")
                            ),
                            secondary_y=True
                        )

            title_str = f"【{site_options[selected_site_id]}】 水質推移グラフ ({selected_sheet_type})"
            fig.update_layout(
                title=dict(text=title_str, font=dict(size=18, color="#2E7D32")),
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(l=40, r=40, t=60, b=40),
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#F9FBF7"
            )
            fig.update_xaxes(title_text="測定日", showgrid=True, gridcolor="#E8F5E9")
            fig.update_yaxes(title_text=selected_primary_item, secondary_y=False, showgrid=True, gridcolor="#E8F5E9")
            if selected_secondary_item and selected_secondary_item != "(なし)":
                fig.update_yaxes(title_text=selected_secondary_item, secondary_y=True, showgrid=False)

            st.plotly_chart(fig, use_container_width=True)
            st.caption("📸 グラフ右上のカメラアイコンをタップすると、グラフをPNG画像としてワンクリック保存できます。")

# ------------------------------------------
# TAB 3: エクセル帳票ダウンロード (要件⑥)
# ------------------------------------------
with tab3:
    st.markdown("<div class='sub-header'>📑 現場別・月別/年別 エクセル帳票ダウンロード</div>", unsafe_allow_html=True)
    st.info("💡 選択した現場・年度の全水質測定データを、公式様式のExcelファイル (.xlsx) として一括生成・出力します。")

    rep_site_all = df_rep[(df_rep["Site_ID"] == selected_site_id) & (df_rep["Sheet_Type"] == selected_sheet_type)]
    
    if rep_site_all.empty:
        st.warning("該当現場のデータが存在しません。")
    else:
        years = sorted(pd.to_datetime(rep_site_all["Date"]).dt.year.unique(), reverse=True)
        selected_year = st.selectbox("📅 出力対象年を選択", options=years)

        def generate_formatted_excel(site_id, sheet_type, year):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"{year}年_{sheet_type}"

            site_name = site_options[site_id]
            ws.append([f"◆ {site_name} 【{sheet_type}】 {year}年 水質管理月報・年間データ一覧表"])
            ws.append([f"出力日時: {get_jst_today().strftime('%Y/%m/%d')} (KBL Management App System)"])
            ws.append([])

            rep_year = rep_site_all[pd.to_datetime(rep_site_all["Date"]).dt.year == year].sort_values("Date")
            rec_year = df_rec[df_rec["Report_ID"].isin(rep_year["Report_ID"])]

            locs_in_sheet = df_loc[(df_loc["Site_ID"] == site_id) & (df_loc["Loc_ID"].isin(df_item[df_item["Sheet_Type"] == sheet_type]["Loc_ID"]))].sort_values("Display_Order")

            header_loc = ["日付"]
            header_item = [""]

            item_col_map = {}
            col_idx = 2

            for _, loc in locs_in_sheet.iterrows():
                l_items = df_item[(df_item["Loc_ID"] == loc["Loc_ID"]) & (df_item["Sheet_Type"] == sheet_type)].sort_values("Display_Order")
                for _, it in l_items.iterrows():
                    header_loc.append(loc["Loc_Name"])
                    unit_str = f"({it['Unit_or_Options']})" if pd.notna(it["Unit_or_Options"]) and str(it["Unit_or_Options"]) != "nan" else ""
                    header_item.append(f"{it['Item_Name']}{unit_str}")
                    item_col_map[it["Item_ID"]] = col_idx
                    col_idx += 1

            header_loc.append("備考")
            header_item.append("特記事項")

            ws.append(header_loc)
            ws.append(header_item)

            for _, rep in rep_year.iterrows():
                r_date = rep["Date"]
                r_id = rep["Report_ID"]
                r_notes = rep["Notes"] if pd.notna(rep["Notes"]) else ""

                row_data = [""] * col_idx
                row_data[0] = r_date
                row_data[-1] = r_notes

                recs = rec_year[rec_year["Report_ID"] == r_id]
                for _, rec in recs.iterrows():
                    it_id = rec["Item_ID"]
                    val = rec["Value"]
                    if it_id in item_col_map:
                        row_data[item_col_map[it_id] - 1] = val

                ws.append(row_data)

            header_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
            sub_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

            for col_i in range(1, col_idx + 1):
                ws.cell(row=4, column=col_i).fill = header_fill
                ws.cell(row=4, column=col_i).font = Font(bold=True, color="1B5E20")
                ws.cell(row=5, column=col_i).fill = sub_fill
                ws.cell(row=5, column=col_i).font = Font(bold=True)

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            wb.save(output)
            return output.getvalue()

        excel_data = generate_formatted_excel(selected_site_id, selected_sheet_type, selected_year)
        file_name = f"{selected_year}_{site_options[selected_site_id]}_{selected_sheet_type}.xlsx"

        st.download_button(
            label=f"📥 【{selected_year}年】 {site_options[selected_site_id]} エクセル帳票をダウンロード",
            data=excel_data,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ------------------------------------------
# TAB 4: データベース全件閲覧
# ------------------------------------------
with tab4:
    st.markdown("<div class='sub-header'>🗃️ データベース全件閲覧 (Master & Daily Logs)</div>", unsafe_allow_html=True)
    st.info("💡 データベースに登録されている全テーブルの内容を閲覧できます。")

    sub_tab1, sub_tab2, sub_tab3, sub_tab4, sub_tab5 = st.tabs([
        "🏢 Site Master",
        "📍 Location Master",
        "🏷️ Item Master",
        "📅 Daily Report",
        "📝 Inspection Records"
    ])

    with sub_tab1:
        st.dataframe(df_site, use_container_width=True)

    with sub_tab2:
        st.dataframe(df_loc[df_loc["Site_ID"] == selected_site_id], use_container_width=True)

    with sub_tab3:
        st.dataframe(df_item[df_item["Loc_ID"].isin(df_loc[df_loc["Site_ID"] == selected_site_id]["Loc_ID"])], use_container_width=True)

    with sub_tab4:
        st.dataframe(df_rep[df_rep["Site_ID"] == selected_site_id].sort_values("Date", ascending=False), use_container_width=True)

    with sub_tab5:
        rep_ids = df_rep[df_rep["Site_ID"] == selected_site_id]["Report_ID"]
        st.dataframe(df_rec[df_rec["Report_ID"].isin(rep_ids)], use_container_width=True)