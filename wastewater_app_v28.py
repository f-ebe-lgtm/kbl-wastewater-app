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
    st.write("関係者専用のシステムです。パスワードを入力してログインしてください。")
    
    pwd_input = st.text_input("パスワード", type="password", key="login_pwd_key")
    if st.button("🔓 ログイン") or pwd_input:
        try:
            correct_pwd = st.secrets["PASSWORD"]
        except Exception:
            correct_pwd = "kbl2026"
            
        if pwd_input == correct_pwd:
            st.session_state["authenticated"] = True
            st.rerun()
        elif pwd_input:
            st.error("❌ パスワードが正しくありません。")
    st.stop()



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

    def find_item(df_scope, name_prefix):
        if df_scope.empty:
            return df_scope
        match = df_scope[df_scope['Item_Name'] == name_prefix]
        if match.empty:
            match = df_scope[df_scope['Item_Name'].str.startswith(name_prefix, na=False)]
        return match

    # 1. 有機割合 [%] (MLVSS/MLSS*100) & 無機割合 [%] (100 - 有機割合)
    locs_in_site = df_loc_all[df_loc_all['Site_ID'] == selected_site_id]['Loc_ID'].unique()
    for loc_id in locs_in_site:
        loc_items = df_item_all[(df_item_all['Loc_ID'] == loc_id) & (df_item_all['Sheet_Type'] == selected_sheet_type)]
        mlss_item = find_item(loc_items, 'MLSS')
        mlvss_item = find_item(loc_items, 'MLVSS')
        org_item = find_item(loc_items, '有機割合')
        inorg_item = find_item(loc_items, '無機割合')
        
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
        
        isou_item = find_item(site_items, '移送量')
        hensou_item = find_item(site_items, '返送量')
        
        aeration_end = site_items[site_items['Loc_Name'] == '曝気槽（末端側）']
        mlss_item = find_item(aeration_end, 'MLSS')
        temp_item = find_item(aeration_end, '水温')
        sv30_item = find_item(aeration_end, 'SV30')
        
        isou_v = _parse_val(isou_item.iloc[0]['Item_ID']) if not isou_item.empty else None
        hensou_v = _parse_val(hensou_item.iloc[0]['Item_ID']) if not hensou_item.empty else None
        mlss_v = _parse_val(mlss_item.iloc[0]['Item_ID']) if not mlss_item.empty else None
        temp_v = _parse_val(temp_item.iloc[0]['Item_ID']) if not temp_item.empty else None
        sv30_v = _parse_val(sv30_item.iloc[0]['Item_ID']) if not sv30_item.empty else None
        
        load_item = find_item(site_items, '水面積負荷')
        return_ratio_item = find_item(site_items, '返送率')
        settling_item = site_items[site_items['Item_Name'] == '沈降速度']
        ratio_item = site_items[site_items['Item_Name'] == '沈降速度/水面積負荷']
        
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
    if os.path.exists(DB_FILENAME_V2):
        return DB_FILENAME_V2
    elif os.path.exists(DB_FILENAME_V1):
        return DB_FILENAME_V1
    else:
        v2_abs = "/workspace/artifacts/wastewater-appsheet-db-v2.xlsx"
        v1_abs = "/workspace/artifacts/wastewater-appsheet-db.xlsx"
        if os.path.exists(v2_abs):
            return v2_abs
        return v1_abs

db_path = get_db_path()

# ==========================================
# 2. データ読み込み ＆ キャッシュ処理 (パーセント・不等号クレンジング強化)
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

df_site, df_loc, df_item, df_rep, df_rec = load_all_data(db_path)

if df_site is None:
    st.stop()

# タイトル表示

# スマホ入力用：テンキー（数字キーボード）自動呼び出し機能
components.html("""
<script>
function setNumericInputMode() {
    try {
        const pDoc = window.parent.document;
        const inputs = pDoc.querySelectorAll('input[type="text"]');
        inputs.forEach(input => {
            if (!input.id.includes('login') && !input.getAttribute('inputmode')) {
                input.setAttribute('inputmode', 'decimal');
            }
        });
    } catch(e) {}
}
setNumericInputMode();
const observer = new MutationObserver(setNumericInputMode);
observer.observe(window.parent.document.body, { childList: true, subtree: true });
</script>
""", height=0, width=0)

st.markdown("<div class='main-header'>🌱 排水処理点検・水質データ管理システム (KBL Management App)</div>", unsafe_allow_html=True)

# サイドバー：グローバル選択ヘッダー
st.sidebar.image("https://img.icons8.com/color/96/000000/sprout.png", width=64)
st.sidebar.title("📌 対象設定")

# ① 現場の選択
site_options = dict(zip(df_site["Site_ID"], df_site["Site_Name"]))
selected_site_id = st.sidebar.selectbox(
    "① 現場を選択",
    options=list(site_options.keys()),
    format_func=lambda x: site_options[x]
)

# ② シートの選択（点検管理表 or 計量証明）
available_sheet_types = ["点検管理表", "計量証明"]

selected_sheet_type = st.sidebar.radio(
    "② シート種別を選択",
    options=available_sheet_types,
    horizontal=True
)

st.sidebar.markdown("---")
st.sidebar.info(f"**選択中の現場**: {site_options[selected_site_id]}\n\n**シート**: {selected_sheet_type}")

# ==========================================
# 3. メイン画面タブ構成 (要件①〜⑥)
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📝 ①〜④ 点検データ入力", 
    "📊 ⑤ 動的2軸・複数槽比較水質グラフ", 
    "📑 ⑥ エクセル帳票ダウンロード",
    "🗃️ データベース全件閲覧"
])

# ------------------------------------------
# TAB 1: 点検データ入力 (要件①〜④) + 異常値チェック
# ------------------------------------------
with tab1:
    st.markdown("<div class='sub-header'>📝 点検結果の新規入力 ＆ 蓄積 (過去データとの異常値チェック機能付き)</div>", unsafe_allow_html=True)
    
    col_input1, col_input2 = st.columns(2)
    
    with col_input1:
        # ③ 日付の選択 (原則当日、変更可)
        input_date = st.date_input("③ 点検日を選択", datetime.date.today())
        input_date_str = input_date.strftime("%Y/%m/%d")
        
    with col_input2:
        # 該当現場・シート種別に適した Location フィルタ
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
        
        # 選択された Loc & Sheet_Type に対応する Item を抽出
        loc_items = df_item[(df_item["Loc_ID"] == selected_loc_id) & (df_item["Sheet_Type"] == selected_sheet_type)].sort_values("Display_Order")
        
        if loc_items.empty:
            st.warning("指定された槽・シート種別の点検項目が登録されていません。")
        else:
            # 既存データのプレロード判定
            existing_rep = df_rep[(df_rep["Site_ID"] == selected_site_id) & (df_rep["Date"] == input_date_str) & (df_rep["Sheet_Type"] == selected_sheet_type)]
            existing_values = {}
            if not existing_rep.empty:
                rep_id_exist = existing_rep.iloc[0]["Report_ID"]
                rec_exist = df_rec[df_rec["Report_ID"] == rep_id_exist]
                existing_values = dict(zip(rec_exist["Item_ID"], rec_exist["Value"]))
                st.info(f"💡 {input_date_str} の既存データが読み込まれました。必要に応じて内容を更新してください。")
                
            # 過去データの統計情報算出 (異常値チェック用)
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

            st.info("💡 **自動計算項目（有機割合・無機割合・水面積負荷・沈降速度・返送率など）** は、入力された基礎数値から保存時に**全自動計算されExcelデータベースへ直接格納**されます（入力ボックスは非表示としています）。")
            
            # 自動計算項目を入力フォームから除外
            calc_item_names = ['有機割合', '有機割合 (%)', '無機割合', '沈降速度', '水面積負荷', '沈降速度/水面積負荷', '返送率', '返送率 (%)']
            input_items = loc_items[~loc_items['Item_Name'].isin(calc_item_names)].sort_values("Display_Order")

            with st.form("inspection_input_form"):
                form_data = {}
                cols = st.columns(3)
                
                for idx, (_, item) in enumerate(input_items.iterrows()):
                    item_id = item["Item_ID"]
                    item_name = item["Item_Name"]
                    vtype = item["Value_Type"]
                    unit = str(item["Unit_or_Options"]) if pd.notna(item["Unit_or_Options"]) else ""
                    
                    default_val = existing_values.get(item_id, "")
                    
                    label_str = item_name if vtype == "Enum" or not unit or unit == "nan" else f"{item_name} ({unit})"
                    col_target = cols[idx % 3]
                    
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
                    # 異常値チェック
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

                    # Excelへの書き込み処理
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
                            
                    # 自動計算処理の実行 (有機割合, 無機割合, 水面積負荷, 沈降速度, 沈降速度/水面積負荷, 返送率)
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
                    st.success(f"✅ {input_date_str} 『{loc_options[selected_loc_id]}』 の点検データを正常に保存しました！")
                    st.rerun()

# ------------------------------------------
# TAB 2: 動的2軸・複数槽比較水質グラフ (要件⑤)
# ------------------------------------------
with tab2:
    st.markdown("<div class='sub-header'>📊 期間・複数槽(最大3箇所)・1軸/2軸の比較水質グラフ</div>", unsafe_allow_html=True)
    
    # 該当現場・シート種別で使用可能な全 Location を取得
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
            # 槽の選択 (最大3つまで複数選択可能)
            selected_loc_ids = st.multiselect(
                "📍 比較する槽（測定箇所）を選択 (最大3つ)",
                options=all_loc_ids,
                default=all_loc_ids[:min(3, len(all_loc_ids))],
                format_func=lambda x: graph_loc_options[x],
                max_selections=3,
                key="g_multilocs"
            )
            
        if not selected_loc_ids:
            st.info("比較する槽を1つ以上選択してください。")
        else:
            # 選択された全槽に属する項目を抽出
            sel_items = df_item[(
                df_item["Loc_ID"].isin(selected_loc_ids)
            ) & (df_item["Sheet_Type"] == selected_sheet_type)].sort_values("Display_Order")
            
            # ユニークな項目名リストを作成 (単位付きラベル)
            unique_item_names = []
            item_name_to_unit = {}
            for _, r in sel_items.iterrows():
                iname = r["Item_Name"]
                uopt = str(r["Unit_or_Options"]) if pd.notna(r["Unit_or_Options"]) else ""
                if iname not in unique_item_names:
                    unique_item_names.append(iname)
                if iname in ["無機割合", "有機割合", "返送率"]:
                    item_name_to_unit[iname] = " (%)"
                elif uopt and uopt != "nan":
                    item_name_to_unit[iname] = f" ({uopt})"
                elif iname not in item_name_to_unit:
                    item_name_to_unit[iname] = ""
                    
            item_display_map = {iname: f"{iname}{item_name_to_unit.get(iname, '')}" for iname in unique_item_names}
            
            with col_g2:
                selected_item_y1 = st.selectbox(
                    "📈 第1縦軸 (左軸/実線・●)",
                    options=unique_item_names,
                    format_func=lambda x: item_display_map[x],
                    key="g_y1_name"
                )
                
            with col_g3:
                y2_options = [None] + unique_item_names
                selected_item_y2 = st.selectbox(
                    "📉 第2縦軸 (右軸/実線・◆ - オプション)",
                    options=y2_options,
                    format_func=lambda x: "なし (1軸のみ)" if x is None else item_display_map[x],
                    key="g_y2_name"
                )

            st.markdown("---")
            st.markdown("##### 📅 グラフ表示期間の設定")

            site_reps = df_rep[(df_rep["Site_ID"] == selected_site_id) & (df_rep["Sheet_Type"] == selected_sheet_type)]
            if not site_reps.empty:
                site_dates_dt = pd.to_datetime(site_reps["Date"]).sort_values().drop_duplicates()
                all_date_strs = site_dates_dt.dt.strftime("%Y/%m/%d").tolist()
                min_d = site_dates_dt.min().date()
                max_d = site_dates_dt.max().date()
            else:
                all_date_strs = [datetime.date.today().strftime("%Y/%m/%d")]
                min_d = datetime.date(2023, 1, 1)
                max_d = datetime.date.today()

            # 状態更新用コールバック関数 (最も近い実点検日へ自動アラインメント)
            def update_range_state(target_s_d, target_e_d):
                target_s_dt = pd.to_datetime(target_s_d)
                target_e_dt = pd.to_datetime(target_e_d)
                
                closest_s_str = min(all_date_strs, key=lambda x: abs((pd.to_datetime(x) - target_s_dt).days))
                closest_e_str = min(all_date_strs, key=lambda x: abs((pd.to_datetime(x) - target_e_dt).days))
                
                st.session_state["g_start_d"] = pd.to_datetime(closest_s_str).date()
                st.session_state["g_end_d"] = pd.to_datetime(closest_e_str).date()
                st.session_state["sel_start_str"] = closest_s_str
                st.session_state["sel_end_str"] = closest_e_str

            # 初期セッション状態の設定
            if "g_start_d" not in st.session_state or st.session_state.get("g_last_site") != selected_site_id or st.session_state.get("g_last_sheet") != selected_sheet_type:
                update_range_state(min_d, max_d)
                st.session_state["g_last_site"] = selected_site_id
                st.session_state["g_last_sheet"] = selected_sheet_type

            # 1. プリセットボタン ＆ リセットボタン群
            btn_c1, btn_c2, btn_c3, btn_c4, btn_c5 = st.columns([1.5, 1, 1, 1, 1])
            with btn_c1:
                if st.button("🔄 期間リセット (全期間)", key="btn_reset"):
                    update_range_state(min_d, max_d)
                    st.rerun()
                    
            with btn_c2:
                if st.button("過去3ヶ月", key="btn_3m"):
                    s_3m = max(min_d, (pd.to_datetime(max_d) - pd.DateOffset(months=3)).date())
                    update_range_state(s_3m, max_d)
                    st.rerun()
                    
            with btn_c3:
                if st.button("過去6ヶ月", key="btn_6m"):
                    s_6m = max(min_d, (pd.to_datetime(max_d) - pd.DateOffset(months=6)).date())
                    update_range_state(s_6m, max_d)
                    st.rerun()
                    
            with btn_c4:
                if st.button("過去1年", key="btn_1y"):
                    s_1y = max(min_d, (pd.to_datetime(max_d) - pd.DateOffset(years=1)).date())
                    update_range_state(s_1y, max_d)
                    st.rerun()
                    
            with btn_c5:
                if st.button("過去2年", key="btn_2y"):
                    s_2y = max(min_d, (pd.to_datetime(max_d) - pd.DateOffset(years=2)).date())
                    update_range_state(s_2y, max_d)
                    st.rerun()

            # 2. 開始点検日・終了点検日のドロップダウン選択
            d_col1, d_col2 = st.columns(2)
            
            curr_start_str = st.session_state["g_start_d"].strftime("%Y/%m/%d") if isinstance(st.session_state["g_start_d"], (datetime.date, pd.Timestamp)) else all_date_strs
            curr_start_idx = all_date_strs.index(curr_start_str) if curr_start_str in all_date_strs else 0
            
            with d_col1:
                sel_start_str = st.selectbox("📅 開始点検日を選択", options=all_date_strs, index=curr_start_idx, key="sel_start_str")
                st.session_state["g_start_d"] = pd.to_datetime(sel_start_str).date()

            curr_end_str = st.session_state["g_end_d"].strftime("%Y/%m/%d") if isinstance(st.session_state["g_end_d"], (datetime.date, pd.Timestamp)) else all_date_strs[-1]
            curr_end_idx = all_date_strs.index(curr_end_str) if curr_end_str in all_date_strs else len(all_date_strs) - 1
            
            with d_col2:
                sel_end_str = st.selectbox("📅 終了点検日を選択", options=all_date_strs, index=curr_end_idx, key="sel_end_str")
                st.session_state["g_end_d"] = pd.to_datetime(sel_end_str).date()

            start_d = st.session_state["g_start_d"]
            end_d = st.session_state["g_end_d"]

            # 選択期間内の点検回数表示
            active_dates_in_range = [d for d in site_dates_dt.dt.date if start_d <= d <= end_d]
            st.info(f"💡 **選択中の期間 ({start_d.strftime('%Y/%m/%d')} 〜 {end_d.strftime('%Y/%m/%d')}) に含まれる点検日**: **{len(active_dates_in_range)} 回**")

            st.markdown("---")
            
            # 対象期間の Daily Report を抽出
            site_reps_filtered = df_rep[(
                df_rep["Site_ID"] == selected_site_id
            ) & (
                df_rep["Sheet_Type"] == selected_sheet_type
            )].copy()
            site_reps_filtered["Date_dt"] = pd.to_datetime(site_reps_filtered["Date"])
            
            if start_d and end_d:
                site_reps_filtered = site_reps_filtered[
                    (site_reps_filtered["Date_dt"].dt.date >= start_d) &
                    (site_reps_filtered["Date_dt"].dt.date <= end_d)
                ]
            site_reps_filtered = site_reps_filtered.sort_values("Date_dt")
            
            # 各槽のカラーパレット定義 (Y1とY2で明確に異なる対比色・すべて実線)
            # 槽1: 左軸=深緑 (#2E7D32) / 右軸=鮮やかブルー (#1976D2)
            # 槽2: 左軸=黄緑 (#7CB342) / 右軸=パープル (#8E24AA)
            # 槽3: 左軸=オレンジ (#F57C00) / 右軸=ターコイズシアン (#00ACC1)
            loc_colors = [
                {"y1": "#2E7D32", "y2": "#1976D2", "label": "槽1(緑/青)"},
                {"y1": "#7CB342", "y2": "#8E24AA", "label": "槽2(黄緑/紫)"},
                {"y1": "#F57C00", "y2": "#00ACC1", "label": "槽3(橙/水色)"}
            ]
            
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            has_data = False
            summary_data = []
            
            for l_idx, loc_id in enumerate(selected_loc_ids):
                loc_name = graph_loc_options[loc_id]
                color_cfg = loc_colors[l_idx % len(loc_colors)]
                
                # 第1縦軸 (Y1)
                item_y1_row = df_item[(
                    df_item["Loc_ID"] == loc_id
                ) & (
                    df_item["Sheet_Type"] == selected_sheet_type
                ) & (
                    df_item["Item_Name"] == selected_item_y1
                )]
                
                if not item_y1_row.empty:
                    item_y1_id = item_y1_row.iloc[0]["Item_ID"]
                    rec_y1 = df_rec[df_rec["Item_ID"] == item_y1_id][["Report_ID", "Value_Num"]]
                    df_plot_y1 = site_reps_filtered.merge(rec_y1, on="Report_ID", how="left")
                    
                    valid_y1 = df_plot_y1.dropna(subset=["Value_Num"])
                    if not valid_y1.empty:
                        has_data = True
                        fig.add_trace(
                            go.Scatter(
                                x=df_plot_y1["Date"],
                                y=df_plot_y1["Value_Num"],
                                name=f"【{loc_name}】{item_display_map[selected_item_y1]}",
                                mode="lines+markers",
                                connectgaps=False,
                                line=dict(color=color_cfg["y1"], width=3),
                                marker=dict(size=8, symbol="circle")
                            ),
                            secondary_y=False
                        )
                        summary_data.append({
                            "槽名": loc_name,
                            "軸": "左軸",
                            "項目": item_display_map[selected_item_y1],
                            "平均値": f"{valid_y1['Value_Num'].mean():.2f}",
                            "最新値": f"{valid_y1.iloc[-1]['Value_Num']:.2f}",
                            "最小値": f"{valid_y1['Value_Num'].min():.2f}",
                            "最大値": f"{valid_y1['Value_Num'].max():.2f}"
                        })
                
                # 第2縦軸 (Y2 - オプション)
                if selected_item_y2:
                    item_y2_row = df_item[(
                        df_item["Loc_ID"] == loc_id
                    ) & (
                        df_item["Sheet_Type"] == selected_sheet_type
                    ) & (
                        df_item["Item_Name"] == selected_item_y2
                    )]
                    
                    if not item_y2_row.empty:
                        item_y2_id = item_y2_row.iloc[0]["Item_ID"]
                        rec_y2 = df_rec[df_rec["Item_ID"] == item_y2_id][["Report_ID", "Value_Num"]]
                        df_plot_y2 = site_reps_filtered.merge(rec_y2, on="Report_ID", how="left")
                        
                        valid_y2 = df_plot_y2.dropna(subset=["Value_Num"])
                        if not valid_y2.empty:
                            has_data = True
                            fig.add_trace(
                                go.Scatter(
                                    x=df_plot_y2["Date"],
                                    y=df_plot_y2["Value_Num"],
                                    name=f"【{loc_name}】{item_display_map[selected_item_y2]} (右軸)",
                                    mode="lines+markers",
                                    connectgaps=False,
                                    line=dict(color=color_cfg["y2"], width=3),
                                    marker=dict(size=8, symbol="diamond")
                                ),
                                secondary_y=True
                            )
                            summary_data.append({
                                "槽名": loc_name,
                                "軸": "右軸",
                                "項目": item_display_map[selected_item_y2],
                                "平均値": f"{valid_y2['Value_Num'].mean():.2f}",
                                "最新値": f"{valid_y2.iloc[-1]['Value_Num']:.2f}",
                                "最小値": f"{valid_y2['Value_Num'].min():.2f}",
                                "最大値": f"{valid_y2['Value_Num'].max():.2f}"
                            })
                            
            if not has_data:
                st.warning("選択した期間・槽・項目には数値データが存在しません。")
            else:
                title_locs_str = " vs ".join([graph_loc_options[lid] for lid in selected_loc_ids])
                title_text = f"【{site_options[selected_site_id]}】 ({title_locs_str}) 水質変化比較グラフ"
                
                fig.update_layout(
                    title_text=title_text,
                    xaxis_title="点検日 (Date)",
                    template="plotly_white",
                    hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                
                fig.update_yaxes(title_text=f"<b>{item_display_map[selected_item_y1]}</b>", secondary_y=False, title_font=dict(color="#2E7D32"))
                if selected_item_y2:
                    fig.update_yaxes(title_text=f"<b>{item_display_map[selected_item_y2]}</b>", secondary_y=True, title_font=dict(color="#1565C0"))
                    
                st.plotly_chart(fig, use_container_width=True)
                st.caption("📸 グラフ右上のカメラアイコンをタップすると、グラフをPNG画像としてワンクリック保存できます。")
                
                if summary_data:
                    st.markdown("##### 📈 槽別・項目別 統計サマリー一覧")
                    df_sum = pd.DataFrame(summary_data)
                    st.dataframe(df_sum, use_container_width=True)

# ------------------------------------------
# TAB 3: エクセル帳票ダウンロード (要件⑥)
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
        
        mi_cols = [("基本情報", "日付")] + [(loc_name, item_name) for _, loc_name, item_name in col_defs] + [("基本情報", "備考")]
        raw_col_keys = ["Date"] + [item_id for item_id, _, _ in col_defs] + ["Notes"]
        
        df_preview = df_raw[raw_col_keys].copy()
        df_preview.columns = pd.MultiIndex.from_tuples(mi_cols)
        
        st.markdown("##### プレビュー (最新10件 - 2段ヘッダー構造)")
        st.dataframe(df_preview.tail(10), use_container_width=True)
        
        def generate_formatted_excel(selected_site_id, selected_sheet_type, site_name):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"{selected_sheet_type}_集計"
            
            ws.append([f"◆ {site_name} 【{selected_sheet_type}】 全点検データ一覧表"])
            ws.append([f"出力日時: {datetime.datetime.now().strftime('%Y/%m/%d %H:%M')}"])
            ws.append([])
            
            header1 = ["日付"] + [loc_name for _, loc_name, _ in col_defs] + ["備考"]
            header2 = ["日付"] + [item_name for _, _, item_name in col_defs] + ["備考"]
            
            ws.append(header1)
            ws.append(header2)
            
            for _, r in df_raw.iterrows():
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
                    
            green_fill = PatternFill(start_color="43A047", end_color="43A047", fill_type="solid")
            header_font = Font(name="游ゴシック", size=10, bold=True, color="FFFFFF")
            data_font = Font(name="游ゴシック", size=10)
            thin_border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9")
            )
            
            ws["A1"].font = Font(name="游ゴシック", size=14, bold=True, color="388E3C")
            
            for r in range(4, 6):
                for c in range(1, total_cols + 1):
                    cell = ws.cell(row=r, column=c)
                    cell.fill = green_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = thin_border
                    
            for r_idx in range(6, ws.max_row + 1):
                for c_idx in range(1, total_cols + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.font = data_font
                    cell.border = thin_border
                    val_str = str(cell.value or "").replace(",", "").strip()
                    try:
                        num_v = float(val_str)
                        cell.value = num_v
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    except ValueError:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                        
            ws.freeze_panes = "B6"
            
            for col in ws.columns:
                max_l = max(len(str(cell.value or "")) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_l + 3, 12)
                
            wb.save(output)
            output.seek(0)
            return output
            
        excel_bytes = generate_formatted_excel(selected_site_id, selected_sheet_type, site_options[selected_site_id])
        
        st.download_button(
            label="📥 原本と同仕様のExcelデータ (.xlsx) をダウンロードする",
            data=excel_bytes,
            file_name=f"{site_options[selected_site_id]}_{selected_sheet_type}_全データ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ------------------------------------------
# TAB 4: データベース全件閲覧
# ------------------------------------------
with tab4:
    st.markdown("<div class='sub-header'>🗃️ データベース マスター＆実績テーブル閲覧</div>", unsafe_allow_html=True)
    
    table_choice = st.selectbox("閲覧するテーブルを選択", ["Site Master", "Location Master", "Item Master", "Daily Report", "Inspection Records"])
    
    if table_choice == "Site Master":
        st.dataframe(df_site, use_container_width=True)
    elif table_choice == "Location Master":
        st.dataframe(df_loc[df_loc["Site_ID"] == selected_site_id], use_container_width=True)
    elif table_choice == "Item Master":
        st.dataframe(df_item[df_item["Sheet_Type"] == selected_sheet_type], use_container_width=True)
    elif table_choice == "Daily Report":
        st.dataframe(df_rep[df_rep["Site_ID"] == selected_site_id], use_container_width=True)
    elif table_choice == "Inspection Records":
        st.dataframe(df_rec.head(500), use_container_width=True)
