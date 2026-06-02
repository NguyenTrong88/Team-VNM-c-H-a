import streamlit as st
import pandas as pd

st.set_page_config(page_title="Phần mềm Lọc Dữ Liệu Excel", layout="wide")

st.title("📊 ỨNG DỤNG LỌC DỮ LIỆU EXCEL TỪ GOOGLE DRIVE")
st.subheader("📋 Đang đọc dữ liệu từ sheet: DSKH (Tiêu đề dòng 4, dữ liệu từ dòng 5)")
st.write("Dữ liệu được cập nhật theo thời gian thực từ file Excel trên Drive của bạn.")

# ================================================================
# CẤU HÌNH — CHỈ CẦN CHỈNH PHẦN NÀY
# ================================================================
FILE_ID = "1ovZyqNg6hQVqEHXpfBnd1-zfdHcDa7TD"

# Các cột muốn hiển thị trong bảng (A=0, B=1, C=2, ...)
COLS_HIEN_THI = [
    1,   # B
    2,   # C — Mã KH
    3,   # D — Tên KH
    6,   # G
    7,   # H
    8,   # I
    9,   # J
    10,  # K
    11,  # L
    12,  # M
    13,  # N
    14,  # O
    15,  # P
    16,  # Q
    17,  # R
    18,  # S
    19,  # T
    20,  # U
    21,  # V
    30,  # AE
    31,  # AF
    32,  # AG
    33,  # AH
    34,  # AI
    35,  # AJ
    36,  # AK
    37,  # AL
    38,  # AM
    39,  # AN
    40,  # AO
    41,  # AP
    42,  # AQ
    51,  # AZ
]

COL_MA_KH  = 2   # Cột C — Mã KH
COL_TEN_KH = 3   # Cột D — Tên KH

# Bộ lọc cố định sidebar
COLS_LOC_CO_DINH = [
    "Tên NVBH",
    "Mã KH",
    "Tên KH",
    "V_SHOP TL",
    "MẸ VÀ BÉ TL",
    "VNM_SHOP TL",
    "VIP_SHOP TL",
    "PSDS",
]

# Vị trí cột gốc Excel cho bảng thống kê (A=0, B=1, ...)
COL_G  = 6    # V_SHOP TL           — Chỉ tiêu
COL_H  = 7    # V_SHOP TB           — Chỉ tiêu
COL_I  = 8    # MẸ VÀ BÉ TL        — Chỉ tiêu
COL_J  = 9    # MẸ VÀ BÉ SB_BDD TB — Chỉ tiêu
COL_K  = 10   # MẸ VÀ BÉ SBPS TB   — Chỉ tiêu
COL_L  = 11   # VNM_SHOP TL         — Chỉ tiêu
COL_M  = 12   # VNM_SHOP TB         — Chỉ tiêu
COL_N  = 13   # VIP_SHOP TL         — Chỉ tiêu
COL_O  = 14   # VIP_SHOP TB         — Chỉ tiêu
COL_P  = 15   # V_SHOP TL           — Thực hiện
COL_Q  = 16   # MẸ VÀ BÉ SB_BDD TB — Thực hiện
COL_R  = 17   # MẸ VÀ BÉ SBPS TB   — Thực hiện
COL_S  = 18   # VNM_SHOP TL         — Thực hiện
COL_T  = 19   # VIP_SHOP TL         — Thực hiện
COL_AU = 46   # A-SBPS              — Thực hiện
# ================================================================

excel_url = f"https://drive.google.com/uc?export=download&id={FILE_ID}"

@st.cache_data(ttl=10)
def load_data():
    try:
        df_raw = pd.read_excel(excel_url, sheet_name='DSKH', header=None, engine='openpyxl')
        return df_raw
    except ValueError:
        st.error("❌ Không tìm thấy sheet tên là 'DSKH'. Vui lòng kiểm tra lại tên sheet.")
        return None
    except Exception as e:
        st.error(f"❌ Không thể kết nối đến file Excel trên Drive. Lỗi: {e}")
        return None

df_raw = load_data()

if df_raw is not None:
    df_raw_full = df_raw.copy()

    valid_cols = [c for c in COLS_HIEN_THI if c < df_raw.shape[1]]
    df_raw_display = df_raw.iloc[:, valid_cols]

    # ----------------------------------------------------------------
    # CHỈ 1 DÒNG TIÊU ĐỀ (dòng 4 = index 3), dữ liệu từ dòng 5 (index 4)
    # ----------------------------------------------------------------
    row_header = df_raw_display.iloc[3].fillna("").astype(str).str.strip().tolist()

    # Chuẩn hóa tên cột
    headers_for_filter = []
    for idx, r in enumerate(row_header):
        r_clean = "" if r.lower().startswith("unnamed:") else r
        if r_clean == "":
            headers_for_filter.append(f"Cột_Trống_{idx+1}")
        else:
            if r_clean not in headers_for_filter:
                headers_for_filter.append(r_clean)
            else:
                headers_for_filter.append(f"{r_clean}_{idx}")

    # Dữ liệu từ dòng 5 (index 4) trở đi
    df_data = df_raw_display.iloc[4:].copy()
    df_data.columns = headers_for_filter
    df_data = df_data.reset_index(drop=True)

    # df_data_full giữ nguyên tất cả cột gốc để tra số liệu
    df_data_full = df_raw_full.iloc[4:].copy().reset_index(drop=True)

    TEN_KH_DISPLAY_IDX = valid_cols.index(COL_TEN_KH) if COL_TEN_KH in valid_cols else -1

    # ================================================================
    # SIDEBAR BỘ LỌC CỐ ĐỊNH
    # ================================================================
    st.sidebar.header("Bộ Lọc Dữ Liệu DSKH")
    search_query = st.sidebar.text_input("🔍 Tìm kiếm nhanh (Mã, Tên, SĐT...):")

    cols_loc_ton_tai = [col for col in COLS_LOC_CO_DINH if col in headers_for_filter]

    filtered_df = df_data.copy()

    if search_query:
        mask = df_data.astype(str).apply(
            lambda x: x.str.contains(search_query, case=False, na=False)
        ).any(axis=1)
        filtered_df = filtered_df[mask]

    for col in cols_loc_ton_tai:
        unique_vals = [str(val).strip() for val in df_data[col].unique()
                       if str(val).strip() not in ["", "nan", "NaT", "None"]]
        unique_vals = sorted(list(set(unique_vals)))
        selected_vals = st.sidebar.multiselect(f"Lọc theo {col}:", options=unique_vals)
        if selected_vals:
            filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]

    filtered_df_full = df_data_full.loc[filtered_df.index]

    # ================================================================
    # BẢNG HIỂN THỊ CHÍNH
    # ================================================================
    def render_cell(val):
        raw = str(val).strip()
        if raw.lower() in ["nan", "nat", "null", "#n/a", ""]:
            return "", True
        num = pd.to_numeric(val, errors='coerce')
        if pd.notna(num):
            if num == 0:
                return "", True
            return f"{round(num):,}", False
        return raw, False

    html_rows = ""
    for _, row in filtered_df.iterrows():
        html_rows += "<tr>"
        for idx, val in enumerate(row):
            content, is_zero = render_cell(val)
            bg = " style='background-color:#ffe5e5;'" if is_zero else ""
            num_check = pd.to_numeric(val, errors='coerce')
            extra_class = "num-cell" if pd.notna(num_check) else ""
            if idx == TEN_KH_DISPLAY_IDX:
                html_rows += f"<td class='text-wrap-column'{bg}>{content}</td>"
            else:
                html_rows += f"<td class='{extra_class}'{bg}>{content}</td>"
        html_rows += "</tr>"

    html_header = "<tr>"
    for idx, h in enumerate(row_header):
        h_display = "" if h.lower().startswith("unnamed:") else h
        html_header += f"<th class='text-wrap-column'>{h_display}</th>" if idx == TEN_KH_DISPLAY_IDX else f"<th>{h_display}</th>"
    html_header += "</tr>"

    table_html = f"""
    <style>
        .table-container {{ max-height: 500px; overflow-y: auto; border: 1px solid #dee2e6; font-family: sans-serif; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th, td {{ padding: 8px 10px; border: 1px solid #dee2e6; text-align: left; white-space: nowrap; }}
        .num-cell {{ text-align: right; font-variant-numeric: tabular-nums; }}
        .text-wrap-column {{ white-space: normal !important; min-width: 180px !important; max-width: 220px !important; word-break: break-word; }}
        thead tr th {{ position: sticky; top: 0; background-color: #f1f3f5; color: #495057; z-index: 10; }}
        tbody tr:hover td {{ background-color: #f0f4ff !important; }}
    </style>
    <div class="table-container">
        <table>
            <thead>{html_header}</thead>
            <tbody>{html_rows}</tbody>
        </table>
    </div>
    """
    st.components.v1.html(table_html, height=520, scrolling=False)

    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 Tải danh sách đã lọc về máy (.CSV)",
        data=csv,
        file_name="dskh_da_loc.csv",
        mime="text/csv",
    )

    # ================================================================
    # PHẦN 3: BẢNG THỐNG KÊ
    # ================================================================
    st.markdown("---")
    st.subheader("📊 Thông tin chi tiết khách hàng sau khi lọc")

    def sum_col(df_full, col_idx):
        if col_idx is None or col_idx >= df_full.shape[1]:
            return None
        total = pd.to_numeric(df_full.iloc[:, col_idx], errors='coerce').sum()
        return float(total) if pd.notna(total) else None

    def fmt_stat(val):
        if val is None:
            return None, None
        if val == 0:
            return "", True
        return f"{round(val):,}", False

    def get_str_full(df_full, col_idx):
        if col_idx >= df_full.shape[1] or len(df_full) == 0:
            return ""
        v = df_full.iloc[0, col_idx]
        if str(v).strip().lower() in ["nan", "nat", "null", "#n/a", ""]:
            return ""
        return str(v).strip()

    def stat_td(val, is_zero, no_data=False):
        if no_data:
            return '<td class="no-data-cell">—</td>'
        if is_zero:
            return '<td class="zero-cell"></td>'
        return f'<td class="num-cell">{val}</td>'

    if len(filtered_df_full) == 0:
        st.info("⚠️ Không có dữ liệu. Vui lòng kiểm tra lại bộ lọc.")
    else:
        so_kh = filtered_df_full.iloc[:, COL_MA_KH].nunique()
        if so_kh == 1:
            ma_kh_hien  = get_str_full(filtered_df_full, COL_MA_KH)
            ten_kh_hien = get_str_full(filtered_df_full, COL_TEN_KH)
        else:
            ma_kh_hien  = f"({so_kh} khách hàng)"
            ten_kh_hien = "(nhiều khách hàng)"

        s = {
            'G':  sum_col(filtered_df_full, COL_G),
            'H':  sum_col(filtered_df_full, COL_H),
            'I':  sum_col(filtered_df_full, COL_I),
            'J':  sum_col(filtered_df_full, COL_J),
            'K':  sum_col(filtered_df_full, COL_K),
            'L':  sum_col(filtered_df_full, COL_L),
            'M':  sum_col(filtered_df_full, COL_M),
            'N':  sum_col(filtered_df_full, COL_N),
            'O':  sum_col(filtered_df_full, COL_O),
            'P':  sum_col(filtered_df_full, COL_P),
            'Q':  sum_col(filtered_df_full, COL_Q),
            'R':  sum_col(filtered_df_full, COL_R),
            'S':  sum_col(filtered_df_full, COL_S),
            'T':  sum_col(filtered_df_full, COL_T),
            'AU': sum_col(filtered_df_full, COL_AU),
        }

        def con_lai(ct, th):
            if ct is None or th is None:
                return None
            return ct - th

        rows_stat = [
            ("V_SHOP",    "TL",        'G',  'P',  con_lai(s['G'], s['P'])),
            ("V_SHOP",    "TB",        'H',  None, None),
            ("MẸ VÀ BÉ", "TL",        'I',  None, None),
            ("MẸ VÀ BÉ", "SB_BDD TB", 'J',  'Q',  con_lai(s['J'], s['Q'])),
            ("MẸ VÀ BÉ", "SBPS TB",   'K',  'R',  con_lai(s['K'], s['R'])),
            ("VNM_SHOP",  "TL",        'L',  'S',  con_lai(s['L'], s['S'])),
            ("VNM_SHOP",  "TB",        'M',  None, None),
            ("VIP_SHOP",  "TL",        'N',  'T',  con_lai(s['N'], s['T'])),
            ("VIP_SHOP",  "TB",        'O',  None, None),
            ("A-SBPS",    "",          None, 'AU', None),
        ]

        rowspan_map = {}
        for nhom, _, _, _, _ in rows_stat:
            rowspan_map[nhom] = rowspan_map.get(nhom, 0) + 1

        html_stat = """
        <style>
            .stat-wrap { max-width: 740px; font-family: sans-serif; }
            .kh-info { margin-bottom: 10px; border-collapse: collapse; width: 100%; max-width: 740px; }
            .kh-info td { border: 1px solid #2d7ab8; padding: 7px 16px; }
            .kh-label { background-color: #3a8fd4; color: white; font-weight: bold; text-align: center; min-width: 160px; }
            .kh-value { background: white; min-width: 200px; }
            .stat-table { border-collapse: collapse; font-size: 14px; width: 100%; max-width: 740px; }
            .stat-table th { background-color: #3a8fd4; color: white; padding: 8px 14px; border: 1px solid #2d7ab8; text-align: center; }
            .stat-table td { padding: 7px 14px; border: 1px solid #dee2e6; }
            .group-cell { font-weight: bold; background-color: #cce8f4; vertical-align: middle; text-align: center; color: #1a5f8a; border: 1px solid #2d7ab8 !important; }
            .label-cell { color: #212529; padding-left: 14px; background-color: #e8f4fc; }
            .num-cell { text-align: right; color: #212529; font-variant-numeric: tabular-nums; }
            .zero-cell { background-color: #ffe5e5; }
            .no-data-cell { background-color: #f8f9fa; text-align: center; color: #adb5bd; }
        </style>
        <div class="stat-wrap">
        """

        html_stat += f"""
        <table class="kh-info">
            <tr><td class="kh-label">Mã KH</td><td class="kh-value">{ma_kh_hien}</td><td style="border:none"></td><td style="border:none"></td></tr>
            <tr><td class="kh-label">Tên KH</td><td class="kh-value">{ten_kh_hien}</td><td style="border:none"></td><td style="border:none"></td></tr>
        </table>
        <table class="stat-table">
            <thead>
                <tr>
                    <th style="width:150px"></th>
                    <th style="width:120px"></th>
                    <th style="width:130px">Chỉ tiêu</th>
                    <th style="width:130px">Thực Hiện</th>
                    <th style="width:130px">Còn Lại</th>
                </tr>
            </thead>
            <tbody>
        """

        prev_group = None
        for nhom, loai, key_ct, key_th, val_cl in rows_stat:
            ct_td = stat_td(*fmt_stat(s[key_ct])) if key_ct else '<td class="no-data-cell">—</td>'
            th_td = stat_td(*fmt_stat(s[key_th])) if key_th else '<td class="no-data-cell">—</td>'
            cl_td = stat_td(*fmt_stat(val_cl))     if val_cl is not None else '<td class="no-data-cell">—</td>'

            group_td = ""
            if nhom != prev_group:
                rowspan = rowspan_map[nhom]
                group_td = f'<td class="group-cell" rowspan="{rowspan}">{nhom}</td>'
                prev_group = nhom

            html_stat += f"""
                <tr>
                    {group_td}
                    <td class="label-cell">{loai}</td>
                    {ct_td}{th_td}{cl_td}
                </tr>
            """

        html_stat += "</tbody></table></div>"
        st.components.v1.html(html_stat, height=430, scrolling=False)

    st.markdown(f"**Tổng số dòng lọc được:** {len(filtered_df)} dòng")
