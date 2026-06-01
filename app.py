import streamlit as st
import pandas as pd

st.set_page_config(page_title="Phần mềm Lọc Dữ Liệu Excel", layout="wide")

st.title("📊 ỨNG DỤNG LỌC DỮ LIỆU EXCEL TỪ GOOGLE DRIVE")
st.subheader("📋 Đang đọc dữ liệu từ sheet: DSKH (Cố định 2 dòng tiêu đề)")
st.write("Dữ liệu được cập nhật theo thời gian thực từ file Excel trên Drive của bạn.")

# ================================================================
# CẤU HÌNH — CHỈ CẦN CHỈNH PHẦN NÀY
# ================================================================
FILE_ID = "1ovZyqNg6hQVqEHXpfBnd1-zfdHcDa7TD"

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

COL_H = 7
COL_Q = 16
COL_I = 8
COL_J = 9
COL_K = 10
COL_R = 17
COL_L = 11
COL_S = 18
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

    row_4 = df_raw_display.iloc[3].fillna("").astype(str).str.strip().tolist()
    row_5 = df_raw_display.iloc[4].fillna("").astype(str).str.strip().tolist()

    headers_for_filter = []
    for idx, r4 in enumerate(row_4):
        r4_clean = "" if r4.lower().startswith("unnamed:") else r4
        if r4_clean == "":
            headers_for_filter.append(f"Cột_Trống_{idx+1}")
        else:
            if r4_clean not in headers_for_filter:
                headers_for_filter.append(r4_clean)
            else:
                headers_for_filter.append(f"{r4_clean}_{idx}")

    df_data = df_raw_display.iloc[5:].copy()
    df_data.columns = headers_for_filter
    df_data = df_data.reset_index(drop=True)

    df_data_full = df_raw_full.iloc[5:].copy().reset_index(drop=True)

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
    # XỬ LÝ GIÁ TRỊ: số = làm tròn, 0 = để trống + tô đỏ nhạt
    # ================================================================
    def render_cell(val, col_name):
        """
        Trả về (html_content, is_zero)
        - Số = 0 hoặc rỗng/nan: is_zero=True → ô đỏ nhạt, nội dung trống
        - Số khác 0: làm tròn, format có dấu phẩy
        - Text: giữ nguyên
        """
        raw = str(val).strip()
        if raw.lower() in ["nan", "nat", "null", "#n/a", ""]:
            return ("", True)

        num = pd.to_numeric(val, errors='coerce')
        if pd.notna(num):
            if num == 0:
                return ("", True)
            else:
                return (f"{round(num):,}", False)
        else:
            return (raw, False)

    # Tạo HTML rows với logic tô màu
    html_rows = ""
    for _, row in filtered_df.iterrows():
        html_rows += "<tr>"
        for idx, val in enumerate(row):
            content, is_zero = render_cell(val, headers_for_filter[idx])
            style = " style='background-color:#ffe5e5;'" if is_zero else ""
            if idx == TEN_KH_DISPLAY_IDX:
                html_rows += f"<td class='text-wrap-column'{style}>{content}</td>"
            else:
                # Căn phải nếu là số
                num_check = pd.to_numeric(val, errors='coerce')
                align = " num-cell" if pd.notna(num_check) else ""
                html_rows += f"<td class='{align}'{style}>{content}</td>"
        html_rows += "</tr>"

    html_header_4 = "<tr>"
    for idx, r4 in enumerate(row_4):
        r4_display = "" if r4.lower().startswith("unnamed:") else r4
        html_header_4 += f"<th class='text-wrap-column'>{r4_display}</th>" if idx == TEN_KH_DISPLAY_IDX else f"<th>{r4_display}</th>"
    html_header_4 += "</tr>"

    html_header_5 = "<tr>"
    for idx, r5 in enumerate(row_5):
        r5_display = "" if r5.lower().startswith("unnamed:") else r5
        html_header_5 += f"<th class='text-wrap-column'>{r5_display}</th>" if idx == TEN_KH_DISPLAY_IDX else f"<th>{r5_display}</th>"
    html_header_5 += "</tr>"

    table_html = f"""
    <style>
        .table-container {{
            max-height: 500px;
            overflow-y: auto;
            border: 1px solid #dee2e6;
            font-family: sans-serif;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        th, td {{
            padding: 8px 10px;
            border: 1px solid #dee2e6;
            text-align: left;
            white-space: nowrap;
        }}
        .num-cell {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .text-wrap-column {{
            white-space: normal !important;
            min-width: 180px !important;
            max-width: 220px !important;
            word-break: break-word;
        }}
        thead tr:nth-child(1) th {{
            position: sticky;
            top: 0;
            background-color: #f1f3f5;
            color: #495057;
            z-index: 10;
        }}
        thead tr:nth-child(2) th {{
            position: sticky;
            top: 33px;
            background-color: #f8f9fa;
            color: #6c757d;
            z-index: 9;
        }}
        tbody tr:hover td {{
            background-color: #f0f4ff !important;
        }}
        /* Ô đỏ nhạt khi hover vẫn giữ màu đỏ nhạt hơn */
        tbody tr:hover td[style*="ffe5e5"] {{
            background-color: #ffd0d0 !important;
        }}
    </style>
    <div class="table-container">
        <table>
            <thead>{html_header_4}{html_header_5}</thead>
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

    def fmt_stat(df_full, col_idx):
        """Tính tổng, làm tròn. Nếu = 0 trả về None (tô đỏ nhạt)."""
        if col_idx >= df_full.shape[1]:
            return None, True
        total = pd.to_numeric(df_full.iloc[:, col_idx], errors='coerce').sum()
        if pd.isna(total) or total == 0:
            return "", True   # is_zero = True
        return f"{round(total):,}", False

    def get_str(df_full, col_idx):
        if col_idx >= df_full.shape[1] or len(df_full) == 0:
            return ""
        v = df_full.iloc[0, col_idx]
        if str(v).strip().lower() in ["nan", "nat", "null", "#n/a", ""]:
            return ""
        return str(v).strip()

    if len(filtered_df_full) == 0:
        st.info("⚠️ Không có dữ liệu. Vui lòng kiểm tra lại bộ lọc.")
    else:
        so_kh = filtered_df_full.iloc[:, COL_MA_KH].nunique()

        if so_kh == 1:
            ma_kh_hien  = get_str(filtered_df_full, COL_MA_KH)
            ten_kh_hien = get_str(filtered_df_full, COL_TEN_KH)
        else:
            ma_kh_hien  = f"({so_kh} khách hàng)"
            ten_kh_hien = "(nhiều khách hàng)"

        st.markdown(f"**Mã KH:** {ma_kh_hien}")
        st.markdown(f"**Tên KH:** {ten_kh_hien}")
        st.markdown("")

        rows_stat = [
            ("V_SHOP",    "TL",        COL_H, COL_Q),
            ("V_SHOP",    "TB",        COL_I, None),
            ("MẸ VÀ BÉ", "TL",        COL_J, None),
            ("MẸ VÀ BÉ", "SB_BDD TB", COL_K, COL_R),
            ("MẸ VÀ BÉ", "SBPS TB",   COL_L, COL_S),
        ]

        rowspan_map = {}
        for nhom, _, _, _ in rows_stat:
            rowspan_map[nhom] = rowspan_map.get(nhom, 0) + 1

        html_stat = """
        <style>
            .stat-table { border-collapse: collapse; font-family: sans-serif; font-size: 14px; width: 100%; max-width: 620px; }
            .stat-table th { background-color: #f1f3f5; color: #495057; padding: 8px 14px; border: 1px solid #dee2e6; text-align: center; }
            .stat-table td { padding: 7px 14px; border: 1px solid #dee2e6; text-align: left; }
            .stat-table .group-cell { font-weight: bold; background-color: #f8f9fa; vertical-align: middle; text-align: center; color: #343a40; }
            .stat-table .label-cell { color: #495057; padding-left: 18px; }
            .stat-table .num-cell { text-align: right; color: #212529; font-variant-numeric: tabular-nums; }
            .stat-table .zero-cell { background-color: #ffe5e5; text-align: right; }
            .stat-table .no-data-cell { background-color: #f8f9fa; text-align: center; color: #adb5bd; }
        </style>
        <table class="stat-table">
            <thead>
                <tr>
                    <th style="width:130px"></th>
                    <th style="width:120px"></th>
                    <th style="width:160px">Chỉ tiêu</th>
                    <th style="width:160px">Thực hiện</th>
                </tr>
            </thead>
            <tbody>
        """

        prev_group = None
        for nhom, loai, col_ct, col_th in rows_stat:
            ct_val, ct_zero = fmt_stat(filtered_df_full, col_ct)
            if col_th is not None:
                th_val, th_zero = fmt_stat(filtered_df_full, col_th)
            else:
                th_val, th_zero = None, None

            group_td = ""
            if nhom != prev_group:
                rowspan = rowspan_map[nhom]
                group_td = f'<td class="group-cell" rowspan="{rowspan}">{nhom}</td>'
                prev_group = nhom

            ct_class = "zero-cell" if ct_zero else "num-cell"

            if th_val is None:
                th_html = '<td class="no-data-cell">—</td>'
            else:
                th_class = "zero-cell" if th_zero else "num-cell"
                th_html = f'<td class="{th_class}">{th_val}</td>'

            html_stat += f"""
                <tr>
                    {group_td}
                    <td class="label-cell">{loai}</td>
                    <td class="{ct_class}">{ct_val}</td>
                    {th_html}
                </tr>
            """

        html_stat += "</tbody></table>"
        st.components.v1.html(html_stat, height=240, scrolling=False)

    st.markdown(f"**Tổng số dòng lọc được:** {len(filtered_df)} dòng")
