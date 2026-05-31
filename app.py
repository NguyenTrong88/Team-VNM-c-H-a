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

# Vị trí cột theo Excel (A=0, B=1, C=2, D=3, ...)
COL_MA_KH  = 2   # Cột C — Mã KH
COL_TEN_KH = 3   # Cột D — Tên KH

# Vị trí cột số liệu cho bảng thống kê
COL_H = 7    # V_SHOP TL     — Chỉ tiêu
COL_Q = 16   # V_SHOP TL     — Thực hiện
COL_I = 8    # V_SHOP TB     — Chỉ tiêu (không có Thực hiện)
COL_J = 9    # MẸ VÀ BÉ TL  — Chỉ tiêu (không có Thực hiện)
COL_K = 10   # MẸ VÀ BÉ SB_BDD TB — Chỉ tiêu
COL_R = 17   # MẸ VÀ BÉ SB_BDD TB — Thực hiện
COL_L = 11   # MẸ VÀ BÉ SBPS TB   — Chỉ tiêu
COL_S = 18   # MẸ VÀ BÉ SBPS TB   — Thực hiện

# Cột dùng trong bộ lọc sidebar (dùng tên tiêu đề từ dòng 4)
COT_TINH_TONG = "PSDS T4"
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
    # Lưu lại df_raw gốc (giữ đủ cột) để tra cứu số liệu theo vị trí cột
    df_raw_full = df_raw.copy()

    # Trích xuất dòng tiêu đề (dòng 4 và dòng 5, index 3 và 4)
    row_4 = df_raw.iloc[3].fillna("").astype(str).str.strip().tolist()
    row_5 = df_raw.iloc[4].fillna("").astype(str).str.strip().tolist()

    # Chuẩn hóa tên cột cho bộ lọc
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

    # Dữ liệu từ dòng 6 trở đi
    df_data = df_raw.iloc[5:].copy()
    df_data.columns = headers_for_filter
    df_data = df_data.reset_index(drop=True)

    # Tạo df_data_full giữ nguyên vị trí cột (không đổi tên) để tra cứu theo index
    df_data_full = df_raw_full.iloc[5:].copy().reset_index(drop=True)

    # --- SIDEBAR BỘ LỌC ---
    st.sidebar.header("Bộ Lọc Dữ Liệu DSKH")
    search_query = st.sidebar.text_input("🔍 Tìm kiếm nhanh (Mã, Tên, SĐT...):")

    filter_options = [col for col in headers_for_filter if not col.startswith("Cột_Trống_")]
    selected_filter_cols = st.sidebar.multiselect(
        "Chọn các cột bạn muốn lọc chi tiết:",
        options=filter_options,
        default=filter_options[:2] if len(filter_options) >= 2 else filter_options
    )

    # Áp dụng lọc — lưu lại index để đồng bộ với df_data_full
    filtered_df = df_data.copy()
    if search_query:
        mask = df_data.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        filtered_df = filtered_df[mask]

    for col in selected_filter_cols:
        unique_vals = [str(val).strip() for val in df_data[col].unique() if str(val).strip() != ""]
        unique_vals = sorted(list(set(unique_vals)))
        selected_vals = st.sidebar.multiselect(f"Lọc theo {col}:", options=unique_vals)
        if selected_vals:
            filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]

    # Lấy các dòng tương ứng từ df_data_full theo index đã lọc
    filtered_df_full = df_data_full.loc[filtered_df.index]

    # --- BẢNG HIỂN THỊ CHÍNH ---
    filtered_df_clean = filtered_df.copy().fillna("")
    filtered_df_clean = filtered_df_clean.map(
        lambda x: "" if str(x).strip().lower() in ["nan", "nat", "null", "#n/a"] else x
    )

    # Cột Tên KH dùng để wrap chữ — lấy vị trí trong headers sau khi map
    TEN_KH_INDEX = COL_TEN_KH  # vẫn giữ nguyên vì không bỏ cột

    html_rows = ""
    for _, row in filtered_df_clean.iterrows():
        html_rows += "<tr>"
        for idx, val in enumerate(row):
            if idx == TEN_KH_INDEX:
                html_rows += f"<td class='text-wrap-column'>{val}</td>"
            else:
                html_rows += f"<td>{val}</td>"
        html_rows += "</tr>"

    html_header_4 = "<tr>"
    for idx, r4 in enumerate(row_4):
        r4_display = "" if r4.lower().startswith("unnamed:") else r4
        html_header_4 += f"<th class='text-wrap-column'>{r4_display}</th>" if idx == TEN_KH_INDEX else f"<th>{r4_display}</th>"
    html_header_4 += "</tr>"

    html_header_5 = "<tr>"
    for idx, r5 in enumerate(row_5):
        r5_display = "" if r5.lower().startswith("unnamed:") else r5
        html_header_5 += f"<th class='text-wrap-column'>{r5_display}</th>" if idx == TEN_KH_INDEX else f"<th>{r5_display}</th>"
    html_header_5 += "</tr>"

    table_html = f"""
    <style>
        .table-container {{ max-height: 500px; overflow-y: auto; border: 1px solid #dee2e6; font-family: sans-serif; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th, td {{ padding: 8px 10px; border: 1px solid #dee2e6; text-align: left; white-space: nowrap; }}
        .text-wrap-column {{ white-space: normal !important; min-width: 180px !important; max-width: 220px !important; word-break: break-word; }}
        thead tr:nth-child(1) th {{ position: sticky; top: 0; background-color: #f1f3f5; color: #495057; z-index: 10; }}
        thead tr:nth-child(2) th {{ position: sticky; top: 33px; background-color: #f8f9fa; color: #6c757d; z-index: 9; }}
        tbody tr:hover {{ background-color: #f8f9fa; }}
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
    # PHẦN 3: BẢNG THỐNG KÊ LAYOUT MỚI — LẤY SỐ LIỆU THEO VỊ TRÍ CỘT
    # ================================================================
    st.markdown("---")
    st.subheader("📊 Thông tin chi tiết khách hàng sau khi lọc")

    def fmt_col(df_full, col_idx):
        """Tính tổng cột theo vị trí index, format có dấu phẩy."""
        if col_idx >= df_full.shape[1]:
            return f"⚠️ Cột {col_idx} không tồn tại"
        total = pd.to_numeric(df_full.iloc[:, col_idx], errors='coerce').sum()
        if pd.isna(total) or total == 0:
            return "—"
        return f"{total:,.0f}"

    def get_str(df_full, col_idx):
        """Lấy giá trị text đầu tiên theo vị trí index."""
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

        # Cấu trúc bảng: (Nhóm, Loại, col_chi_tieu, col_thuc_hien hoặc None)
        rows_stat = [
            ("V_SHOP",    "TL",        COL_H, COL_Q),
            ("V_SHOP",    "TB",        COL_I, None),
            ("MẸ VÀ BÉ", "TL",        COL_J, None),
            ("MẸ VÀ BÉ", "SB_BDD TB", COL_K, COL_R),
            ("MẸ VÀ BÉ", "SBPS TB",   COL_L, COL_S),
        ]

        # Đếm rowspan mỗi nhóm
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
            .stat-table .empty-cell { text-align: center; color: #adb5bd; }
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
            ct_val = fmt_col(filtered_df_full, col_ct)
            th_val = fmt_col(filtered_df_full, col_th) if col_th is not None else None

            group_td = ""
            if nhom != prev_group:
                rowspan = rowspan_map[nhom]
                group_td = f'<td class="group-cell" rowspan="{rowspan}">{nhom}</td>'
                prev_group = nhom

            th_html = f'<td class="num-cell">{th_val}</td>' if th_val is not None else '<td class="empty-cell">—</td>'

            html_stat += f"""
                <tr>
                    {group_td}
                    <td class="label-cell">{loai}</td>
                    <td class="num-cell">{ct_val}</td>
                    {th_html}
                </tr>
            """

        html_stat += "</tbody></table>"
        st.components.v1.html(html_stat, height=240, scrolling=False)

    st.markdown(f"**Tổng số dòng lọc được:** {len(filtered_df)} dòng")
