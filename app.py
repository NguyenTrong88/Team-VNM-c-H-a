import streamlit as st
import pandas as pd

# Cấu hình giao diện trang web hiển thị rộng rãi
st.set_page_config(page_title="Phần mềm Lọc Dữ Liệu Excel", layout="wide")

st.title("📊 ỨNG DỤNG LỌC DỮ LIỆU EXCEL TỪ GOOGLE DRIVE")
st.subheader("📋 Đang đọc dữ liệu từ sheet: DSKH (Cố định 2 dòng tiêu đề)")
st.write("Dữ liệu được cập nhật theo thời gian thực từ file Excel trên Drive của bạn.")

# --- BƯỚC THAY ĐỔI THÔNG TIN CỦA BẠN ---
FILE_ID = "1ovZyqNg6hQVqEHXpfBnd1-zfdHcDa7TD"

# Tên các cột thật trong Excel tương ứng với từng chỉ tiêu
COT_MA_KH   = "Mã KH"       # Cột chứa Mã KH
COT_TEN_KH  = "Tên KH"      # Cột chứa Tên KH

# V_SHOP
COT_H = "CỘT H"   # V_SHOP TL  - Chỉ tiêu
COT_Q = "CỘT Q"   # V_SHOP TL  - Thực hiện
COT_I = "CỘT I"   # V_SHOP TB  - Chỉ tiêu (không có Thực hiện)

# MẸ VÀ BÉ
COT_J = "CỘT J"   # MẸ VÀ BÉ TL      - Chỉ tiêu (không có Thực hiện)
COT_K = "CỘT K"   # MẸ VÀ BÉ SB_BDD TB - Chỉ tiêu
COT_R = "CỘT R"   # MẸ VÀ BÉ SB_BDD TB - Thực hiện
COT_L = "CỘT L"   # MẸ VÀ BÉ SBPS TB   - Chỉ tiêu
COT_S = "CỘT S"   # MẸ VÀ BÉ SBPS TB   - Thực hiện
# --------------------------------------

excel_url = f"https://drive.google.com/uc?export=download&id={FILE_ID}"

@st.cache_data(ttl=10)
def load_data():
    try:
        df_raw = pd.read_excel(excel_url, sheet_name='DSKH', header=None, engine='openpyxl')
        return df_raw
    except ValueError:
        st.error("❌ Không tìm thấy sheet tên là 'DSKH' trong file Excel của bạn.")
        return None
    except Exception as e:
        st.error(f"❌ Không thể kết nối đến file Excel trên Drive. Lỗi: {e}")
        return None

df_raw = load_data()

if df_raw is not None:
    # Bỏ cột A (cột đầu tiên)
    df_raw = df_raw.iloc[:, 1:]

    row_4 = df_raw.iloc[3].fillna("").astype(str).str.strip().tolist()
    row_5 = df_raw.iloc[4].fillna("").astype(str).str.strip().tolist()

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

    df_data = df_raw.iloc[5:].copy()
    df_data.columns = headers_for_filter
    df_data = df_data.reset_index(drop=True)

    # --- SIDEBAR BỘ LỌC ---
    st.sidebar.header("Bộ Lọc Dữ Liệu DSKH")
    search_query = st.sidebar.text_input("🔍 Tìm kiếm nhanh (Mã, Tên, SĐT...):")

    filter_options = [col for col in headers_for_filter if not col.startswith("Cột_Trống_")]
    selected_filter_cols = st.sidebar.multiselect(
        "Chọn các cột bạn muốn lọc chi tiết:",
        options=filter_options,
        default=filter_options[:2] if len(filter_options) >= 2 else filter_options
    )

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

    # --- BẢNG HTML CHÍNH ---
    filtered_df_clean = filtered_df.copy().fillna("")
    filtered_df_clean = filtered_df_clean.map(
        lambda x: "" if str(x).strip().lower() in ["nan", "nat", "null", "#n/a"] else x
    )

    TEN_KH_INDEX = 2

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

    # ============================================================
    # PHẦN 3: BẢNG THỐNG KÊ THEO LAYOUT MỚI
    # ============================================================
    st.markdown("---")
    st.subheader("📊 Thông tin chi tiết khách hàng sau khi lọc")

    def get_val(df, col):
        """Lấy giá trị đầu tiên của cột, trả về '' nếu không có hoặc cột không tồn tại."""
        if col in df.columns and len(df) > 0:
            v = df[col].iloc[0]
            if str(v).strip().lower() in ["nan", "nat", "null", "#n/a", ""]:
                return ""
            return v
        return ""

    def fmt(df, col):
        """Tính tổng cột số, format có dấu phẩy. Trả về '' nếu cột không tồn tại."""
        if col in df.columns and len(df) > 0:
            total = pd.to_numeric(df[col], errors='coerce').sum()
            if pd.isna(total):
                return ""
            return f"{total:,.0f}"
        return f"⚠️ Không tìm thấy cột '{col}'"

    # Lấy thông tin KH (nếu lọc được đúng 1 KH thì hiện, nhiều KH thì hiện tổng hợp)
    so_kh = filtered_df[COT_MA_KH].nunique() if COT_MA_KH in filtered_df.columns else 0

    if len(filtered_df) == 0:
        st.info("⚠️ Không có dữ liệu. Vui lòng kiểm tra lại bộ lọc.")
    else:
        # Hiển thị Mã KH / Tên KH
        if so_kh == 1:
            ma_kh_hien  = get_val(filtered_df, COT_MA_KH)
            ten_kh_hien = get_val(filtered_df, COT_TEN_KH)
            label_ma  = f"**Mã KH:** {ma_kh_hien}"
            label_ten = f"**Tên KH:** {ten_kh_hien}"
        else:
            label_ma  = f"**Mã KH:** ({so_kh} khách hàng)"
            label_ten = f"**Tên KH:** (nhiều khách hàng)"

        st.markdown(label_ma)
        st.markdown(label_ten)
        st.markdown("")

        # Xây dựng bảng HTML thống kê
        rows_stat = [
            # (Nhóm, Loại, Chỉ tiêu col, Thực hiện col)
            ("V_SHOP",    "TL",       COT_H, COT_Q),
            ("V_SHOP",    "TB",       COT_I, None),
            ("MẸ VÀ BÉ", "TL",       COT_J, None),
            ("MẸ VÀ BÉ", "SB_BDD TB",COT_K, COT_R),
            ("MẸ VÀ BÉ", "SBPS TB",  COT_L, COT_S),
        ]

        html_stat = """
        <style>
            .stat-table {
                border-collapse: collapse;
                font-family: sans-serif;
                font-size: 14px;
                width: 100%;
                max-width: 600px;
            }
            .stat-table th {
                background-color: #f1f3f5;
                color: #495057;
                padding: 8px 14px;
                border: 1px solid #dee2e6;
                text-align: center;
            }
            .stat-table td {
                padding: 7px 14px;
                border: 1px solid #dee2e6;
                text-align: left;
            }
            .stat-table .group-cell {
                font-weight: bold;
                background-color: #f8f9fa;
                vertical-align: middle;
                text-align: center;
                color: #343a40;
            }
            .stat-table .label-cell {
                color: #495057;
                padding-left: 18px;
            }
            .stat-table .num-cell {
                text-align: right;
                color: #212529;
                font-variant-numeric: tabular-nums;
            }
            .stat-table .empty-cell {
                text-align: center;
                color: #adb5bd;
            }
        </style>
        <table class="stat-table">
            <thead>
                <tr>
                    <th style="width:130px"></th>
                    <th style="width:120px"></th>
                    <th style="width:150px">Chỉ tiêu</th>
                    <th style="width:150px">Thực hiện</th>
                </tr>
            </thead>
            <tbody>
        """

        # Tính rowspan cho từng nhóm
        from itertools import groupby
        groups = {}
        for nhom, loai, c_ct, c_th in rows_stat:
            groups[nhom] = groups.get(nhom, 0) + 1

        prev_group = None
        for nhom, loai, c_ct, c_th in rows_stat:
            ct_val = fmt(filtered_df, c_ct)
            th_val = fmt(filtered_df, c_th) if c_th else None

            group_td = ""
            if nhom != prev_group:
                rowspan = groups[nhom]
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

        st.components.v1.html(html_stat, height=230, scrolling=False)

    st.markdown(f"**Tổng số dòng lọc được:** {len(filtered_df)} dòng")
