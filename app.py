import streamlit as st
import pandas as pd

# Cấu hình giao diện trang web hiển thị rộng rãi
st.set_page_config(page_title="Phần mềm Lọc Dữ Liệu Excel", layout="wide")

st.title("📊 ỨNG DỤNG LỌC DỮ LIỆU EXCEL TỪ GOOGLE DRIVE")
st.subheader("📋 Đang đọc dữ liệu từ sheet: DSKH (Cố định 2 dòng tiêu đề)")
st.write("Dữ liệu được cập nhật theo thời gian thực từ file Excel trên Drive của bạn.")

# --- BƯỚC THAY ĐỔI THÔNG TIN CỦA BẠN ---
FILE_ID = "1ovZyqNg6hQVqEHXpfBnd1-zfdHcDa7TD"
COT_TINH_TONG = "PSDS T4" 
# --------------------------------------

# Đường dẫn tải trực tiếp file Excel từ Google Drive
excel_url = f"https://drive.google.com/uc?export=download&id={FILE_ID}"

# Hàm đọc dữ liệu gốc và tách thành 2 tầng tiêu đề độc lập
@st.cache_data(ttl=10)
def load_data():
    try:
        df_raw = pd.read_excel(excel_url, sheet_name='DSKH', header=None, engine='openpyxl')
        return df_raw
    except ValueError:
        st.error("❌ Không tìm thấy sheet tên là 'DSKH' trong file Excel của bạn. Vui lòng kiểm tra lại tên sheet.")
        return None
    except Exception as e:
        st.error(f"❌ Không thể kết nối đến file Excel trên Drive. Lỗi: {e}")
        return None

df_raw = load_data()

if df_raw is not None:
    # 1. Trích xuất dòng 4 (Tiêu đề thật) và dòng 5 (Tiêu đề ảo)
    row_4 = df_raw.iloc[3].fillna("").astype(str).str.strip().tolist()
    row_5 = df_raw.iloc[4].fillna("").astype(str).str.strip().tolist()
    
    # Chuẩn hóa tên cột
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

    # Lấy dữ liệu thực tế từ dòng 6 trở đi
    df_data = df_raw.iloc[5:].copy()
    df_data.columns = headers_for_filter
    df_data = df_data.reset_index(drop=True)

    # 👉 Bỏ cột A (cột đầu tiên)
    df_data = df_data.drop(df_data.columns[0], axis=1)
    headers_for_filter = headers_for_filter[1:]
    row_4 = row_4[1:]
    row_5 = row_5[1:]

    # --- PHẦN 1: BỘ LỌC ---
    st.sidebar.header("Bộ Lọc Dữ Liệu DSKH")
    search_query = st.sidebar.text_input("🔍 Tìm kiếm nhanh (Mã, Tên, SĐT...):")
    
    filter_options = [col for col in headers_for_filter if not col.startswith("Cột_Trống_")]
    selected_filter_cols = st.sidebar.multiselect("Chọn các cột bạn muốn lọc chi tiết:", options=filter_options, default=filter_options[:2] if len(filter_options)>=2 else filter_options)
    
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

    # --- PHẦN 2: BẢNG HTML ---
    filtered_df_clean = filtered_df.copy().fillna("")
    filtered_df_clean = filtered_df_clean.map(lambda x: "" if str(x).strip().lower() in ["nan", "nat", "null", "#n/a"] else x)
    
    html_rows = ""
    for _, row in filtered_df_clean.iterrows():
        html_rows += "<tr>"
        for idx, val in enumerate(row):
            if idx == 2:  # Giả sử cột Tên KH sau khi bỏ cột A nằm ở vị trí index=2
                html_rows += f"<td class='text-wrap-column'>{val}</td>"
            else:
                html_rows += f"<td>{val}</td>"
        html_rows += "</tr>"
        
    html_header_4 = "<tr>"
    for idx, r4 in enumerate(row_4):
        r4_display = "" if r4.lower().startswith("unnamed:") else r4
        if idx == 2:
            html_header_4 += f"<th class='text-wrap-column'>{r4_display}</th>"
        else:
            html_header_4 += f"<th>{r4_display}</th>"
    html_header_4 += "</tr>"
    
    html_header_5 = "<tr>"
    for idx, r5 in enumerate(row_5):
        r5_display = "" if r5.lower().startswith("unnamed:") else r5
        if idx == 2:
            html_header_5 += f"<th class='text-wrap-column'>{r5_display}</th>"
        else:
            html_header_5 += f"<th>{r5_display}</th>"
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
        tbody tr:hover {{
            background-color: #f8f9fa;
        }}
    </style>
    <div class="table-container">
        <table>
            <thead>
                {html_header_4}
                {html_header_5}
            </thead>
            <tbody>
                {html_rows}
            </tbody>
        </table>
    </div>
    """
    
    st.components.v1.html(table_html, height=520, scrolling=False)
    
    # Nút tải CSV
    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 Tải danh sách đã lọc về máy (.CSV)",
        data=csv,
        file_name="dskh_da_loc.csv",
        mime="text/csv",
    )
    
    # --- PHẦN 3: THỐNG KÊ ---
    st.markdown("---") 
    st.subheader("📊 Khu vực tính tổng dữ liệu sau khi lọc")
    
    col_thong_ke_1, col_thong_ke_2 = st.columns(2)
    with col_thong_ke_1:
        st.metric(label="Tổng số dòng dữ liệu lọc được", value=f"{len(filtered_df)} dòng")
    with col_thong_ke_2:
        if COT_TINH_TONG in filtered_df.columns:
            solieu_so = pd.to_numeric(filtered_df[COT_TINH_TONG], errors='coerce').fillna(0)
            tong_gia_tri = solieu_so.sum()
            st.metric(label=f"Tổng cộng của cột [{COT_TINH_TONG}]", value=f"{tong_gia_tri:,.0f}")
        else:
            st.info(f"💡 Để tính tổng số tiền/doanh số, hãy nhập đúng tên cột ở dòng
