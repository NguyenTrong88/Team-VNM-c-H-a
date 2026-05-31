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
    selected_filter_cols = st.sidebar.multiselect("Chọn các cột bạn muốn lọc chi tiết:", options=filter_options, default=filter_options[:2] if len(filter
