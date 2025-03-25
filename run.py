import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from io import BytesIO
import xlsxwriter

# Bước 1: Kết nối đến Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r'D:\PAST\STREAMLIT\credentials.json', scope)
client = gspread.authorize(creds)
spreadsheet = client.open('streamlit')
sheet = spreadsheet.get_worksheet(0)  # Trang tính đầu tiên (chỉ số 0)
sheet2 = spreadsheet.get_worksheet(1)  # Trang tính thứ hai (chỉ số 1)

# Bước 2: Hiển thị dữ liệu từ Google Sheets trong Streamlit
st.title('Bộ danh mục chuẩn')
df_sieuthi_danhmuc = pd.DataFrame()
df_danhmuc_size = pd.DataFrame()

def Transform_SizeDanhMuc(data_frame):
    # Sử dụng pd.melt để unpivot các cột 'Danh mục 2500', 'Danh mục 2000', ...
    df_unpivot = pd.melt(data_frame, 
                        id_vars=['Mã cơ sở', 'TÊN HỆ THỐNG'], 
                        value_vars=['Danh mục 2500 (final)', 'Danh mục 2000 (final)', 'Danh mục 1500 (final)', 'Danh mục 1000 (final)'],
                        var_name='Size danh mục', 
                        value_name='Giá trị')
    # Nếu muốn loại bỏ '(final)' từ 'Size danh mục'
    df_unpivot['Size danh mục'] = df_unpivot['Size danh mục'].str.replace(' (final)', '', regex=False)
    df_unpivot['Size danh mục'] = df_unpivot['Size danh mục'].str.replace('Danh mục ', '', regex=False)
    return df_unpivot

file_danhmuc_size = st.file_uploader("Chọn file dữ liệu Danh mục theo Size siêu thị")
if file_danhmuc_size is not None:
    df_danhmuc_size = pd.read_excel(file_danhmuc_size)
    df_danhmuc_size = Transform_SizeDanhMuc(df_danhmuc_size)
    df_danhmuc_size = df_danhmuc_size[df_danhmuc_size['Giá trị'] == 1].drop(columns='Giá trị')
    df_danhmuc_size['Size danh mục'] = df_danhmuc_size['Size danh mục'].astype('int64')
    st.write("Load thành công")
    st.dataframe(df_danhmuc_size, height=210)

file_sieuthi_size = st.file_uploader("Chọn file dữ liệu Siêu thị")
if file_sieuthi_size is not None:
    df_sieuthi_danhmuc = pd.read_excel(file_sieuthi_size)
    st.write("Load thành công")
    st.dataframe(df_sieuthi_danhmuc, height=210)

try: df_danhmuc_chuan = pd.merge(df_sieuthi_danhmuc, df_danhmuc_size, on='Size danh mục', how='outer')
except: pass

# Hàm chuyển DataFrame thành tệp Excel
@st.cache_data
def convert_df_to_excel(df):
    output = BytesIO()  # Tạo một bộ nhớ đệm
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)  # Ghi DataFrame vào Excel
    processed_data = output.getvalue()  # Lấy dữ liệu từ bộ nhớ đệm
    return processed_data

# Chuyển DataFrame thành file Excel
try: excel_data = convert_df_to_excel(df_danhmuc_chuan)
except: pass

# Tạo nút tải xuống
try:
    st.download_button(
        label="Download bộ danh mục chuẩn",
        data=excel_data,
        file_name="Danh_muc_chuan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except: pass
