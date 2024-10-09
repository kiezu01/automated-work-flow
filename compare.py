import os
import numpy as np
import pandas as pd

# Đường dẫn đến thư mục chứa các file Excel
dir_path = os.path.dirname(os.path.realpath(__file__))
data_path = os.path.join(dir_path, 'Bank Statement')

# Danh sách để lưu trữ các DataFrame
data_frames = []

# Đọc file All data.xlsx và Payments.xlsx và lưu vào danh sách
files = ['All data.xlsx', 'Payments.xlsx']
for file in files:
    df = pd.read_excel(os.path.join(data_path, file), engine='openpyxl')
    data_frames.append(df)

# data_frames[0] là All data, data_frames[1] là Payments
all_data, payments = data_frames[0], data_frames[1]

# So sánh cột B của 2 file, nếu Payments chưa có thì lưu vào file temp
unmatched_data = all_data[~all_data.iloc[:, 1].isin(payments.iloc[:, 1])]
unmatched_data.to_excel(os.path.join(data_path, 'temp.xlsx'), index=False)

# Đọc lại file temp.xlsx và xử lý trùng lặp
temp_data = pd.read_excel(os.path.join(data_path, 'temp.xlsx'), engine='openpyxl')
temp_data.iloc[:, 1] = temp_data.iloc[:, 1].astype(str)
temp_data.drop_duplicates(subset=temp_data.columns[1], keep='first', inplace=True)

# Chuyển đổi ngày tháng từ yyyy-mm-dd sang dd/mm/yyyy nếu cột đang ở định dạng chuỗi
def format_date(date_str):
    try:
        year, month, day = date_str.split('-')
        return f'{year}-{month}-{day}'
    except:
        return date_str  # Nếu gặp lỗi, trả về nguyên bản

temp_data.iloc[:, 0] = temp_data.iloc[:, 0].apply(format_date)

# Chuyển cột B sang kiểu số và lọc dữ liệu
temp_data.iloc[:, 1] = pd.to_numeric(temp_data.iloc[:, 1], errors='coerce')

# Lọc các giá trị lớn hơn 9588 và nhỏ hơn 20000 ở cột B và lọc các giá trị lớn hơn 9994093242781086
filtered_data = temp_data[((temp_data.iloc[:, 1] > 9588) & (temp_data.iloc[:, 1] < 20000)) | 
                          (temp_data.iloc[:, 1] > 9994093242750022)]

# chuyen cot B sang kieu chuoi
filtered_data.iloc[:, 1] = filtered_data.iloc[:, 1].astype(str)
filtered_data.to_excel(os.path.join(data_path, 'temp.xlsx'), index=False)

# Mở file kết quả
os.startfile(os.path.join(data_path, 'temp.xlsx'))
