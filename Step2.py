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
    df.iloc[:, 1] = df.iloc[:, 1].astype(str)
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

# filtered_data khong null
filtered_data = temp_data[~temp_data.iloc[:, 1].isnull()]
# kiem tra them 1 lan nua cot B voi du lieu trong Payments xem co trung nhau khong, neu trung thi xoa
filtered_data = filtered_data[~filtered_data.iloc[:, 1].isin(payments.iloc[:, 1])]
# chuyen cot B sang kieu chuoi
filtered_data.iloc[:, 1] = filtered_data.iloc[:, 1].astype(str)
filtered_data.to_excel(os.path.join(data_path, 'temp.xlsx'), index=False)

if filtered_data.empty:
    print('Không có dữ liệu mới được thêm vào')
else: 
    os.startfile(os.path.join(data_path, 'temp.xlsx'))


