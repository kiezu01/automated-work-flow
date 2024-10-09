import re
import pandas as pd
import os
import numpy as np
import logging

# Thiết lập logging để theo dõi quá trình
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Đường dẫn đến thư mục chứa các file Excel
def get_directory_path():
    return os.path.dirname(os.path.realpath(__file__))

# Đọc danh sách các file có đuôi là .xlsx
def get_excel_files(data_path):
    return [f for f in os.listdir(data_path) if f.endswith('.xlsx')]

# Kiểm tra và đọc file Excel, bỏ qua các file lỗi
def read_excel_files(data_files, data_path):
    data = []
    for file in data_files:
        file_path = os.path.join(data_path, file)
        try:
            df = pd.read_excel(file_path, header=None, engine='openpyxl')
            data.append(df)
            logging.info(f"Đã đọc file {file}")
        except ValueError as e:
            logging.error(f"Bỏ qua file {file} do lỗi: {e}")
    return data

# Chuẩn hóa định dạng ngày tháng
def normalize_date_format(data, date_column_index, original_format, target_format='%Y-%m-%d'):
    for i in range(len(data)):
        if data[i].shape[1] > date_column_index:  # Kiểm tra có cột ngày tháng không
            try:
                data[i].iloc[:, date_column_index] = pd.to_datetime(
                    data[i].iloc[:, date_column_index],
                    format=original_format,
                    errors='coerce'
                ).dt.strftime(target_format)
            except Exception as e:
                logging.error(f"Lỗi khi chuẩn hóa ngày tháng: {e}")
    return data

# Xử lý dữ liệu
def process_data(data):
    # Chuẩn hóa cột ngày tháng (giả sử cột 0 là cột ngày tháng)
    data = normalize_date_format(data, date_column_index=0, original_format='%d/%m/%Y %I:%M:%S %p')

    for i in range(len(data)):
        if data[i].shape[1] > 4:  # Kiểm tra có đủ cột trước khi truy cập
            # Xử lý cột E
            data[i].iloc[:, 4] = data[i].iloc[:, 4].replace(np.nan, 'X')
            data[i] = data[i][data[i].iloc[:, 4] != 'X']
        
        if data[i].shape[1] > 1:
            # Xử lý cột B
            data[i].iloc[:, 1] = data[i].iloc[:, 1].astype(str)
            data[i].drop_duplicates(subset=data[i].columns[1], keep='first', inplace=True)
            data[i] = data[i][data[i].iloc[:, 1].apply(lambda x: str(x).isdigit())]
        
        if data[i].shape[1] > 2:
            # Xử lý cột C
            data[i].iloc[:, 2] = data[i].iloc[:, 2].str.replace(' ', '')
            data[i].iloc[:, 2] = data[i].iloc[:, 2].str.upper()

            # Thêm cột số hợp đồng
            data[i][data[i].shape[1]] = data[i].iloc[:, 2].apply(extract_agreement_number)
    
    return pd.concat(data)

# Hàm trích xuất số hợp đồng từ chuỗi
def extract_agreement_number(s):
    patterns = [r'GDP\s?-?\s?(\d+)', r'DP\s?-?\s?(\d+)', r'GMNV\s?-?\s?(\d+)', r'MNV\s?-?\s?(\d+)', r'GDL\s?-?\s?(\d+)', r'DL\s?-?\s?(\d+)']
    for pattern in patterns:
        match = re.search(pattern, s)
        if match:
            return match.group(1)
    return np.nan

def save_and_clean_data(merged_data, output_path):
    # Xử lý và loại bỏ trùng lặp ngay trước khi lưu để tránh việc lưu lại nhiều lần
    merged_data.iloc[:, 1] = merged_data.iloc[:, 1].astype(str)
    merged_data.drop_duplicates(subset=merged_data.columns[1], keep='first', inplace=True)

    # Lưu vào Excel
    merged_data.to_excel(output_path, index=False)
    logging.info("Dữ liệu đã được hợp nhất và làm sạch, lưu vào file All data.xlsx")

def main():
    # Đường dẫn
    dir_path = get_directory_path()
    data_path = os.path.join(dir_path, 'data')
    output_dir = os.path.join(dir_path, 'Bank Statement')

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Đọc các file Excel
    data_files = get_excel_files(data_path)
    data = read_excel_files(data_files, data_path)

    # Xử lý dữ liệu
    merged_data = process_data(data)

    # Lưu dữ liệu vào file Excel
    output_path = os.path.join(output_dir, 'All data.xlsx')
    save_and_clean_data(merged_data, output_path)

if __name__ == "__main__":
    main()
