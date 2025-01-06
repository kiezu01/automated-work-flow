import re
import pandas as pd
import os
import numpy as np
import logging
import warnings

# Loai bo warning
warnings.filterwarnings("ignore", category=UserWarning)

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
def normalize_date_format(data, date_column_index, target_format='%Y-%m-%d'):
    for i in range(len(data)):
        if data[i].shape[1] > date_column_index:  # Kiểm tra có cột ngày tháng không
            try:
                # Đọc ngày tháng và chuyển thành chuỗi
                data[i].iloc[:, date_column_index] = pd.to_datetime(
                    data[i].iloc[:, date_column_index], 
                    errors='coerce',  # Nếu không chuyển đổi được thì sẽ thành NaT
                    dayfirst=True  # Đảm bảo rằng ngày là dd/mm/yyyy
                )
                
                # Thêm bước để chỉ lấy phần ngày (loại bỏ giờ)
                data[i].iloc[:, date_column_index] = data[i].iloc[:, date_column_index].astype(str).str.split(' ').str[0]
                data[i].iloc[:, date_column_index] = data[i].iloc[:, date_column_index].str.replace(r'\s.*', '')

                # Loại bỏ các dòng có NaT trong cột ngày tháng
                data[i] = data[i].dropna(subset=[data[i].columns[date_column_index]])

            except Exception as e:
                logging.error(f"Lỗi khi chuẩn hóa ngày tháng: {e}")
    return data



def process_data(data):
    # Chuẩn hóa cột ngày tháng (giả sử cột 0 là cột ngày tháng)
    data = normalize_date_format(data, date_column_index=0, target_format='%Y-%m-%d')

    for i in range(len(data)):
        if data[i].shape[1] > 4:  # Kiểm tra có đủ cột trước khi truy cập
            # Xử lý cột E
            data[i].iloc[:, 4] = data[i].iloc[:, 4].replace(np.nan, 'X')
            data[i] = data[i][data[i].iloc[:, 4] != 'X']
        
        if data[i].shape[1] > 1:
            # Chuyển cột B thành chuỗi và loại bỏ các giá trị trùng lặp
            data[i].iloc[:, 1] = data[i].iloc[:, 1].astype(str)
            data[i].drop_duplicates(subset=data[i].columns[1], keep='first', inplace=True)
        
        if data[i].shape[1] > 2:
            # Xử lý cột C
            data[i].iloc[:, 2] = data[i].iloc[:, 2].str.replace(' ', '')
            data[i].iloc[:, 2] = data[i].iloc[:, 2].str.upper()

            # Thêm cột số hợp đồng
            data[i][data[i].shape[1]] = data[i].iloc[:, 2].apply(extract_agreement_number)
    
    return pd.concat(data)



# Hàm trích xuất số hợp đồng từ chuỗi
def extract_agreement_number(s):
    patterns = [r'GDP\s?-?\s?(\d+)', r'DP\s?-?\s?(\d+)', r'GMNV\s?-?\s?(\d+)', r'MNV\s?-?\s?(\d+)'
                r'GDL\s?-?\s?(\d+)', r'DL\s?-?\s?(\d+)', r'HD\s?-?\s?(\d+)', r'SHD\s?-?\s?(\d+)'
                r'GOCR\s?-?\s?(\d+)',r'DPL\s?-?\s?(\d+)', r'GDPL\s?-?\s?(\d+)']
    for pattern in patterns:
        match = re.search(pattern, s)
        if match:
            return match.group(1)
    return np.nan


# Lưu dữ liệu vào file Excel và loại bỏ các bản sao từ cột B
def save_and_clean_data(merged_data, output_path):
    merged_data.to_excel(output_path, index=False)

    # Đọc lại và xử lý file vừa lưu để xóa các dòng trùng lặp ở cột B
    df = pd.read_excel(output_path, engine='openpyxl')
    df.iloc[:, 1] = df.iloc[:, 1].astype(str)
    df.drop_duplicates(subset=df.columns[1], keep='first', inplace=True)
    
    df.to_excel(output_path, index=False)
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

    # xoa cac hang o cot ngay thang neu khong dung dinh dang yyyy-mm-dd
    merged_data = merged_data[merged_data.iloc[:, 0].str.match(r'\d{4}-\d{2}-\d{2}')]
    
    # Lưu dữ liệu vào file Excel
    output_path = os.path.join(output_dir, 'All data.xlsx')
    save_and_clean_data(merged_data, output_path)
    # os.startfile(output_path)

if __name__ == "__main__":
    main()
