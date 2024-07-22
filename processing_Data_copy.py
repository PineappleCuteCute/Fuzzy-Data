import pandas as pd

# Đọc tệp Excel
file_path = '/Users/daomanh/Desktop/Fuzzy Data/Data copy.xlsx'
df = pd.read_excel(file_path, sheet_name='FuzzyData_heart')

# Hàm để chuyển đổi các giá trị
def convert_value(value):
    mapping = {
        0: 'Low',
        1: 'Medium',
        2: 'High',
        3: 'Very High',
        4: 'Extremely High'
    }
    return mapping.get(value, value)

# Áp dụng hàm chuyển đổi cho các cột cần thiết
columns_to_convert = ['cp', 'exang', 'slope', 'ca', 'thal']
for column in columns_to_convert:
    df[column] = df[column].apply(convert_value)

# Lưu lại tệp Excel với các giá trị đã chuyển đổi
output_file_path = '/Users/daomanh/Desktop/Fuzzy Data/Data copy.xlsx'
df.to_excel(output_file_path, index=False, sheet_name='FuzzyData_heart')

print(f"Đã lưu tệp đã chuyển đổi tại {output_file_path}")
