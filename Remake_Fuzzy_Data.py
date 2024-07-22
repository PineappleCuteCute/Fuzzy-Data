import pandas as pd

# Đọc dữ liệu từ sheet 'heart' trong file Excel
df = pd.read_excel('Data.xlsx', sheet_name='heart')

# Định nghĩa hàm chuyển đổi giá trị cho từng thuộc tính
def convert_age(value):
    if value < 40:
        return 'Low'
    elif 40 <= value <= 60:
        return 'Medium'
    elif 61 <= value <= 80:
        return 'High'

def convert_trestbps(value):
    if value < 90:
        return 'Low'
    elif 90 <= value <= 120:
        return 'Medium'
    elif value > 120:
        return 'High'

def convert_chol(value):
    if value < 200:
        return 'Medium'
    elif 200 <= value <= 240:
        return 'High'
    elif value > 240:
        return 'Very High'

def convert_fbs(value):
    if value == 0:
        return 'Medium'
    else:
        return 'High'

def convert_restecg(value):
    if value == 0:
        return 'Medium'
    elif value == 1:
        return 'High'
    else:
        return 'Very High'

def convert_thalach(value):
    if value < 100:
        return 'Medium'
    else:
        return 'High'

def convert_oldpeak(value):
    if value < 2:
        return 'Low'
    else:
        return 'High'

# Áp dụng hàm chuyển đổi cho từng cột trong DataFrame
df['age'] = df['age'].apply(convert_age)
df['trestbps'] = df['trestbps'].apply(convert_trestbps)
df['chol'] = df['chol'].apply(convert_chol)
df['fbs'] = df['fbs'].apply(convert_fbs)
df['restecg'] = df['restecg'].apply(convert_restecg)
df['thalach'] = df['thalach'].apply(convert_thalach)
df['oldpeak'] = df['oldpeak'].apply(convert_oldpeak)

# Ghi kết quả vào sheet mới trong file Excel với chế độ 'replace'
with pd.ExcelWriter('Data.xlsx', engine='openpyxl', mode='a') as writer:
    df.to_excel(writer, sheet_name='FuzzyData', index=False)