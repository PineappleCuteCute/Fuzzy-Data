import pandas as pd
from sklearn.model_selection import train_test_split

# Đọc dữ liệu từ sheet 'FuzzyData_heart' trong file Excel
df = pd.read_excel('Data.xlsx', sheet_name='Fuzzy_Data_heart')

# Chia dữ liệu thành tập train và tập test với tỉ lệ 80-20
train_df, test_df = train_test_split(df, test_size=0.3, random_state=42)

# Lưu tập train vào sheet 'FuzzyData'
with pd.ExcelWriter('Data.xlsx', engine='openpyxl', mode='a') as writer:
    train_df.to_excel(writer, sheet_name='FuzzyData2', index=False)

# Lưu tập test vào sheet 'FuzzyTest'
with pd.ExcelWriter('Data.xlsx', engine='openpyxl', mode='a') as writer:
    test_df.to_excel(writer, sheet_name='FuzzyTest2', index=False)
