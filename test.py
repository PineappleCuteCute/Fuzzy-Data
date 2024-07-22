import pandas as pd

# Đọc dữ liệu từ file Excel
file_path = 'Data.xlsx'
sheet_name = 'heart'
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Tính số lượng nhãn 0 và 1 trong cột target
count_label_0 = df[df['target'] == 0].shape[0]
count_label_1 = df[df['target'] == 1].shape[0]

print(f'Số lượng nhãn 0: {count_label_0}')
print(f'Số lượng nhãn 1: {count_label_1}')
