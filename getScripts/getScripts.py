import os
import pandas as pd
from openpyxl import Workbook

# Tạo một đối tượng Workbook
workbook = Workbook()
# Lấy active worksheet
worksheet = workbook.active

# Đường dẫn tới thư mục chứa các file .md
dir_path = '/Users/tr1pt3n/Developer/LOTL/LOLBAS/_lolbas'

# Tạo một list để lưu dữ liệu của các file .md
data = []

# Duyệt qua các tệp tin .md trong thư mục và các thư mục con
for root, dirs, files in os.walk(dir_path):
    for file_name in files:
        if file_name.endswith('.md'):
            # Đường dẫn đầy đủ đến tệp tin .md
            file_path = os.path.join(root, file_name)

            # Đọc nội dung của tệp tin .md
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

                # Tìm chỉ số của dòng chứa "Full_Path:"
                full_path_index = None
                for i, line in enumerate(lines):
                    if line.strip() == "Full_Path:":
                        full_path_index = i
                        break

                # Tìm chỉ số của dòng chứa "Code_Sample:"
                code_sample_index = None
                for i, line in enumerate(lines):
                    if line.strip() == "Code_Sample:":
                        code_sample_index = i
                        break

                # Tìm chỉ số của dòng chứa "Detection:"
                detection_index = None
                for i, line in enumerate(lines):
                    if line.strip() == "Detection:":
                        detection_index = i
                        break

                # Lấy nội dung từ "Full_Path:" tới "Code_Sample:" hoặc từ "Full_Path:" tới "Detection:"
                if full_path_index is not None and (code_sample_index is not None or detection_index is not None):
                    full_path_content = file_name
                    detection_content = ''
                    for line in lines[full_path_index + 1:]:
                        line = line.strip()
                        if line == '':
                            continue
                        if line.startswith('Full_Path:'):
                            break
                        if code_sample_index is not None and line.startswith('Code_Sample:'):
                            break
                        if detection_index is not None and line.startswith('Detection:'):
                            break
                        detection_content += line + ' '

                    data.append([full_path_content, detection_content])

# Tạo một DataFrame từ list data
df = pd.DataFrame(data, columns=['File_Name', 'Content'])

# In nội dung của DataFrame
print(df)

# Ghi DataFrame vào file Excel
df.to_excel('/Users/tr1pt3n/Developer/LOTL/scripts.xlsx', index=False)