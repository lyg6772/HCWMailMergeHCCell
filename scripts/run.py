import win32com.client
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

# read cell file 

def get_cell_file_path():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="한셀파일 선택", filetypes=[("한셀 파일", "*.cell"), ("모든 파일", "*.*")])
    return file_path

def get_hwp_template_file_path():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="한글 템플릿 파일 선택", filetypes=[("한글 파일", "*.hwp"), ("모든 파일", "*.*")])
    return file_path

cell_file_path = get_cell_file_path()
hwp_file_path = get_hwp_template_file_path()

# cell_file_path = 'C:/Users/빅스피오나/Desktop/출석근태신고서_데이터.cell'
# hwp_file_path = 'C:/Users/빅스피오나/Desktop/출석 근태 신고서_양식.hwp'
#cell -> xslx 저장
xlsx_file_path = cell_file_path.replace('cell', 'xlsx')
hcell = win32com.client.Dispatch("HCell.Application")
hcell.Visible = False
cell_data = hcell.Workbooks.Open(cell_file_path)
if os.path.exists(xlsx_file_path):
    os.remove(xlsx_file_path)
cell_data.SaveAs(xlsx_file_path, FileFormat=51)
hcell.Quit()

# 엑셀파일 open
df = pd.read_excel(xlsx_file_path)

data_list = df.to_dict(orient="records")
current_dir = os.getcwd()
idx = 0
result_file_name = hwp_file_path.split('/')[-1].split('_')[0]
try:
# 2. 한글을 제어해 메일머지를 수행

    # 한글 파일 열기 (메일머지 템플릿)

    for data in data_list:
        # 필드 데이터 설정
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.Open(hwp_file_path)
        # hwp.XHwpWindows.Item(0).Visible = True
        key_list = [str(x) for x in data.keys()]
        value_list = [str(x) for x in data.values()]
        hwp.PutFieldText(Field='\x02'.join(key_list), Text='\x02'.join(value_list))

        hwp.SaveAs(f"{current_dir}/{result_file_name}_{idx}.hwp")
        hwp.Quit()
        idx = idx+1
    # 저장 및 종료
except Exception as e:
    print(e)
    if hwp:
        hwp.Quit()
