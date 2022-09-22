from openpyxl.styles import PatternFill
import json
import os

# 변경시 셀 색상 변경 용도
changeFill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')


"""
file, folder CRUD
"""
def read_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        res = json.load(f)
    return res  

def write_json(path: str, data):
    with open(path, 'w+', encoding='utf-8') as f:
        json.dump(data, f, indent="\t", ensure_ascii=False)

def make_folder(path: str):
    if not os.path.exists(path):
        os.makedirs(path)



# 새로운 데이터를 번역 시트에 추가하는 함수
def add_data_to_sheet(tag: str, sheet, data_list):
    # 1. 번역 시트의 첫 열은 한국어입니다.
    sheet_list = [i[0].value for i in sheet.iter_rows(min_row=2)]
    
    for key in list(set(data_list)):
        if key is not None and key not in sheet_list:
            print("새로 등록, 번역 확인 필요 : {} - {}".format(tag, key))
            sheet.append([key])