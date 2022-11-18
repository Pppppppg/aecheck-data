import json
import os
from openpyxl.worksheet.worksheet import Worksheet

def write_json(path: str, json_data):
    """json 값을 해당 경로에 파일로 저장하는 함수"""
    with open(path, 'w+', encoding='utf-8') as f:
        json.dump(json_data, f, indent="\t", ensure_ascii=False)

def make_folder(path: str):
    """경로에 해당하는 폴더가 없을 경우 생성하는 함수"""
    if not os.path.exists(path):
        os.makedirs(path)


def add_data_to_sheet(tag: str, sheet: Worksheet, data_list):
    """
    새로운 데이터를 번역 시트에 추가하는 함수
    """
    
    # 번역 시트의 첫 열은 한국어
    sheet_list = [i[0].value for i in sheet.iter_rows(min_row=2)]
    
    for key in list(set(data_list)):
        if key is not None and key not in sheet_list:
            print("새로 등록, 번역 확인 필요 : {} - {}".format(tag, key))
            sheet.append([key])