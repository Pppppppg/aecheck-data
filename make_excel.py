from openpyxl import load_workbook
from function import add_data_to_sheet
from config import FILE_NAME

print("Load Excel...")
workbook = load_workbook(FILE_NAME, data_only=True)


# 1. 캐릭터 이름 번역
# 모든 캐릭터 이름들을 모아서 중복을 제거한 배열을 생성
per_sheet = workbook["퍼스널리티"]
char_list = list(set([i[0].value for i in per_sheet.iter_rows(min_row=2)]))
print("대상 캐릭터", len(char_list), "명")

add_data_to_sheet("캐릭터", workbook["캐릭번역"], char_list)


# 2. 직업서 번역
char_sheet = workbook["캐릭터"]
book_list = list(set([i[12].value for i in char_sheet.iter_rows(min_row=2)]))

add_data_to_sheet("직업서", workbook["캐릭번역"], book_list)


# 3. 특성 번역
per_list = []
per_sheet = workbook["퍼스널리티"]
for row in per_sheet.iter_rows(min_row=2):
    per_list += row[2].value.split(",")

add_data_to_sheet("퍼스널리티", workbook["특성번역"], per_list)


# Save
workbook.save(FILE_NAME)
print("Save Excel. COMPLETE")