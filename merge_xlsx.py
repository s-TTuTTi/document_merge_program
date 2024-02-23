from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from copy import copy


def copy_sheet_data(source_sheet, target_sheet, start_row=1):
    # 병합된 셀 정보 복사
    target_sheet.merged_cells = copy(source_sheet.merged_cells)
    print(target_sheet.merged_cells)
    target_sheet.merged_cells.add('Z6:Z7')
    print(target_sheet.merged_cells)
    # 각 열의 넓이를 복사
    for col_letter, column_dimensions in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[col_letter].width = column_dimensions.width

    # 각 행의 높이를 복사
    for row_letter, row_dimensions in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[row_letter].height = row_dimensions.height
    # 각 셀의 데이터 및 스타일 복사

    for row in source_sheet.iter_rows(min_row=1, max_row=source_sheet.max_row, min_col=1,
                                      max_col=source_sheet.max_column):
        for cell in row:
            # 새로운 셀 생성
            new_cell = target_sheet[get_column_letter(cell.column) + str(start_row + cell.row - 1)]

            # 셀 데이터 복사
            new_cell.value = cell.value

            # 셀 스타일 복사
            new_cell.font = copy(cell.font)
            new_cell.border = copy(cell.border)
            new_cell.fill = copy(cell.fill)
            new_cell.number_format = copy(cell.number_format)
            new_cell.protection = copy(cell.protection)
            new_cell.alignment = copy(cell.alignment)


# 원본 워크북 로드
wb = load_workbook(filename='xlsx_sample/test2.xlsx')
ws = wb["참여학사조직"]

twb = load_workbook(filename='xlsx_sample/test1.xlsx')
tws = twb["창원대"]

# 새로운 워크북 생성
newwb = Workbook()
newwb.remove(newwb.active)
newws = newwb.create_sheet("Page 1")

# 데이터 및 스타일 복사
copy_sheet_data(ws, newws)
copy_sheet_data(tws, newws, start_row=ws.max_row+2)

# 새로운 워크북 저장
newwb.save('newEXCEL.xlsx')

# 원본 워크북과 새로운 워크북 닫기
wb.close()
newwb.close()

# 원본 워크북과 새로운 워크북 닫기
wb.close()
newwb.close()