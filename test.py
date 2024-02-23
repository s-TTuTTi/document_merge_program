import pandas as pd

def merge_xlsx(file_list, output_file):
    # 파일 목록에서 DataFrame을 담을 리스트 생성
    df_list = []

    for file in file_list:
        # 파일 목록에서 하나씩 파일을 가져와 DataFrame 형식으로 저장
        file_df = pd.read_excel(file)
        df_list.append(file_df)

    # pd.concat() 함수를 사용하여 데이터프레임들을 병합
    merged_df = pd.concat(df_list, ignore_index=True)

    # 결과를 Excel 파일로 저장
    merged_df.to_excel(output_file, index=False)

# 사용 예시
file_list = ['xlsx_sample/test1.xlsx', 'xlsx_sample/test2.xlsx']
merge_xlsx(file_list, 'xlsx_sample/output.xlsx')