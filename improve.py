import pandas as pd

# 엑셀 파일 경로
file_path = '파일.xlsx'
output_file_path = '새로운_파일.xlsx'

# 엑셀 파일 불러오기
df = pd.read_excel(file_path)

# 'O'열 초기화
df['O'] = ''

# 조건과 매칭될 값을 정의한 딕셔너리
mapping_dict = {
    ('A', 'B'): 'A1',
    ('A', 'C'): 'A2',
    ('A', 'F'): 'FA'
    # 추가적인 조건에 따라 더 많은 값들을 설정할 수 있습니다.
    # , ...
}

# M1과 N1 열의 데이터에 따라 다른 값으로 변경 (첫 번째 행부터 비교 시작)
for index in range(len(df)):
    current_row = df.iloc[index]

    # '상품명'에 해당하는 열 찾기
    m_column = [col for col in df.columns if '상품명' in col or 'product' in col]
    if not m_column:
        raise ValueError(f"No '상품명' or 'product' column found in the dataframe. Columns: {df.columns}")

    # '옵션명'에 해당하는 열 찾기
    n_column = [col for col in df.columns if '옵션명' in col or 'option' in col]
    if not n_column:
        raise ValueError(f"No '옵션명' or 'option' column found in the dataframe. Columns: {df.columns}")

    # 조건에 따라 'O' 열 자동으로 입력
    key = (current_row[m_column[0]], current_row[n_column[0]])
    if key in mapping_dict:
        df.at[index, 'O'] = mapping_dict[key]

# 변경된 데이터프레임을 새로운 엑셀 파일로 저장
df.to_excel(output_file_path, index=False)

print("작업이 완료되었습니다.")
