import pandas as pd

# 엑셀 파일 경로
file_path = 'openyour.xlsx'
output_file_path = 'loadsave.xlsx'

# 엑셀 파일 불러오기
df = pd.read_excel(file_path)

# 'O' 열 초기화
df['O'] = ''

# '상품명'과 '옵션명' 값을 튜플로 묶어 딕셔너리로 만들기
mapping_dict = dict(zip(zip(df['상품명'], df['옵션명']), df['O']))

# 'O' 값 입력
df['O'] = list(map(lambda row: mapping_dict.get((row['상품명'], row['옵션명']), ''), df[['상품명', '옵션명']].to_records(index=False)))

# 딕셔너리를 직접 엑셀 파일로 저장
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    # 데이터프레임을 엑셀 파일에 쓰기
    df.to_excel(writer, index=False)

    # 딕셔너리를 새로운 시트로 추가
    mapping_df = pd.DataFrame(list(mapping_dict.keys()), columns=['M', 'N'])
    mapping_df['O'] = ''
    mapping_df.to_excel(writer, sheet_name='Mapping', index=False)

print("작업이 완료되었습니다.")
