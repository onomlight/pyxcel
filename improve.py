import pandas as pd
from tkinter import messagebox
# 엑셀 파일 경로
file_path = 'bigtest.xlsx'
output_file_path = 'new__새로운_파일.xlsx'

# 엑셀 파일 불러오기
df = pd.read_excel(file_path)

# 'O'열 초기화
df['O'] = ''

# 조건과 매칭될 값을 정의한 딕셔너리
mapping_dict = {
('메르헨트 대용량 샴푸 1500ml 베이비파우더 약산성 퍼퓸 두피 지성 향기좋은 미용실 사춘기 청소년','[메르헨트 옵션 선택]: 5.티트리&로즈 샴푸 1.5L_머스크향'):'♡4W_OMHC1038_메헨 티트리 샴푸 화이트머스크 1.5L 1개',
('[메르헨트] 메르헨트 퍼퓸 대용량 샴푸 1000ml  x 1개 약산성 퍼퓸 두피 미용실 향기좋은 지성 건성','1) 퍼퓸 화이트솝'):'♡4WC_OMHC1062_메헨 퍼퓸 샴푸 화이트솝 1L 1개',
('1+1 메르헨트 대용량 디퓨저 500ml 시트리코 실내방향제 화장실 사무실 인테리어 아로마','1) 향선택: 메르헨트 디퓨저 500ml 2개 런드리룸'):'',

('1+1 메르헨트 대용량 디퓨저 500ml 시트리코 실내방향제 화장실 사무실 인테리어 아로마','1) 향선택: 메르헨트 디퓨저 500ml 2개 시트리코'):'',

('1+1 메르헨트 대용량 디퓨저 500ml 시트리코 실내방향제 화장실 사무실 인테리어 아로마','1) 향선택: 메르헨트 디퓨저 500ml 2개 시트리코'):'',
('1+1 메르헨트 대용량 디퓨저 500ml 시트리코 실내방향제 화장실 사무실 인테리어 아로마','1) 향선택: 메르헨트 디퓨저 500ml 2개 준포레스트'):'',

('1+1 섬유향수 섬유탈취제 250ml 7종향 룸스프레이 드레스퍼퓸','1+1 섬유향수 250ml 화이트머스크'):'',
    # 추가적인 조건에 따라 더 많은 값들을 설정할 수 있습니다.
    # , ...
}

highlight_indices = []
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
    else:
        highlight_indices.append(index)

# 변경된 데이터프레임을 새로운 엑셀 파일로 저장
# df_style = df.style
# df_style.applymap(lambda x: f"background-color: {'lightgrey'}", subset=pd.IndexSlice[highlight_indices, ['N']])
# df_style.to_excel(output_file_path, index=False)

# print("작업이 완료되었습니다.")

# if 'N' in df.columns:
#     df_style = df.style
#     df_style.applymap(lambda x: f"background-color: {'lightgrey'}", subset=pd.IndexSlice[highlight_indices, ['N']])
#     df_style.to_excel(output_file_path, index=False)
#     print("작업이 완료되었습니다.")
# else:
#     raise ValueError("No 'N' column found in the dataframe. Columns: {df.columns}")

# 'N' 열이 존재하는 경우에만 스타일 적용
# if any('N' in col.upper() for col in df.columns):
#     df_style = df.style
#     df_style.applymap(lambda x: f"background-color: {'lightgrey'}", subset=pd.IndexSlice[highlight_indices, [col for col in df.columns if 'N' in col.upper()]])
#     df_style.to_excel(output_file_path, index=False)
#     print("작업이 완료되었습니다.")
# else:
#     raise ValueError("No 'N' column found in the dataframe. Columns: {df.columns}")
#pip install Jinja2


# 'N' 열이 존재하는 경우에만 스타일 적용
if any('N' in col.upper() for col in df.columns):
    df_style = df.style
    df_style = df_style.applymap(lambda x: f"background-color: {'lightgrey'}", subset=pd.IndexSlice[highlight_indices, [col for col in df.columns if 'N' in col.upper()]])
    df_style.to_excel(output_file_path, index=False)
    print("작업이 완료되었습니다.")

    # 메시지 박스 표시
    messagebox.showinfo("완료", "작업이 완료되었습니다!")


else:
    raise ValueError("No 'N' column found in the dataframe. Columns: {df.columns}")
