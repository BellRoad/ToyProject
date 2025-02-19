'''
merge_sales.py에서 저장된 csv파일을 순서대로 하나의 csv파일에 저장하는 프로그램
'''

import pandas as pd
import glob

# CSV 파일들이 있는 디렉토리 경로
path = './'  # 현재 디렉토리를 가정. 필요시 변경하세요.

# 모든 CSV 파일 목록 가져오기
all_files = glob.glob(path + "2024*.csv")

# 빈 리스트 생성
li = []

# 각 파일을 읽어서 리스트에 추가
for filename in all_files:
    df = pd.read_csv(filename, encoding='ansi')
    li.append(df)

# 모든 데이터프레임을 하나로 합치기
combined_df = pd.concat(li, axis=0, ignore_index=True)

# 결과를 새 CSV 파일로 저장
combined_df.to_csv("combined_2024.csv", index=False, encoding='utf-8-sig')

print("모든 CSV 파일이 성공적으로 통합되었습니다.")
