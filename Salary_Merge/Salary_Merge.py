'''
위하고의 월별 급여 내역 파일 "(주)동심-202401.xlsx"~"(주)동심-202412.xlsx"에서
사원명, 부서, 직급, 지급액계, 국민연금, 건강보험, 장기요양보험, 소득세, 지방소득세 칼럼만 선택하여
A열에 월을 입력하고 다음열에 차례대로 위의 칼럼을 입력.
이후 "급여내역 2024 동심.xlsx"에 저장하는 프로그램
'''

import pandas as pd

# 결과를 저장할 빈 데이터프레임 초기화
final_df = pd.DataFrame()

# 파일명 패턴에 따라 반복 처리
for month in range(1, 13):
    # 파일명 생성
    file_name = f"(주)동심-2024{month:02d}.xlsx"
    
    # 엑셀 파일 읽기 (1, 2행을 MultiIndex로 처리)
    df = pd.read_excel(file_name, header=[0, 1])
    
    # 선택할 MultiIndex 열 정의
    selected_columns = [
        ('사원명', 'Unnamed: 1_level_1'),
        ('부서', 'Unnamed: 2_level_1'),
        ('직급', 'Unnamed: 3_level_1'),
        ('수당', '지급액계'),
        ('공제', '국민연금'),
        ('공제', '건강보험'),
        ('공제', '고용보험'),        
        ('공제', '장기요양보험료'),
        ('공제', '소득세'),
        ('공제', '지방소득세')
    ]
    
    # 필요한 열만 선택
    extracted_df = df[selected_columns]
    
    # 열 이름 간소화
    extracted_df.columns = ['사원명', '부서', '직급', '지급액계', '국민연금', '건강보험', '고용보험', '장기요양보험', '소득세', '지방소득세']
    
    # 마지막 행('합계') 제외
    extracted_df = extracted_df.iloc[:-1]
    
    # "월" 열 추가
    extracted_df.insert(0, '월', f"{month:02d}월")
    
    # 결과 데이터프레임에 추가
    final_df = pd.concat([final_df, extracted_df], ignore_index=True)

# 최종 결과를 새 엑셀 파일로 저장
output_file = "급여내역 2024 동심.xlsx"
final_df.to_excel(output_file, index=False)

print(f"파일이 저장되었습니다: {output_file}")
