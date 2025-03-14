import pandas as pd
import glob
import re
import os

# 파일 패턴 정의
patterns = ['재고리스트(누계)*동심*.xlsx', '재고리스트(누계)*5월5일*.xlsx', '재고리스트(누계)*킨더*.xlsx']

for pattern in patterns:
    # 파일 목록 가져오기
    files = glob.glob(pattern)
    
    if not files:
        print(f"'{pattern}' 패턴에 해당하는 파일이 없습니다.")
        continue
    
    # 데이터프레임 리스트 초기화
    dfs = []
    date_ranges = set()  # 날짜 범위 저장
    
    # 파일 읽기 및 병합
    for file in files:
        # 파일 이름에서 날짜 추출 (yyyymmdd-yyyymmdd 형식)
        match = re.search(r'\d{8}-\d{8}', file)
        if match:
            date_ranges.add(match.group())  # 날짜 범위 추가
        
        df = pd.read_excel(file)
        dfs.append(df)
    
    merged_df = pd.concat(dfs, ignore_index=True)
    
    # 열 순서 재정의
    new_columns = [
        '품목코드', '품목명', '전재고 수량', '전재고 금액', '입고 수량', '현재고 수량', '현재고 금액',
        '출고 수량', '반품 수량', 'A/S 수량', '조정 수량', '가공 수량', '입고 금액', '출고 금액', '반품 금액', 'A/S 금액', '조정 금액', '가공 금액'
    ]
    merged_df = merged_df.reindex(columns=new_columns)
    
    # 새 열 추가 (출고 합계)
    merged_df.insert(
        loc=merged_df.columns.get_loc('현재고 수량'),
        column='출고 합계',
        value=merged_df['출고 수량'] + merged_df['반품 수량'] + merged_df['A/S 수량'] + merged_df['조정 수량'] + merged_df['가공 수량']
    )
    
    # 품목명 기준 그룹화 및 합계 계산
    numeric_cols = merged_df.select_dtypes(include='number').columns.tolist()
    grouped_df = (merged_df.groupby('품목명', sort=False).agg({
        '품목코드': 'first',
        **{col: 'sum' for col in numeric_cols if col != '품목코드'}
    })
    .sort_index(ascending=True)
    .reset_index())
    
    # 출력 파일명 생성 (날짜 범위 포함)
    date_range_str = '_'.join(sorted(date_ranges))
    if '동심' in pattern:
        output_file = f'재고리스트(누계) 동심 통합 {date_range_str}.xlsx'
    elif '5월5일' in pattern:
        output_file = f'재고리스트(누계) 5월5일 통합 {date_range_str}.xlsx'
    elif '킨더' in pattern:
        output_file = f'재고리스트(누계) 킨더 통합 {date_range_str}.xlsx'
    
    # 최종 파일 저장
    grouped_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"파일 생성 완료: {output_file}")

'''
    # 원본 파일 삭제 (추가된 부분)
    for file in files:
        try:
            os.remove(file)
            print(f"원본 파일 삭제 완료: {file}")
        except Exception as e:
            print(f"파일 삭제 실패: {file}, 오류: {str(e)}")
'''    

os.system("pause")