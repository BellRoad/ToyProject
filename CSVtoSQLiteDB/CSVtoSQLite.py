import sqlite3
import csv
import os

def create_db_from_csv(csv_filename):
    # CSV 파일 이름에서 확장자 제거
    base_name = os.path.splitext(csv_filename)[0]
    
    # DB 파일 이름 생성
    db_filename = f"{base_name}.db"
    
    # SQLite 연결
    conn = sqlite3.connect(db_filename)
    cursor = conn.cursor()
    
    try:
        # CSV 파일 열기
        with open(csv_filename, 'r', newline='', encoding='utf-8') as csvfile:
            csv_reader = csv.reader(csvfile)
            headers = next(csv_reader)  # 첫 번째 행을 헤더로 사용
            
            # 테이블 생성 (테이블 이름과 컬럼 이름에 따옴표 추가)
            table_name = f'"{base_name}"'
            columns = ','.join([f'"{header}" TEXT' for header in headers])
            cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} ({columns})")
            
            # 데이터 삽입 (매개변수화된 쿼리 사용)
            placeholders = ','.join(['?' for _ in headers])
            insert_query = f"INSERT INTO {table_name} VALUES ({placeholders})"
            cursor.executemany(insert_query, csv_reader)
        
        conn.commit()
        print(f"'{csv_filename}'의 내용이 '{db_filename}' 데이터베이스의 '{base_name}' 테이블에 성공적으로 저장되었습니다.")
    
    except Exception as e:
        print(f"오류 발생: {e}")
    
    finally:
        conn.close()

# CSV 파일 이름 입력 받기
csv_filename = input("처리할 CSV 파일의 이름을 입력하세요(확장자포함) : ")

# 프로그램 실행
if os.path.exists(csv_filename):
    create_db_from_csv(csv_filename)
else:
    print(f"'{csv_filename}' 파일을 찾을 수 없습니다. 파일이 현재 폴더에 있는지 확인해주세요.")
