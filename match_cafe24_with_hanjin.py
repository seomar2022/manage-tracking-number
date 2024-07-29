import pandas as pd

def match_cafe24_with_hanjin(cafe24_file, hanjin_file, output_file):
    # 카페24 주문 데이터 불러오기
    cafe24_data = pd.read_csv(cafe24_file, encoding='utf-8')
    print("카페24 데이터 불러오기 완료.")
    
    # 한진택배 배송 데이터 불러오기
    hanjin_data = pd.read_excel(hanjin_file, engine='openpyxl')
    print("한진택배 데이터 불러오기 완료.")
    
    # 데이터 매칭 (예: 주문번호를 기준으로)
    matched_data = pd.merge(cafe24_data, hanjin_data[['주문번호', '운송장번호']], how='left', on='주문번호')
    print("데이터 매칭 완료.")
    
    # D1 셀에 "수량" 추가 
    matched_data.insert(3, '수량', '') 

    # 결과 저장
    matched_data.to_csv(output_file, index=False, encoding='utf-8-sig')
    print("결과 저장 완료")

# 파일 경로 설정
cafe24_file = 'excel_sample_old.csv'
hanjin_file = 'hanjin.xlsx'
output_file = 'excel_sample_old.csv'

# 매칭 실행
match_cafe24_with_hanjin(cafe24_file, hanjin_file, output_file)
