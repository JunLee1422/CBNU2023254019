import pandas as pd



import pandas as pd
import numpy as np
import os


# 현재 스크립트 파일의 디렉토리 경로를 가져오기
script_dir = os.path.dirname(os.path.abspath(__file__))

# 파일 경로를 조합하여 MRP_final.xlsx 파일의 전체 경로를 얻기
file_path = os.path.join(script_dir, 'MRP_final.xlsx')

# 파일을 읽어오기
data = pd.read_excel(file_path)

# 데이터 정의
mps_data = {
    '품목코드': ['A', 'B', 'D', 'A', 'B', 'D', 'A', 'B', 'D'],
    '품목명': ['계량기 A', '계량기 B', '부분 조립품 D', '계량기 A', '계량기 B', '부분 조립품 D', '계량기 A', '계량기 B', '부분 조립품 D'],
    '수량': [1250, 470, 270, 850, 360, 250, 550, 560, 320],
    '납기': [9, 9, 9, 13, 13, 13, 17, 17, 17]
}

bom_data = {
    'Parent': ['A', 'A', 'B', 'C'],
    'Child': ['C', 'D', 'C', 'D'],
    'Qty': [1, 1, 1, 2]
}

irf_data = {
    '품목코드': ['A', 'B', 'C', 'D'],
    '현재재고': [50, 60, 40, 200],
    '인도기간': [2, 2, 1, 1],
    '안전재고': [0, 0, 5, 20],
    '예정입고량': [0, 10, 0, 100],
    '예정입고일': [0, 5, 0, 4],
    '주문량': [1, 1, 2000, 5000]
}

# 데이터프레임 생성
mps_df = pd.DataFrame(mps_data)
bom_df = pd.DataFrame(bom_data)
irf_df = pd.DataFrame(irf_data)

# 결과 테이블 초기화
result_data = []

# 각 주차에 대해 MRP 계산
for week in range(4, 18):
    for item in ['A', 'B', 'C', 'D']:
        item_mps = mps_df[(mps_df['품목코드'] == item) & (mps_df['납기'] == week)]
        item_irf = irf_df[irf_df['품목코드'] == item]

        if not item_mps.empty:
            demand = item_mps['수량'].sum()
        else:
            demand = 0

        scheduled_receipts = item_irf['예정입고량'].values[0]
        projected_inventory = item_irf['현재재고'].values[0] + scheduled_receipts
        net_requirements = demand - projected_inventory

        if net_requirements > 0:
            planned_orders = net_requirements
        else:
            planned_orders = 0

        planned_receipts = max(net_requirements, 0)

        # 결과 테이블 업데이트
        result_data.append({'품목코드': item, '주차': week, '구분': '총소요량', '값': demand})
        result_data.append({'품목코드': item, '주차': week, '구분': '예정입고', '값': scheduled_receipts})
        result_data.append({'품목코드': item, '주차': week, '구분': '예상재고', '값': projected_inventory})
        result_data.append({'품목코드': item, '주차': week, '구분': '순소요량', '값': net_requirements})
        result_data.append({'품목코드': item, '주차': week, '구분': '계획수주', '값': planned_orders})
        result_data.append({'품목코드': item, '주차': week, '구분': '계획발주', '값': planned_receipts})

# 결과 데이터프레임 생성
result_df = pd.DataFrame(result_data)

# 테이블 출력
result_pivot = result_df.pivot(index=['품목코드', '구분'], columns='주차', values='값')
result_pivot.reset_index(inplace=True)

# 결과 데이터프레임을 엑셀 파일로 저장
result_pivot.to_excel('MRP_output2.xlsx', index=False)

# 결과 데이터프레임을 바로 출력
result_pivot