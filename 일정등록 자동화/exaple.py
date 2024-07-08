import pandas as pd
import pprint

#같은 경로에있는 csv파일을 읽어서 데이터 가져오기
data = pd.read_excel('C:/Users/admin/Desktop/김동환/Worksapce/Selenium/일정등록 자동화/일정등록 폼.xlsx')  


#데이터를 1개의 list에 담기
listdata = []
for i, row in data.iterrows():
    listdata.append([
        row['시작일자'] if pd.notna(row['시작일자']) else '',
        f"{row['시작시']:02}" if pd.notna(row['시작시']) else '',  # 2자리로 포맷팅
        f"{row['시작분']:02}" if pd.notna(row['시작분']) else '',  # 2자리로 포맷팅
        f"{row['종료시']:02}" if pd.notna(row['종료시']) else '',  # 2자리로 포맷팅
        f"{row['종료분']:02}" if pd.notna(row['종료분']) else '',  # 2자리로 포맷팅
        row['방문목적'] if pd.notna(row['방문목적']) else '',
        row['지원사이트'] if pd.notna(row['지원사이트']) else '',
        row['유지보수구분'] if pd.notna(row['유지보수구분']) else '',
        row['문제점'] if pd.notna(row['문제점']) else '',
        row['조치방법'] if pd.notna(row['조치방법']) else '',
        row['조치결과'] if pd.notna(row['조치결과']) else '',
        row['업체담당자'] if pd.notna(row['업체담당자']) else ''
    ])

pprint.pprint(listdata)
exit()