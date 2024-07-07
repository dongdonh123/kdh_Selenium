import pandas as pd

data = pd.read_csv('경비청구 폼.csv')  #csv파일을 읽겠다
# print(data) #출력하기

# print(data['날짜(YYYY-MM-DD)']) #첫번째열만 출력
print(data['날짜(YYYY-MM-DD)'].count()) #첫번째열 개수 출력 (총 반복횟수 확인)

##리스트를 각각 만들기
###############################################

# 번호 = data['번호'].values

# 날짜 = data['날짜(YYYY-MM-DD)'].values

# 계정항목 = data['계정항목'].values

# 사용금액 = data['사용금액'].values

# 사용내역 = data['사용내역'].values

###########################################

#한꺼번에 담기
listdata = []
for i, row in data.iterrows():
    listdata.append([row['날짜(YYYY-MM-DD)'], row['계정항목'],row['사용금액'],row['사용내역']])

# print(listdata)

# print("0번째" +listdata[0][0])


for i in range(8): 
    print(listdata[i])
    for j in range(4):
        print(listdata[i][j])