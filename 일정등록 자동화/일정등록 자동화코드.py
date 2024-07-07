from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd


options = Options()
options.add_argument("--start-maximized")
options.add_experimental_option("detach",True)

driver = webdriver.Chrome(options=options)

url = "https://gw.datastreams.co.kr/"

driver.get(url)
time.sleep(1)

print("STPE 1 : 로그인 후 일정등록 화면으로 이동하기 !!")
# 1. ID입력
로그인ID = driver.find_element(By.NAME, "loginid")
로그인ID.send_keys("donghwan.kim")

# 2. PASSWD 입력
패스워드 = driver.find_element(By.NAME, "password")
패스워드.send_keys("dhkim0014*")

# 3. 로그인버튼 입력
time.sleep(1)
로그인버튼 = driver.find_element(By.CSS_SELECTOR, ".button-ds-primary")
로그인버튼.click()

# 4. 상단 제품 클릭
time.sleep(5)
상단제품 = driver.find_element(By.LINK_TEXT, '제품')
상단제품.click()

# 4. 좌측 유지보수관리 클릭
time.sleep(3)
좌측유지보수관리 = driver.find_element(By.XPATH, "//a[contains(., '유지보수관리')]")
좌측유지보수관리.click()


# 5. 좌측 유지보수관리 > 유지보수 일정관리 클릭
time.sleep(1)
유지보수_일정관리 = driver.find_element(By.LINK_TEXT, '유지보수 일정관리')
유지보수_일정관리.click()

#같은 경로에있는 csv파일을 읽어서 데이터 가져오기
data = pd.read_csv('C:/Users/admin/Desktop/김동환/Worksapce/Selenium/일정등록 자동화/일정등록 폼.csv')  

#데이터를 1개의 list에 담기
listdata = []
for i, row in data.iterrows():
    listdata.append([row['시작일자'], row['시작시'],row['시작분'],row['종료시'],row['종료분'],row['방문목적'],row['지원사이트'],row['유지보수구분'],row['문제점'],row['조치방법'],row['조치결과'],row['업체담당자']])

print("STPE 2 : 경비청구 자동 입력 시작 !!")

for i in range(data['시작일자'].count()):

    # 내부 iframe으로 전환
    time.sleep(2)
    iframe = driver.find_element(By.ID, 'frameContent')
    driver.switch_to.frame(iframe)

    # 6. 등록버튼 클릭
    time.sleep(5)
    등록버튼 = driver.find_element(By.XPATH, "//button[text()='등록']")
    등록버튼.click()

    # 모든 창 핸들 가져오기
    window_handles = driver.window_handles

    # 팝업 창으로 전환 (기본적으로 두 번째 창이 팝업 창이므로 인덱스 1 사용)
    driver.switch_to.window(window_handles[1])

    #1 시작 일자
    time.sleep(3)
    시작일자 = driver.find_element(By.NAME, "pmStartDate")
    시작일자.send_keys(Keys.CONTROL, 'a')
    시작일자.send_keys(Keys.BACKSPACE)
    시작일자.send_keys(listdata[i][0])

    #2 시작/종료 시간
    time.sleep(0.5)
    시작시 = driver.find_element(By.NAME, "pmStartHour")
    driver.execute_script("arguments[0].setAttribute('value', arguments[1])", 시작시, listdata[i][1])
    time.sleep(0.5)
    시작분 = driver.find_element(By.NAME, "pmStartMin")
    driver.execute_script("arguments[0].setAttribute('value', arguments[1])", 시작분, listdata[i][2])
    time.sleep(0.5)
    종료시 = driver.find_element(By.NAME, "pmEndHour")
    driver.execute_script("arguments[0].setAttribute('value', arguments[1])", 종료시, listdata[i][3])
    time.sleep(0.5)
    종료분 = driver.find_element(By.NAME, "pmEndMin")
    driver.execute_script("arguments[0].setAttribute('value', arguments[1])", 종료분, listdata[i][4])

    #3 방문목적
    time.sleep(1.5)
    방문목적 = driver.find_element(By.NAME, "pmTitle")
    방문목적.send_keys(listdata[i][5])

    #4 지원 사이트 팝업이생성되는 사이트일때 처리##################################
    time.sleep(1.5)
    지원사이트 = driver.find_element(By.NAME, "pmSiteNm")
    지원사이트.send_keys(listdata[i][6])

    #5 사업부
    time.sleep(1.5)
    사업부 = driver.find_element(By.NAME, "pmCmpDept")
    사업부.send_keys(listdata[i][6])

    #6 유지보수 구분
    time.sleep(1.5)
    유지보수구분 = driver.find_element(By.NAME, "pmMitncGb")
    if listdata[i][7] == "-1":
        유지보수구분.send_keys(Keys.ARROW_UP)
    elif listdata[i][7] == "0":
        pass
    else:
        for _ in range(listdata[i][7]):
            유지보수구분.send_keys(Keys.ARROW_DOWN)
    
    #7. 처리 담당자
    time.sleep(1.5)
    처리담당자 = driver.find_element(By.NAME, "pmEmpNm")
    처리담당자.send_keys(Keys.CONTROL, 'a')
    처리담당자.send_keys(Keys.BACKSPACE)
    처리담당자.send_keys("헬프데스크")
    처리담당자.send_keys(Keys.ENTER)
    time.sleep(1.5)

    window_handles = driver.window_handles
    driver.switch_to.window(window_handles[2])
    
    time.sleep(1.5)
    헬프데스크 = driver.find_element(By.LINK_TEXT, '헬프데스크')
    헬프데스크.click()
    time.sleep(1.5)
    driver.switch_to.window(window_handles[1])

    #8. 팀구분
    time.sleep(1.5)
    팀구분 = driver.find_element(By.NAME, "pmTssubteam")
    팀구분.send_keys(Keys.ARROW_DOWN)
    팀구분.send_keys(Keys.ARROW_DOWN)

    #버그로 뜨는 헬프데스크 선택창 닫기처리
    time.sleep(1.5)
    window_handles = driver.window_handles
    driver.switch_to.window(window_handles[2])
    time.sleep(1.5)
    헬프데스크 = driver.find_element(By.LINK_TEXT, '헬프데스크')
    헬프데스크.click()
    time.sleep(1.5)
    driver.switch_to.window(window_handles[1])

    #9. 문제점
    time.sleep(1.5)
    문제점 = driver.find_element(By.NAME, "pmIssueCntn")
    문제점.send_keys(listdata[i][8])

    #10. 조치방법
    time.sleep(1.5)
    조치방법 = driver.find_element(By.NAME, "pmActionCntn")
    조치방법.send_keys(listdata[i][9])

    #11. 조치결과
    time.sleep(1.5)
    조치결과 = driver.find_element(By.NAME, "pmResultCntn")
    조치결과.send_keys(listdata[i][10])

    #12 업체담당자
    time.sleep(1.5)
    업체담당자 = driver.find_element(By.NAME, "pmCusCharge")
    업체담당자.send_keys(listdata[i][11]) 

    #13 저장버튼 클릭
    time.sleep(1.5)
    저장버튼 = driver.find_element(By.XPATH, "//button[text()='저장']")
    저장버튼.click()

    #14 저장 성공 알럿 누르고 처음
    time.sleep(2)
    alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
    alert.accept() 
    time.sleep(2)
    alert.accept() 

    #15 부모 창으로 다시 전환
    time.sleep(5)
    driver.switch_to.window(window_handles[0])


    
































# # 7. 프로젝트명 입력
# time.sleep(3)
# 등록팝업_프로젝트명 = driver.find_element(By.NAME, "pmPJTNM")
# 등록팝업_프로젝트명.send_keys("[사내]PVS부문_기술지원팀")
# 등록팝업_프로젝트명.send_keys(Keys.ENTER)


# #.8. 등록일자 입력
# time.sleep(3)
# 등록팝업_등록일자 = driver.find_element(By.NAME, "pmUSEDATE")
# 등록팝업_등록일자.send_keys(listdata[0][0])


# # 9. 실제 사용일자 입력
# time.sleep(3)
# 등록팝업_실제사용일자 = driver.find_element(By.NAME, "pmRealUSEDATE")
# 등록팝업_실제사용일자.send_keys(listdata[0][0])


# # 10. 결제유형 입력
# time.sleep(3)
# 등록팝업_결제유형 = driver.find_element(By.NAME, "selCASHCD")

# # 결제유형 : 개인카드는 아래방향키 2번
# for _ in range(2): 
#     등록팝업_결제유형.send_keys(Keys.ARROW_DOWN)



# #11. 계정항목 입력
# time.sleep(3)
# 등록팝업_계정항목 = driver.find_element(By.NAME, "selACCTCD")

# for _ in range(listdata[0][1]): 
#     등록팝업_계정항목.send_keys(Keys.ARROW_DOWN)

# #12. 사용금액 입력
# time.sleep(3)
# 등록팝업_사용금액 = driver.find_element(By.NAME, "pmAMT")
# 등록팝업_사용금액.send_keys(listdata[0][2])


# #13. 사용내역 입력
# time.sleep(3)
# 등록팝업_사용내역 = driver.find_element(By.NAME, "pmCNTN")
# 등록팝업_사용내역.send_keys(listdata[0][3])


# #14. 저장 입력
# time.sleep(5)
# 등록팝업_저장버튼 = driver.find_element(By.XPATH, "//button[text()='저장']")
# 등록팝업_저장버튼.click()

# # 15. 저장성공 알럿창 확인 
# time.sleep(5)
# alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
# alert.accept()

# print("STPE 3 : 경비청구 추가 저장 반복 시작 !!")

# for i in range(data['사용금액'].count()):

#     # 16. 처음 1회는 등록을 했으니 처음꺼는 넘기기
#     if i == 0:
#         continue;
    
#     #17. 등록일자 입력
#     time.sleep(3)
#     등록팝업_등록일자 = driver.find_element(By.NAME, "pmUSEDATE")
#     등록팝업_등록일자.send_keys(Keys.CONTROL, 'a')
#     등록팝업_등록일자.send_keys(Keys.BACKSPACE)
#     등록팝업_등록일자.send_keys(listdata[i][0])
    
#     #18. 실제사용일자 입력
#     time.sleep(3)
#     등록팝업_실제사용일자 = driver.find_element(By.NAME, "pmRealUSEDATE")
#     등록팝업_실제사용일자.send_keys(Keys.CONTROL, 'a')
#     등록팝업_실제사용일자.send_keys(Keys.BACKSPACE)
#     등록팝업_실제사용일자.send_keys(listdata[i][0])   
    
#     #19. 계정항목 입력
#     time.sleep(3)
#     등록팝업_계정항목 = driver.find_element(By.NAME, "selACCTCD")

#     #위의 키 17번 누르기
#     for _ in range(17):
#         등록팝업_계정항목.send_keys(Keys.ARROW_UP)

#     for _ in range(listdata[i][1]):
#         등록팝업_계정항목.send_keys(Keys.ARROW_DOWN)

#     #20. 사용금액 입력
#     time.sleep(3)
#     등록팝업_사용금액 = driver.find_element(By.NAME, "pmAMT")
#     등록팝업_사용금액.send_keys(Keys.CONTROL, 'a')
#     등록팝업_사용금액.send_keys(Keys.BACKSPACE)
#     등록팝업_사용금액.send_keys(listdata[i][2])
    
#     #21. 사용내역 입력
#     time.sleep(3)
#     등록팝업_사용내역 = driver.find_element(By.NAME, "pmCNTN")
#     등록팝업_사용내역.send_keys(Keys.CONTROL, 'a')
#     등록팝업_사용내역.send_keys(Keys.BACKSPACE)
#     등록팝업_사용내역.send_keys(listdata[i][3])
    
#     #22. 추가저장 버튼
#     time.sleep(5)
#     등록팝업_저장버튼 = driver.find_element(By.XPATH, "//button[text()='추가저장']")
#     등록팝업_저장버튼.click()
    
#     #23. 저장성공 알럿창 확인 
#     time.sleep(5)
#     alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
#     alert.accept()

# print("STPE 4 : 경비청구 자동 저장 종료 !! 데이터 잘 들어갔나 직접 확인 할것 !!")
