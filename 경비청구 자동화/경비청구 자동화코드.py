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

print("STPE 1 : 로그인 후 경비청구 화면으로 이동하기 !!")
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

# 4. 상단 프로젝트 클릭
time.sleep(5)
상단프로젝트 = driver.find_element(By.LINK_TEXT, '프로젝트')
상단프로젝트.click()

# 4. 좌측 프로젝트비용 클릭
time.sleep(3)
좌측프로젝트비용 = driver.find_element(By.XPATH, "//a[contains(., '프로젝트비용')]")
좌측프로젝트비용.click()


# 5. 좌측 프로젝트비용 > 비용관리 클릭
time.sleep(1)
비용관리 = driver.find_element(By.LINK_TEXT, '비용관리')
비용관리.click()

# 내부 iframe으로 전환
time.sleep(2)
iframe = driver.find_element(By.ID, 'frameContent')
driver.switch_to.frame(iframe)

# # iframe 내부의 HTML 코드 가져오기 (html 확인용)
# html = driver.page_source
# with open('copy.html', 'w', encoding='utf-8') as file:
#     file.write(html)


# 6. 등록버튼 클릭
time.sleep(5)
등록버튼 = driver.find_element(By.XPATH, "//button[text()='등록']")
등록버튼.click()

# 모든 창 핸들 가져오기
window_handles = driver.window_handles

# 팝업 창으로 전환 (기본적으로 두 번째 창이 팝업 창이므로 인덱스 1 사용)
driver.switch_to.window(window_handles[1])

# 엑셀 파일 경로
file_path = r'C:\Users\admin\Desktop\김동환\Worksapce\Selenium\경비청구 자동화\경비청구 폼.xlsx'

# 엑셀 파일 읽어오기
data = pd.read_excel(file_path)

#데이터를 1개의 list에 담기
listdata = []
for i, row in data.iterrows():
    listdata.append([row['날짜(YYYY-MM-DD)'], row['계정항목'],row['사용금액'],row['사용내역']])

print("STPE 2 : 경비청구 처음 1회 입력 시작 !!")
# 7. 프로젝트명 입력
time.sleep(3)
등록팝업_프로젝트명 = driver.find_element(By.NAME, "pmPJTNM")
등록팝업_프로젝트명.send_keys("[사내]PVS부문_기술지원팀")
등록팝업_프로젝트명.send_keys(Keys.ENTER)


#.8. 등록일자 입력
time.sleep(3)
등록팝업_등록일자 = driver.find_element(By.NAME, "pmUSEDATE")
등록팝업_등록일자.send_keys(listdata[0][0])


# 9. 실제 사용일자 입력
time.sleep(3)
등록팝업_실제사용일자 = driver.find_element(By.NAME, "pmRealUSEDATE")
등록팝업_실제사용일자.send_keys(listdata[0][0])


# 10. 결제유형 입력
time.sleep(3)
등록팝업_결제유형 = driver.find_element(By.NAME, "selCASHCD")

# 결제유형 : 개인카드는 아래방향키 2번
for _ in range(2): 
    등록팝업_결제유형.send_keys(Keys.ARROW_DOWN)



#11. 계정항목 입력
time.sleep(3)
등록팝업_계정항목 = driver.find_element(By.NAME, "selACCTCD")

for _ in range(listdata[0][1]): 
    등록팝업_계정항목.send_keys(Keys.ARROW_DOWN)

#12. 사용금액 입력
time.sleep(3)
등록팝업_사용금액 = driver.find_element(By.NAME, "pmAMT")
등록팝업_사용금액.send_keys(listdata[0][2])


#13. 사용내역 입력
time.sleep(3)
등록팝업_사용내역 = driver.find_element(By.NAME, "pmCNTN")
등록팝업_사용내역.send_keys(listdata[0][3])


#14. 저장 입력
time.sleep(5)
등록팝업_저장버튼 = driver.find_element(By.XPATH, "//button[text()='저장']")
등록팝업_저장버튼.click()

# 15. 저장성공 알럿창 확인 
time.sleep(5)
alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
alert.accept()

print("STPE 3 : 경비청구 추가 저장 반복 시작 !!")

for i in range(data['사용금액'].count()):

    # 16. 처음 1회는 등록을 했으니 처음꺼는 넘기기
    if i == 0:
        continue;
    
    #17. 등록일자 입력
    time.sleep(3)
    등록팝업_등록일자 = driver.find_element(By.NAME, "pmUSEDATE")
    등록팝업_등록일자.send_keys(Keys.CONTROL, 'a')
    등록팝업_등록일자.send_keys(Keys.BACKSPACE)
    등록팝업_등록일자.send_keys(listdata[i][0])
    
    #18. 실제사용일자 입력
    time.sleep(3)
    등록팝업_실제사용일자 = driver.find_element(By.NAME, "pmRealUSEDATE")
    등록팝업_실제사용일자.send_keys(Keys.CONTROL, 'a')
    등록팝업_실제사용일자.send_keys(Keys.BACKSPACE)
    등록팝업_실제사용일자.send_keys(listdata[i][0])   
    
    #19. 계정항목 입력
    time.sleep(3)
    등록팝업_계정항목 = driver.find_element(By.NAME, "selACCTCD")

    #위의 키 17번 누르기
    for _ in range(17):
        등록팝업_계정항목.send_keys(Keys.ARROW_UP)

    for _ in range(listdata[i][1]):
        등록팝업_계정항목.send_keys(Keys.ARROW_DOWN)

    #20. 사용금액 입력
    time.sleep(3)
    등록팝업_사용금액 = driver.find_element(By.NAME, "pmAMT")
    등록팝업_사용금액.send_keys(Keys.CONTROL, 'a')
    등록팝업_사용금액.send_keys(Keys.BACKSPACE)
    등록팝업_사용금액.send_keys(listdata[i][2])
    
    #21. 사용내역 입력
    time.sleep(3)
    등록팝업_사용내역 = driver.find_element(By.NAME, "pmCNTN")
    등록팝업_사용내역.send_keys(Keys.CONTROL, 'a')
    등록팝업_사용내역.send_keys(Keys.BACKSPACE)
    등록팝업_사용내역.send_keys(listdata[i][3])
    
    #22. 추가저장 버튼
    time.sleep(5)
    등록팝업_저장버튼 = driver.find_element(By.XPATH, "//button[text()='추가저장']")
    등록팝업_저장버튼.click()
    
    #23. 저장성공 알럿창 확인 
    time.sleep(5)
    alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
    alert.accept()

print("STPE 4 : 경비청구 자동 저장 종료 !! 데이터 잘 들어갔나 직접 확인 할것 !!")
