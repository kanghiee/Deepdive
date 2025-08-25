 # ======================== IMPORTS ========================
import os
import time
import datetime
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import get_as_dataframe
from webdriver_manager.chrome import ChromeDriverManager
import gspread


# ======================== LOGGING ========================
def log_step(msg):
    print(f"\n🟦 [STEP] {msg}")

def log_info(msg):
    print(f"✅ {msg}")

def log_warn(msg):
    print(f"⚠️ {msg}")

def log_error(msg):
    print(f"❌ {msg}")

# ======================== STEP 0: 구글시트 먼저 확인 ========================
today = datetime.date.today().strftime('%Y-%m-%d')
log_step(f"구글시트에서 {today} 지그재그 교환출고 주문 불러오기")
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    "/Users/deepdive/Documents/강희/지그재그 교환송장/new_google_API_KEY/mystical-button-438607-h5-ecc3ead3147d.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1PhdbPo4WEcPGJBKLwqDYAny_WpH25cJNKrZzM42bSyU/edit#gid=600415326")

def get_today_df(sheet_name):
    df = get_as_dataframe(spreadsheet.worksheet(sheet_name), evaluate_formulas=True)
    df.dropna(how="all", inplace=True)
    return df[df['출고일'] == today]

df1 = get_today_df("(외부몰)교환출고 raw")
df2 = get_today_df("(불량)교환출고 raw")
df = pd.concat([df1, df2], ignore_index=True)
df = df[df['교환형태'] == '지그재그'].drop_duplicates(subset='주문번호', keep='first')
order_number = df['주문번호'].astype(str).tolist()
ship_number = df['운송장번호'].astype(str).tolist()

if not order_number:
    log_warn("📭 오늘 처리할 지그재그 교환 출고 주문이 없습니다. 스크립트를 종료합니다.")
    exit()

log_info(f"총 지그재그 교환 출고 주문: {len(order_number)}건")


# ======================== BROWSER SETUP ========================
from pathlib import Path
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

log_step("크롬 드라이버 실행 및 지그재그 로그인 페이지 진입")

options = Options()
# options.add_argument('--headless=new')  # 필요 시
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920x1080')
options.add_argument('--start-maximized')

driver = webdriver.Chrome(options=options)

driver.set_page_load_timeout(60)
driver.get("https://partners.kakaostyle.com/login")

# 로그인 폼/버튼 보일 때까지 잠깐만 대기 (페이지 늦게 열릴 때 대비)
WebDriverWait(driver, 15).until(
    EC.any_of(
        EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div[2]/main/form/div/div[2]/div[1]/input')),
        EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'카카오') and contains(text(),'로그인')]"))
    )
)


# ======================== 로그인 ========================
load_dotenv()

ZIGZAG_ID = os.getenv("LOGIN_ID")
ZIGZAG_PW = os.getenv("LOGIN_PW")
log_step("지그재그 로그인 시도")
driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/main/form/div/div[2]/div[1]/input').send_keys(ZIGZAG_ID)
time.sleep(1)
pwd = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/main/form/div/div[2]/div[2]/input')
pwd.clear()
pwd.send_keys(ZIGZAG_PW)
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/main/form/div/div[3]/button').click()
time.sleep(4)

# ======================== 브랜드 선택 및 팝업 제거 ========================

log_step("베리시 브랜드 선택 및 팝업 닫기")

# 브랜드 선택
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//p[text()='베리시']"))
).click()
time.sleep(3)

# 팝업 닫기 (존재할 경우만)
try:
    WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='오늘 하루 동안 모든 창을 열지 않음']/preceding-sibling::span[@data-role='check-mark']"))
    ).click()

    WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='전체닫기' and contains(@class, 'AllCloseButton')]"))
    ).click()

    time.sleep(1)
except Exception as e:
    log_step("팝업이 나타나지 않아 건너뜀")

# ======================== 수거확정 처리 ========================
log_step("수거확정 처리 시작")
# 왼쪽 사이드 메뉴 / 주문 배송 관리 클릭
log_step("왼쪽 사이드 메뉴/ 주문 배송 관리 클릭")
lists = driver.find_element(By.XPATH,'//*[@id="order_item"]/button/span[2]')
lists.click()


time.sleep(0.5)

# 왼쪽 사이드 메뉴 / 주문 배송관리 -> 교환 관리 클릭
log_step('왼쪽 사이드 메뉴 / 주문 배송관리 -> 교환 관리 클릭')
exchange_list = driver.find_element(By.XPATH,'//*[@id="order_item"]/div/div/div/a[8]')
exchange_list.click()

time.sleep(3)


log_step("교환수거완료 클릭")
# 처리상태 클릭
states = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[3]/div[1]/div/div[2]/div')
states.click()
time.sleep(1)


# "교환수거완료" 항목을 클릭
exchange_done = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//li[@role='option' and normalize-space(text())='교환수거완료']"
    ))
)
exchange_done.click()

time.sleep(1)


log_step("검색옵션 주문번호로 변경")
# 키워드 클릭
keyword = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[3]/div[3]/div/div[2]/div/div/div[1]')
keyword.click()

time.sleep(1)



# "주문 번호" 항목을 클릭
order_number_option = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//li[@role='option' and normalize-space(text())='주문 번호']"
    ))
)
order_number_option.click()

time.sleep(1)


log_step("조회기간 1개월 변경")
# 조회기간 1개월 클릭
dateline = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[3]/div[4]/div/div[2]/div[1]/div[2]/div[1]/label[3]')
dateline.click()

time.sleep(3)

log_step('주문번호 입력 및 수거확정 시작')
for i, order_number_value in enumerate(order_number):
    print(f"{i+1}번째 주문번호 입력 중: {order_number_value}")

    # 1. input 요소 찾기
    order_input = driver.find_element(
        By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[3]/div[3]/div/div[2]/div/div/div[2]/input'
    )             
    # 2. input 클릭 후 커서를 끝으로 이동
    order_input.click()
    time.sleep(0.5)
    order_input.send_keys(Keys.END)
    
    # 3. BACKSPACE 여러 번 입력해서 React 내부 state까지 삭제
    for _ in range(50):
        order_input.send_keys(Keys.BACKSPACE)
    time.sleep(0.2)

    # 4. input 이벤트 강제 트리거 (React state 동기화)
    driver.execute_script("""
        const input = arguments[0];
        const event = new Event('input', { bubbles: true });
        input.dispatchEvent(event);
    """, order_input)

    # 5. 디버깅용 출력
    print("초기화 후 value:", order_input.get_attribute("value"))

    # 6. 새로운 주문번호 입력
    order_input.send_keys(order_number_value)

    # 7. 조회 버튼 클릭
    search_btn = driver.find_element(
        By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[4]/button[1]'
    )
    search_btn.click()
    time.sleep(2)  # 결과 로딩 대기


    # 🧠 "0건" 텍스트 확인 → 있으면 continue
    try:
        result_summary = driver.find_element(
            By.XPATH, '//div[contains(text(), "교환수거완료 목록")]'
        ).text

        if '0건' in result_summary:
            print(f"📭 교환수거완료 목록 0건 — {order_number_value}, 다음으로 넘어감")
            continue

    except Exception as e:
        print(f"⚠️ 교환수거완료 텍스트 확인 실패 — {order_number_value}, 다음으로 넘어감")
        continue


    

    time.sleep(1.5)  # 조회 후 대기

    # 체크박스 또는 체크 마크 안전하게 클릭
    try:
        checkbox = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH,
                '//label[input[@type="checkbox"] and span[@data-role="check-mark"]]'
            ))
        )
        checkbox.click()
        print("✅ 주문건 체크박스 클릭 성공")
    except Exception as e:
        print(f"❌ 주문건 체크박스 클릭 실패: {e}")

    time.sleep(1)
    
    log_step("수거확정 클릭")
    sugu_perfect = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div/form/div[2]/div/div[2]/div/button[1]')
    sugu_perfect.click()
    time.sleep(1.5)  # 조회 후 대기
    # 텍스트만으로 찾기 (단, 다른 "확인" 버튼도 있으면 위험함)
    confirm_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[normalize-space(text())="확인"]'))
    )
    confirm_btn.click()
    time.sleep(1.5)  # 조회 후 대기
    # 텍스트만으로 찾기 (단, 다른 "확인" 버튼도 있으면 위험함)
    confirm2_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[normalize-space(text())="확인"]'))
    )
    confirm2_btn.click()
    time.sleep(2)  


########################[송장입력]#########################
log_step('송장입력하기 단계 시작')

log_step('배송 준비중 변경')
# 처리 상태 배송 준비중으로 변경

dropdown_xpath = ('//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[3]/div[1]/div/div[2]/div')

# '클릭 가능한' 상태의 드롭다운을 찾아서 바로 변수에 저장
dropdown = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
)

# 위에서 받은 '신선한' 요소를 바로 클릭
dropdown.click()



time.sleep(1.5)

ship_done = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//li[@role='option' and normalize-space(text())='교환 배송준비중']"
    ))
)
ship_done.click()
time.sleep(2)


log_step("조회기간 1개월 변경")
# 조회기간 1개월 변경
dateline = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[3]/div[4]/div/div[2]/div[1]/div[2]/div[1]/label[3]')
dateline.click()

time.sleep(2)

#교환 송장 입력
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

for i, order_number_value in enumerate(order_number):
    print(f"{i+1}번째 주문번호 입력 중: {order_number_value}")
    
    progress = f"[{i+1}/{len(order_number)}]"
    ship_number_value = ship_number[i]

    # 주문번호 입력창 초기화
    order_input = driver.find_element(
        By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[3]/div[3]/div/div[2]/div/div/div[2]/input'
    )
    order_input.click()
    order_input.send_keys(Keys.END)
    for _ in range(50):
        order_input.send_keys(Keys.BACKSPACE)
    time.sleep(1.5)

    # React state input 동기화
    driver.execute_script("""
        const input = arguments[0];
        const event = new Event('input', { bubbles: true });
        input.dispatchEvent(event);
    """, order_input)
    
    print(f"{progress} 초기화 후 value:", order_input.get_attribute("value"))

    # 주문번호 입력 및 조회
    order_input.send_keys(order_number_value)
    search_btn = driver.find_element(
        By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div/form/div[1]/div[4]/button[1]'
    )
    search_btn.click()
    time.sleep(2.5)

    # 정확한 XPath로 텍스트 가져와서 0건이면 넘어가기
    status_text = driver.find_element(
        By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div/form/div[2]/div/div[1]/div[1]/div/h1/div'
    ).text

    if status_text.strip() == "교환 배송준비중 목록 (0건)":
        print(f"❌ {progress} 조회 결과 없음 — {order_number_value}, 다음으로 넘어감")
        continue

    # 전체 선택
    all_select = driver.find_element(By.XPATH,
        '//*[@id="app"]/div/div[2]/div[2]/div/form/div[2]/div/div[3]/div/table/thead/tr/th[1]/div/label/span'
    )
    all_select.click()
    time.sleep(1.5)

    # 배송 업체 선택
    # 1) '일괄입력' 영역 찾기
    bulk_panel = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//span[normalize-space(.)='일괄입력']/.."))
    )

    # 2) 그 안에서 role=combobox + '배송 업체 선택' 찾기
    dropdown_xpath = ".//div[@role='combobox' and contains(normalize-space(.), '배송 업체 선택')]"
    dropdown = WebDriverWait(bulk_panel, 10).until(
        EC.presence_of_element_located((By.XPATH, dropdown_xpath))
    )

    # 3) 클릭 시도
    
    WebDriverWait(bulk_panel, 5).until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
    dropdown.click()


    # cj 대한통운 선택
    cj_option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((
            By.XPATH,
            '//li[@role="option" and normalize-space(text())="CJ대한통운"]'
        ))
    )
    cj_option.click()
    time.sleep(1.5)

    # 운송장 번호 입력
    invoice_input = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((
            By.XPATH,
            '//div[contains(@class, "css-1k3hx0v")]//input[@placeholder="운송장 번호"]'
        ))
    )
    invoice_input.click()
    invoice_input.send_keys(ship_number_value)
    time.sleep(2)

    # 선택건 적용 버튼 클릭
    apply_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((
            By.XPATH,
            '//div[div/input[@placeholder="운송장 번호"]]//button[normalize-space(text())="선택건 적용"]'
        ))
    )
    apply_btn.click()
    time.sleep(1)

    # 교환 배송중 처리 버튼 클릭
    exchange_ship_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((
            By.XPATH,
            '//button[.//span[normalize-space(text())="교환 배송중 처리"]]'
        ))
    )
    exchange_ship_btn.click()
    time.sleep(1)

    # 확인 버튼 클릭 2번
    for _ in range(2):
        confirm_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//button[normalize-space(text())="확인"]'))
        )
        confirm_btn.click()
        time.sleep(1.4)

    print(f"✅ {progress} 송장 입력 완료: {ship_number_value}")