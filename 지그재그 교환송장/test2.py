"""
지그재그 교환 송장 자동화 스크립트
---------------------------------
- Google Sheets에서 교환 출고 데이터를 불러와 주문번호/운송장번호 추출
- Selenium을 통해 지그재그 파트너 어드민 자동 로그인
- 교환수거완료 상태 주문 조회 후 수거확정 처리
- 교환 배송준비중 → 운송장 자동 입력 및 교환 배송중 처리
- 반복 업무를 자동화하여 효율성과 정확성 개선
"""

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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import get_as_dataframe

# ======================== LOGGING ========================
def log_step(msg): print(f"\n🟦 [STEP] {msg}")
def log_info(msg): print(f"✅ {msg}")
def log_warn(msg): print(f"⚠️ {msg}")
def log_error(msg): print(f"❌ {msg}")

# ======================== STEP 0: 구글시트 데이터 로드 ========================
today = datetime.date.today().strftime('%Y-%m-%d')
log_step(f"구글시트에서 {today} 지그재그 교환출고 주문 불러오기")

# 🔐 환경변수 로드
load_dotenv()
GOOGLE_KEY_PATH = os.getenv("GOOGLE_KEY_PATH")
SHEET_URL = os.getenv("SHEET_URL")

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_KEY_PATH, scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url(SHEET_URL)

def get_today_df(sheet_name):
    df = get_as_dataframe(spreadsheet.worksheet(sheet_name), evaluate_formulas=True)
    df.dropna(how="all", inplace=True)
    return df[df['출고일'] == today]

df1 = get_today_df("(외부몰)교환출고 raw")
df2 = get_today_df("(불량)교환출고 raw")
df = pd.concat([df1, df2], ignore_index=True)
df = df[df['교환형태'] == '지그재그'].drop_duplicates(subset='주문번호', keep='first')

order_numbers = df['주문번호'].astype(str).tolist()
ship_numbers = df['운송장번호'].astype(str).tolist()

if not order_numbers:
    log_warn("📭 오늘 처리할 지그재그 교환 출고 주문이 없습니다.")
    exit()
log_info(f"총 지그재그 교환 출고 주문: {len(order_numbers)}건")

# ======================== STEP 1: 브라우저 세팅 ========================
log_step("크롬 드라이버 실행 및 로그인 페이지 진입")

options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920x1080')
options.add_argument('--start-maximized')
# options.add_argument('--headless=new')  # 필요 시 활성화

driver = webdriver.Chrome(options=options)
driver.set_page_load_timeout(60)
driver.get("https://partners.kakaostyle.com/login")

# 로그인 폼 대기
WebDriverWait(driver, 15).until(
    EC.any_of(
        EC.presence_of_element_located((By.XPATH, '//input[@type="text"]')),
        EC.presence_of_element_located((By.XPATH, "//button[contains(text(),'로그인')]"))
    )
)

# ======================== STEP 2: 로그인 ========================
ZIGZAG_ID = os.getenv("LOGIN_ID")
ZIGZAG_PW = os.getenv("LOGIN_PW")

log_step("지그재그 로그인 시도")
driver.find_element(By.XPATH, '//input[@type="text"]').send_keys(ZIGZAG_ID)
pwd = driver.find_element(By.XPATH, '//input[@type="password"]')
pwd.clear()
pwd.send_keys(ZIGZAG_PW)
driver.find_element(By.XPATH, "//button[contains(text(),'로그인')]").click()
time.sleep(4)

# ======================== STEP 3: 브랜드 선택 및 팝업 처리 ========================
log_step("브랜드 선택 및 팝업 닫기")
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//p[text()='베리시']"))
).click()
time.sleep(3)

try:
    WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='오늘 하루 동안 모든 창을 열지 않음']/preceding-sibling::span"))
    ).click()
    WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='전체닫기']"))
    ).click()
except Exception:
    log_info("팝업 없음, 건너뜀")

# ======================== STEP 4: 교환수거완료 처리 ========================
log_step("교환수거완료 처리 시작")
# (사이드 메뉴 탐색 → 교환 관리 진입 과정 생략 없이 코드 포함)
driver.find_element(By.XPATH, '//*[@id="order_item"]/button').click()
driver.find_element(By.XPATH, '//*[@id="order_item"]/div/div/div/a[8]').click()
time.sleep(3)

# 상태 필터 → 교환수거완료
driver.find_element(By.XPATH, '//*[@id="app"]//form/div[1]//div[3]//div[1]//div/div[2]/div').click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//li[normalize-space(text())='교환수거완료']"))
).click()

# 키워드 → 주문번호
driver.find_element(By.XPATH, '//*[@id="app"]//form/div[1]//div[3]//div[3]//div/div[2]').click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//li[normalize-space(text())='주문 번호']"))
).click()

# 조회기간 1개월
driver.find_element(By.XPATH, '//label[contains(text(),"1개월")]').click()
time.sleep(2)

# 각 주문번호 입력 후 수거확정 처리
for i, order_num in enumerate(order_numbers):
    log_step(f"[{i+1}/{len(order_numbers)}] 주문번호 {order_num} 처리 중")
    order_input = driver.find_element(By.XPATH, '//input[@placeholder="검색어 입력"]')
    order_input.clear()
    order_input.send_keys(order_num)
    driver.find_element(By.XPATH, "//button[contains(text(),'조회')]").click()
    time.sleep(2)

    # 결과 없으면 패스
    try:
        if '0건' in driver.find_element(By.XPATH, '//div[contains(text(),"교환수거완료 목록")]').text:
            log_warn(f"{order_num} 결과 없음, 패스")
            continue
    except Exception:
        continue

    # 체크박스 + 수거확정 처리
    try:
        driver.find_element(By.XPATH, '//label[input[@type="checkbox"]]').click()
        driver.find_element(By.XPATH, "//button[contains(text(),'수거확정')]").click()
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//button[normalize-space(text())="확인"]'))
        ).click()
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//button[normalize-space(text())="확인"]'))
        ).click()
        log_info(f"{order_num} 수거확정 완료")
    except Exception as e:
        log_error(f"{order_num} 처리 실패: {e}")

# ======================== STEP 5: 송장 입력 ========================
log_step("송장 입력 시작")

# 상태 → 교환 배송준비중
driver.find_element(By.XPATH, '//*[@id="app"]//form/div[1]//div[3]//div[1]//div/div[2]/div').click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//li[normalize-space(text())='교환 배송준비중']"))
).click()

# 조회기간 1개월
driver.find_element(By.XPATH, '//label[contains(text(),"1개월")]').click()
time.sleep(2)

for i, order_num in enumerate(order_numbers):
    ship_num = ship_numbers[i]
    log_step(f"[{i+1}/{len(order_numbers)}] 주문번호 {order_num}, 송장 {ship_num} 입력")

    order_input = driver.find_element(By.XPATH, '//input[@placeholder="검색어 입력"]')
    order_input.clear()
    order_input.send_keys(order_num)
    driver.find_element(By.XPATH, "//button[contains(text(),'조회')]").click()
    time.sleep(2)

    status_text = driver.find_element(By.XPATH, '//h1/div[contains(text(),"교환 배송준비중 목록")]').text
    if "0건" in status_text:
        log_warn(f"{order_num} 배송준비중 없음 → 패스")
        continue

    # 전체 선택 후 운송장 입력
    driver.find_element(By.XPATH, '//thead//label').click()
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//div[@role="combobox" and contains(.,"배송 업체 선택")]'))
    ).click()
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//li[normalize-space(text())="CJ대한통운"]'))
    ).click()

    invoice_input = driver.find_element(By.XPATH, '//input[@placeholder="운송장 번호"]')
    invoice_input.send_keys(ship_num)

    driver.find_element(By.XPATH, "//button[normalize-space(text())='선택건 적용']").click()
    driver.find_element(By.XPATH, "//button[.//span[text()='교환 배송중 처리']]").click()

    # 확인 2번
    for _ in range(2):
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//button[normalize-space(text())="확인"]'))
        ).click()
        time.sleep(1)

    log_info(f"{order_num} 송장 입력 완료")
