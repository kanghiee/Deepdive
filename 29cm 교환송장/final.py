# ======================== 1. IMPORTS ========================
import os
import time
import re
import imaplib
import email
import datetime
import pandas as pd
import traceback

from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import get_as_dataframe
from pathlib import Path
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ======================== 2. LOGGING FUNCTIONS ========================
def log_step(step_desc):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\n🟦 [{now}] [STEP] {step_desc}")

def log_info(message):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"✅ [{now}] {message}")

def log_warn(message):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"⚠️ [{now}] {message}")

def log_error(message):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"❌ [{now}] ERROR: {message}")

# ======================== 3. FUNCTIONS ========================
def get_latest_verification_code(email_user, app_password):
    imap = imaplib.IMAP4_SSL('imap.gmail.com')  # 수정된 부분
    imap.login(email_user, app_password)
    imap.select("inbox")

    result, data = imap.search(None, "ALL")
    mail_ids = data[0].split()
    if not mail_ids:
        return None

    latest_email_id = mail_ids[-1]
    result, data = imap.fetch(latest_email_id, "(RFC822)")
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    code_pattern = re.compile(r'\b\d{6}\b')
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode()
                match = code_pattern.search(body)
                if match:
                    return match.group()
    else:
        body = msg.get_payload(decode=True).decode()
        match = code_pattern.search(body)
        if match:
            return match.group()
    return None

def get_today_df(spreadsheet, sheet_name):
    worksheet = spreadsheet.worksheet(sheet_name)
    df = get_as_dataframe(worksheet, evaluate_formulas=True)
    df.dropna(how="all", inplace=True)
    today = datetime.date.today().strftime('%Y-%m-%d')
    return df[df['출고일'] == today]

# ======================== 4. 환경 설정 ========================

log_step("환경 변수 및 웹드라이버 설정")

load_dotenv()
email_user = os.getenv("EMAIL_USER")
app_password = os.getenv("APP_PASSWORD")

# ✅ 시트 먼저 불러오기
log_step("✅ 구글 시트 데이터 선 확인")

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    "/Users/kang_hiee/Documents/강희/Deepdive/지그재그 교환송장/new_google_API_KEY/mystical-button-438607-h5-ecc3ead3147d.json", scope)
client = gspread.authorize(creds)

sheet_url = "https://docs.google.com/spreadsheets/d/1PhdbPo4WEcPGJBKLwqDYAny_WpH25cJNKrZzM42bSyU/edit#gid=600415326"
spreadsheet = client.open_by_url(sheet_url)

df1 = get_today_df(spreadsheet, "(외부몰)교환출고 raw")
df2 = get_today_df(spreadsheet, "(불량)교환출고 raw")
df = pd.concat([df1, df2], ignore_index=True)
df = df[df["교환형태"] == "29CM"].drop_duplicates(subset="주문번호", keep="first")

order_number = df["주문번호"].astype(str).tolist()
ship_number = df["운송장번호"].astype(str).tolist()

# ✅ 데이터 없으면 종료
if not order_number:
    log_warn("📭 오늘 처리할 29CM 교환 출고 주문이 없습니다. 스크립트를 종료합니다.")
    exit()

# ======================== 브라우저 세팅 (29CM) ========================

log_step("크롬 드라이버 실행 및 29CM 로그인 시작")

options = Options()
options.add_argument("--start-maximized")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920x1080")
# 필요시 헤드리스
# options.add_argument("--headless=new")

driver = webdriver.Chrome(options=options)

driver.set_page_load_timeout(60)
driver.get("https://partner-auth.29cm.co.kr/login")

# 로그인 폼 뜰 때까지 버전 무관 대기 (EC.any_of 없이 lambda)
WebDriverWait(driver, 20).until(
    lambda d: d.find_elements(By.XPATH, '//*[@id="__next"]/section/div/div[2]/form/div[1]/div[1]/div/input')
           or d.find_elements(By.XPATH, "//button[contains(.,'로그인')]")
)

# ======================== 5. 로그인 & 이메일 인증 ========================
log_step("29CM 로그인 시작")

try:
    driver.find_element(By.XPATH, '/html/body/div[2]/div/div/button').click()
except:
    pass
time.sleep(1.5)
driver.find_element(By.XPATH, '//*[@id="__next"]/section/div/div[2]/form/div[1]/div[1]/div/input').send_keys("verish")
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="__next"]/section/div/div[2]/form/div[1]/div[2]/div/input').send_keys("Deepdive1!")
time.sleep(1.5)
driver.find_element(By.XPATH, '//*[@id="__next"]/section/div/div[2]/form/button').click()
time.sleep(5)

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table')))
rows = driver.find_elements(By.XPATH, '//table/tbody/tr')

log_step("이강희 선택")
for row in rows:
    try:
        phone_cell = row.find_elements(By.TAG_NAME, 'td')[-1]
        if "010-****-1924" in phone_cell.text:
            row.find_element(By.TAG_NAME, 'label').click()
            log_info("이강희 선택 완료")
            break
    except Exception as e:
        log_error(f"이강희 선택 실패: {e}")

driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/div/form/button[2]').click()
time.sleep(7)

otp = get_latest_verification_code(email_user, app_password)
driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/div/form/div[2]/div/input').send_keys(otp)
time.sleep(1.3)
driver.find_element(By.XPATH, '//*[@id="__next"]/div[2]/div/form/button[2]').click()
time.sleep(4)


log_step("지정된 팝업 패턴 순차 처리 시작")

# 최대 5번까지 다른 팝업이 나타날 가능성을 염두에 두고 반복
for i in range(5):
    # 이번 반복에서 팝업을 처리했는지 여부를 확인하는 변수
    popup_handled = False

    # --- 패턴 1: '오늘 하루 동안 보지않기' 팝업 시도 ---
    try:
        # 팝업을 찾는 시간을 짧게(2초) 설정하여, 없으면 빨리 다음으로 넘어가도록 함
        wait = WebDriverWait(driver, 2)
        
        # 1-1. 체크박스 클릭
        checkbox_locator = (By.XPATH, "//label[contains(., '오늘 하루 동안 보지않기')]/div")
        wait.until(EC.element_to_be_clickable(checkbox_locator)).click()
        log_info("패턴1: '오늘 하루 동안 보지않기' 체크박스 클릭")
        time.sleep(1.5)

        # 1-2. 확인 버튼 클릭 (안정적인 경로로 수정)
        button_locator = (By.XPATH, "//footer[.//label[contains(., '오늘 하루 동안 보지않기')]]//button")
        wait.until(EC.element_to_be_clickable(button_locator)).click()
        log_info("패턴1: 확인 버튼 클릭")
        
        popup_handled = True

    except Exception:
        # 패턴1 팝업이 없으면 조용히 넘어감
        pass

    # --- 패턴 2: '다시 보지 않기' 팝업 시도 ---
    try:
        wait = WebDriverWait(driver, 2)
        
        # 2-1. 체크박스 클릭
        checkbox_locator = (By.XPATH, "//label[contains(., '다시 보지 않기')]/div")
        wait.until(EC.element_to_be_clickable(checkbox_locator)).click()
        log_info("패턴2: '다시 보지 않기' 체크박스 클릭")
        time.sleep(1.5)

        # 2-2. 닫기 버튼 클릭 (안정적인 경로로 수정)
        button_locator = (By.XPATH, "//div[.//label[contains(., '다시 보지 않기')]]//button[contains(.,'닫기')]")
        wait.until(EC.element_to_be_clickable(button_locator)).click()
        log_info("패턴2: 닫기 버튼 클릭")

        popup_handled = True
        
    except Exception:
        # 패턴2 팝업이 없으면 조용히 넘어감
        pass

    # 이번 반복에서 아무 팝업도 처리하지 않았다면, 더 이상 팝업이 없는 것이므로 루프 종료
    if not popup_handled:
        log_info("처리할 팝업을 더 이상 찾지 못했습니다.")
        break
    else:
        # 팝업을 하나 처리했으므로, 다음 팝업이 나타날 시간을 줌
        time.sleep(1.5)

log_step("모든 팝업 처리가 완료되었습니다.")


# ======================== 6. 메뉴 이동 ========================
log_step("29CM 메뉴 이동: 교환처리")

driver.find_element(By.ID, "주문/배송").click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="주문/배송"]/ul/div[6]/a').click()
time.sleep(2)



# ======================== 7. 교환 처리 자동입력 시작 ========================
log_step("교환 송장 자동입력 시작")

log_info(f"오늘 전체 출고건 수: {df.shape[0]}")
log_info(f"29CM 교환 주문 수: {len(order_number)}")

# ======================== 8. 주문번호별 송장 입력 ========================
try:
    for i, order in enumerate(order_number):
        log_step(f"[{i+1}/{len(order_number)}] 주문번호 처리: {order}")
        
        order_input = driver.find_element(By.XPATH, '//*[@id="__next"]/section/section/main/div[2]/div/div/div[2]/div/div[2]/div/div[2]/div/input')
        order_input.clear()
        order_input.send_keys(order)
        driver.find_element(By.XPATH, '//*[@id="__next"]/section/section/main/div[2]/div/div/div[3]/button[1]').click()
        time.sleep(3)

        try:
            status_elem = driver.find_element(By.XPATH, '//*[@id="__next"]/section/section/main/div[2]/section/div[1]/div[2]/div[1]/div/table/tbody/tr/td[2]/p')
            status_text = status_elem.text.strip()
            log_info(f"교환 상태: {status_text}")
        except NoSuchElementException:
            log_warn(f"상태 정보 없음 → 다음 주문으로 넘어감: {order}")
            continue  # 다음 주문번호로 넘어가기

        if status_text in ["교환검수 완료", "교환수거 완료", "교환 수거 중"]:
            driver.find_element(By.XPATH, '//*[@id="__next"]/section/section/main/div[2]/section/div[1]/div[2]/div[1]/div/table/tbody/tr/td[1]/span/label/div').click()
            log_info("체크박스 클릭 완료")
            time.sleep(1.5)

            driver.find_element(By.XPATH, '//*[@id="__next"]/section/section/main/div[2]/section/div[1]/div[1]/div[1]/button[3]').click()
            log_info("교환 출고 처리 클릭 완료")
            time.sleep(2)

            dropdown = driver.find_element(By.XPATH,'//*[@id="react-select-instance-react-select-id-placeholder"]')
            dropdown.click()
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).pause(0.5).send_keys(Keys.ARROW_UP).pause(0.5).send_keys(Keys.ENTER).perform()
            log_info("택배사 선택 완료")
            time.sleep(1.3)

            driver.find_element(By.XPATH, '/html/body/div[8]/div/div[1]/div/section/div/div[2]/div[2]/div/div/input').send_keys(ship_number[i])
            log_info(f"운송장 입력 완료: {ship_number[i]}")
            time.sleep(1.5)

            driver.find_element(By.XPATH, '/html/body/div[8]/div/div[2]/div/button[2]').click()
            log_info(f"교환 처리 완료: {order}")
            time.sleep(1.5)
        else:
            log_warn(f"상태 '{status_text}' 스킵됨: {order}")
except Exception as e:
    log_error(f"전체 루프 중단됨: {traceback.format_exc()}")
    driver.quit()
    exit()