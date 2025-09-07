"""
29CM 교환 송장 자동화 스크립트
---------------------------------
- Google Sheets에서 교환 출고 데이터를 읽어와 주문번호/운송장번호 추출
- Selenium을 통해 29CM 파트너 어드민 로그인 및 OTP 인증 자동화
- 주문번호 검색 후 교환 상태 확인 및 운송장 자동 입력
- 반복 업무 자동화로 처리 속도 및 정확성 개선
"""

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
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import get_as_dataframe

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

# ======================== 3. OTP FUNCTION ========================
def get_latest_verification_code(email_user, app_password):
    """이메일에서 6자리 OTP 코드 추출"""
    imap = imaplib.IMAP4_SSL('imap.gmail.com')
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

# ======================== 4. GOOGLE SHEETS ========================
def get_today_df(spreadsheet, sheet_name):
    worksheet = spreadsheet.worksheet(sheet_name)
    df = get_as_dataframe(worksheet, evaluate_formulas=True)
    df.dropna(how="all", inplace=True)
    today = datetime.date.today().strftime('%Y-%m-%d')
    return df[df['출고일'] == today]

# ======================== 5. MAIN ========================
if __name__ == "__main__":
    load_dotenv()

    EMAIL_USER = os.getenv("EMAIL_USER")
    APP_PASSWORD = os.getenv("APP_PASSWORD")
    ACCOUNT_ID = os.getenv("ACCOUNT_ID")        # 🟢 아이디 환경변수
    ACCOUNT_PW = os.getenv("ACCOUNT_PW")        # 🟢 비밀번호 환경변수
    GOOGLE_KEY_PATH = os.getenv("GOOGLE_KEY")   # 🟢 서비스 계정 키 경로
    SHEET_URL = os.getenv("SHEET_URL")

    # ✅ Google Sheets 인증
    log_step("Google Sheets 데이터 확인")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_KEY_PATH, scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(SHEET_URL)

    df1 = get_today_df(spreadsheet, "(외부몰)교환출고 raw")
    df2 = get_today_df(spreadsheet, "(불량)교환출고 raw")
    df = pd.concat([df1, df2], ignore_index=True)
    df = df[df["교환형태"] == "29CM"].drop_duplicates(subset="주문번호", keep="first")

    order_number = df["주문번호"].astype(str).tolist()
    ship_number = df["운송장번호"].astype(str).tolist()

    if not order_number:
        log_warn("📭 오늘 처리할 29CM 교환 출고 주문이 없습니다. 스크립트를 종료합니다.")
        exit()

    # ======================== 브라우저 세팅 ========================
    log_step("크롬 드라이버 실행 및 로그인")
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    driver.get("https://partner-auth.29cm.co.kr/login")

    # ======================== 로그인 ========================
    WebDriverWait(driver, 20).until(
        lambda d: d.find_elements(By.XPATH, '//input[@type="text"]')
    )
    driver.find_element(By.XPATH, '//input[@type="text"]').send_keys(ACCOUNT_ID)
    driver.find_element(By.XPATH, '//input[@type="password"]').send_keys(ACCOUNT_PW)
    driver.find_element(By.XPATH, "//button[contains(.,'로그인')]").click()
    time.sleep(5)

    # ======================== OTP 인증 ========================
    otp = get_latest_verification_code(EMAIL_USER, APP_PASSWORD)
    driver.find_element(By.XPATH, "//input[@type='number']").send_keys(otp)
    driver.find_element(By.XPATH, "//button[contains(.,'확인')]").click()
    time.sleep(5)

    # ======================== 메뉴 이동 ========================
    log_step("교환처리 메뉴 이동")
    driver.find_element(By.ID, "주문/배송").click()
    time.sleep(2)
    driver.find_element(By.XPATH, '//*[@id="주문/배송"]/ul/div[6]/a').click()
    time.sleep(2)

    # ======================== 교환 처리 자동입력 ========================
    log_step("교환 송장 자동입력 시작")
    log_info(f"오늘 전체 출고건 수: {df.shape[0]}")
    log_info(f"29CM 교환 주문 수: {len(order_number)}")

    try:
        for i, order in enumerate(order_number):
            log_step(f"[{i+1}/{len(order_number)}] 주문번호 처리: {order}")

            order_input = driver.find_element(By.XPATH, '//input[@placeholder="주문번호 입력"]')
            order_input.clear()
            order_input.send_keys(order)
            driver.find_element(By.XPATH, "//button[contains(.,'검색')]").click()
            time.sleep(3)

            try:
                status_elem = driver.find_element(By.XPATH, '//table/tbody/tr/td[2]/p')
                status_text = status_elem.text.strip()
                log_info(f"교환 상태: {status_text}")
            except NoSuchElementException:
                log_warn(f"상태 정보 없음 → 다음 주문으로 넘어감: {order}")
                continue

            if status_text in ["교환검수 완료", "교환수거 완료", "교환 수거 중"]:
                driver.find_element(By.XPATH, '//table/tbody/tr/td[1]//label/div').click()
                log_info("체크박스 클릭 완료")
                time.sleep(1)

                driver.find_element(By.XPATH, "//button[contains(.,'교환 출고 처리')]").click()
                time.sleep(2)

                dropdown = driver.find_element(By.XPATH, '//*[@id="react-select-instance-react-select-id-placeholder"]')
                dropdown.click()
                ActionChains(driver).send_keys(Keys.ARROW_DOWN).pause(0.5).send_keys(Keys.ARROW_UP).pause(0.5).send_keys(Keys.ENTER).perform()
                log_info("택배사 선택 완료")

                driver.find_element(By.XPATH, '//input[@placeholder="운송장번호 입력"]').send_keys(ship_number[i])
                log_info(f"운송장 입력 완료: {ship_number[i]}")

                driver.find_element(By.XPATH, "//button[contains(.,'확인')]").click()
                log_info(f"교환 처리 완료: {order}")
                time.sleep(1.5)
            else:
                log_warn(f"상태 '{status_text}' 스킵됨: {order}")

    except Exception:
        log_error(f"전체 루프 중단됨: {traceback.format_exc()}")
        driver.quit()
        exit()
