Selenium과 Google Sheets API를 활용하여 29CM와 지그재그 파트너 어드민 교환 송장 등록 업무를 자동화한 프로젝트입니다.

주요 기능
- Google Sheets API로 교환 출고 데이터 조회
- Selenium WebDriver로 29CM/지그재그 관리자 페이지 자동 로그인 & OTP 인증
- 주문번호 검색 후 교환 상태 확인
- 조건 충족 시 자동으로 운송장 입력 및 교환 처리 완료

기술 스택
- Python
- Selenium WebDriver
- Google Sheets API (gspread, oauth2client)
- dotenv (환경변수 관리)
- pandas (데이터 가공)

보안 처리
- 계정(ID, PW), Google API Key, 이메일 계정 등은 .env 파일로 관리합니다.
- .gitignore에 .env, 키 파일(*.json)은 반드시 추가하여 공개되지 않도록 했습니다.


긴 XPATH 대신 placeholder, 버튼 텍스트 기반 단순화된 XPATH으로 바꿔 넣었습니다