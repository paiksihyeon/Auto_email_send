import smtplib
from email.mime.text import MIMEText
import pandas as pd

# 엑셀 파일 불러오기
mail_df = pd.read_excel('C:/mail.xlsx')  # 메일 주소가 포함된 파일
number_df = pd.read_excel('C:/number.xlsx')  # 쿠폰 번호가 포함된 파일

# Naver SMTP를 사용하여 이메일을 보내고 터미널에 상태를 출력하는 함수
def send_email_via_naver(email, number, mail_user, mail_password):
    try:
        # SMTP 서버 설정
        #smtp_server = "smtp.gmail.com"
        smtp_server = "smtp.naver.com"
        smtp_port = 587

        # 이메일 메시지 생성
        subject = "한우리 X 아삭 X 웨이브 웨이브 쿠폰 번호 당첨 안내"
        body = ("안녕하세요. 이벤트에 참여해주셔서 감사합니다.\n\n"
                "쿠폰 번호는 '{number}'입니다.\n\n").format(number=number)
        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = mail_user
        msg["To"] = email

        # SMTP 서버에 연결하여 이메일 보내기
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(mail_user, mail_password)
            server.send_message(msg)

        # 성공 상태 반환
        print(f"메일 주소: {email}, 당첨 번호: {number}, 메일 발송: 성공")
        return True  # 성공 시 True 반환
    except Exception as e:
        # 실패 상태 반환
        print(f"메일 주소: {email}, 당첨 번호: {number}, 메일 발송: 실패 ({e})")
        return False  # 실패 시 False 반환

# 모든 수신자에게 이메일 보내는 함수
def send_emails(mail_df, number_df, mail_user, mail_password):
    success_count = 0  # 성공 횟수
    failure_count = 0  # 실패 횟수
    failed_emails = []  # 실패한 이메일 목록

    for email, number in zip(mail_df['mail'], number_df['number']):
        if send_email_via_naver(email, number, mail_user, mail_password):
            success_count += 1  # 성공 시 성공 횟수 증가
        else:
            failure_count += 1  # 실패 시 실패 횟수 증가
            failed_emails.append(email)  # 실패한 이메일 추가

    # 최종 결과 출력
    print("\n총 발송 성공 횟수:", success_count)
    print("총 발송 실패 횟수:", failure_count)
    if failed_emails:
        print("실패한 이메일 목록:", failed_emails)

# 계정 정보로 이메일 발송 실행
send_emails(mail_df, number_df, '@naver.com', '')  # 이메일, 앱 2차비밀번호 사용

