import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

# 엑셀 파일 불러오기
mail_df = pd.read_excel('C:/hanwoori_wave_mail.xlsx')  # 메일 주소가 포함된 파일
number_df = pd.read_excel('C:/hanwoori_wave_number.xlsx')  # 쿠폰 번호가 포함된 파일


# Naver SMTP를 사용하여 이메일을 보내고 터미널에 상태를 출력하는 함수
def send_email_via_naver(email, number, mail_user, mail_password):
    try:
        # SMTP 서버 설정
        smtp_server = "smtp.naver.com"
        smtp_port = 587

        # 이메일 제목
        subject = "[한우리 X 아삭 X 웨이브] 웨이브 이벤트 쿠폰 당첨 안내"

        # HTML 형식의 이메일 본문
        # HTML 형식의 이메일 본문
        body = (f"""
        <html>
        <body>
            <p style="color: black !important;"> ☝️ 하나에서, 우리로! 🙌 </p>
            <p style="color: black !important;">안녕하세요! 한신대학교 학생 홍보대사 <strong><span style="color:purple;">💜한우리 21기💜</span></strong>입니다.</p>
            <p style="color: black !important;">먼저, 한우리 X 아삭 X 웨이브 🍁‘아삭! 너 한우리랑 Wavve 탈래?’🍁 이벤트에 참여해 주셔서 감사합니다.</p>
            <p style="color: black !important;">이벤트에 당첨되셔서 <strong>Wavve 2개월 이용권 쿠폰 번호</strong>를 보내드립니다.</p>
            <br><br>
            <div style="border: 1px solid black !important; padding: 40px 10px !important; text-align: center !important; width: 50% !important;">
                쿠폰 번호는 <strong>'{number}'</strong> 입니다.
            </div>
            <br><br>
            <p style="color: black !important;">앞으로 한우리 21기 많은 관심 부탁드립니다. 감사합니다.</p>
            <br><br>
            <p style="color: black !important;"><strong>[유의 사항]</strong></p>
            <ul style="color: black !important;">
                <li>본 쿠폰은 앱을 통한 쿠폰 등록이 불가하며, PC 혹은 웹을 통한 우회 방식으로 등록을 해주시길 바랍니다.</li>
                <li>본 쿠폰은 2개월 제한이며, 2개월 만료 후 해지하지 않을 시 자동 결제됩니다.</li>
                <li>문의사항💜💌은 한우리 오픈 채팅방을 이용해 주세요. https://open.kakao.com/o/sDObB70g </li>
            </ul>
        </body>
        </html>
        """)

        # MIMEMultipart 객체 생성
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = mail_user
        msg["To"] = email

        # HTML 본문 추가
        msg.attach(MIMEText(body, "html"))

        # SMTP 서버에 연결하여 이메일 보내기
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # TLS 연결 시작
            server.login(mail_user, mail_password)  # SMTP 로그인
            server.send_message(msg)  # 이메일 전송

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
