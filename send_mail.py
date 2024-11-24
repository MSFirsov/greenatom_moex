import smtplib

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


your_email = 'your_email@mail.ru'
your_password = 'your_password'


def send_mail(text, file):

    smtp_server = smtplib.SMTP("smtp.mail.ru", 587)
    smtp_server.starttls()
    smtp_server.login(your_email, your_password)

    # Создание объекта сообщения
    msg = MIMEMultipart()

    # Настройка параметров сообщения
    msg["From"] = your_email
    msg["To"] = your_email
    msg["Subject"] = "moex info"

    # Добавление текста в сообщение
    # text = text_massage
    msg.attach(MIMEText(text, 'plain'))

    # Добавляем вложение (файл Excel)
    with open(f'{file}', 'rb') as f:
        attach = MIMEApplication(f.read(), _subtype='xlsx')
        attach.add_header('Content-Disposition', 'attachment', filename=f'{file}')
        msg.attach(attach)

    # Отправка письма
    smtp_server.sendmail(your_email, your_email, msg.as_string())

    # Закрытие соединения
    smtp_server.quit()


