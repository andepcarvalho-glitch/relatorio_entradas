import os
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

remetente = os.environ.get("SMTP_USER", "intsfiscalbot@gmail.com")
destinatario = os.environ.get("SMTP_DESTINATARIO", "fiscal@ints.org.br")
senha_app = os.environ.get("SMTP_PASSWORD", "")


def enviar_email_com_erro_no_corpo(corpo_email, remetente, senha_app, destinatario, corpo_html=None, assunto=None):
    if not senha_app:
        raise RuntimeError("SMTP_PASSWORD nao definido em .env (senha de app do Gmail).")

    msg = EmailMessage()
    msg['Subject'] = assunto or "📋 Relatório de Erros - Validações"
    msg['From'] = remetente
    msg['To'] = destinatario
    msg.set_content(corpo_email)
    if corpo_html:
        msg.add_alternative(corpo_html, subtype="html")

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(remetente, senha_app)
        smtp.send_message(msg)

    print("📧 Email enviado com sucesso!")
