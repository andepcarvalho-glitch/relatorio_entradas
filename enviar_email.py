import os
import io
import smtplib
import ssl
from email.message import EmailMessage
from validar_relatorio import validar_todos_os_impostos
from gerar_relatorios import getOut_dir

remetente = "intsfiscalbot@gmail.com"
destinatario = "fiscal@ints.org.br"
senha_app = "fnyg iubd qafm avkl"

def capturar_erros_validacao():
    buffer = io.StringIO()
    original_stdout = os.sys.stdout
    os.sys.stdout = buffer

    try:
        print("🚨 INÍCIO DAS VALIDAÇÕES\n")
        out_dir, args = getOut_dir()
        result = validar_todos_os_impostos(out_dir)
        
        
        for i in result:

            filial = str(i).replace(".xlsx", "")
            
            print(f"🏥 {filial}")

            try:
                print(result[i]['AVISO'])
                print("=-"*20)
                print('\n')
            except:

                erros_iss = result[i]['ISS']
                erros_inss = result[i]['INSS']
                erros_csrf = result[i]['CSRF']
                    
                if erros_iss:
                    print(f"Erros ISS: {len(erros_iss)}")
                    for erro in erros_iss:
                        print(erro)
                else:
                    print("✅ Todos os lançamentos de ISS estão consistentes.")
                print("=-"*20)

                
                if erros_inss:
                    print(f"Erros INSS: {len(erros_inss)}")
                    for erro in erros_inss:
                        print(erro)
                else:
                    print("✅ Todos os lançamentos de INSS estão consistentes.")
                print("=-"*20)

                
                if erros_csrf:
                    print(f"Erros CSRF: {len(erros_csrf)}")
                    for erro in erros_csrf:
                        print(erro)
                else:
                    print("✅ Todos os lançamentos de CSRF estão consistentes.")
                print("=-"*20)

                print('\n')
 
        print("\n✅ FIM DAS VALIDAÇÕES")
    finally:
        os.sys.stdout = original_stdout

    return buffer.getvalue()

def enviar_email_com_erro_no_corpo(corpo_email, remetente, senha_app, destinatario):
    msg = EmailMessage()
    msg['Subject'] = "📋 Relatório de Erros - Validações"
    msg['From'] = remetente
    msg['To'] = destinatario
    msg.set_content(corpo_email)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(remetente, senha_app)
        smtp.send_message(msg)


    print("📧 Email enviado com sucesso!")

