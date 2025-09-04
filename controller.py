from gerar_relatorios import main
import os
from enviar_email import capturar_erros_validacao, enviar_email_com_erro_no_corpo, remetente, destinatario, senha_app
from dados import filiais_ref


if __name__ == "__main__":
    print("🚀 Iniciando geração de relatórios...\n")

    out_dir = main()
    
        

    print("\n✅ Geração concluída. Iniciando validações e envio de e-mail...\n")

    erros = capturar_erros_validacao() 
   

    if erros.strip():
        enviar_email_com_erro_no_corpo(erros, remetente, senha_app, destinatario)
    else:
        print("✅ Nenhum erro encontrado. Nenhum email enviado.")
    
    try:
        for arquivo in os.listdir(out_dir):
            caminho_completo = os.path.join(out_dir, arquivo)
            if os.path.isfile(caminho_completo) and arquivo.endswith(".xlsx"):
                os.remove(caminho_completo)
                print("🧹 Planilha removida do diretório.")
    except Exception as e:
        print(f"❌ Erro ao limpar diretório:{e}")

    


    print("\n🏁 Processo finalizado.")