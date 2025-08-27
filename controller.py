from gerar_relatorios import gerar_relatorio, output_dir,filiais_nomes,filial_centros
import os
from enviar_email import capturar_erros_validacao, enviar_email_com_erro_no_corpo, remetente, destinatario, senha_app
from dados import filiais_ref


if __name__ == "__main__":
    print("🚀 Iniciando geração de relatórios...\n")

    for filial_id, centros in filial_centros.items():
        nome_filial = filiais_nomes.get(filial_id)
        filial_ref = filiais_ref.get(filial_id)
        if not nome_filial:
            print(f"❌ Nome da filial {filial_id} não encontrado em filiais_nomes.")
            continue

        output_file = os.path.join(output_dir, f"{filial_ref}.xlsx")
        gerar_relatorio(filial_id, nome_filial, centros, output_file)

    print("\n✅ Geração concluída. Iniciando validações e envio de e-mail...\n")

    erros = capturar_erros_validacao() 
   

    if erros.strip():
        enviar_email_com_erro_no_corpo(erros, remetente, senha_app, destinatario)
    else:
        print("✅ Nenhum erro encontrado. Nenhum email enviado.")
    
    try:
        for arquivo in os.listdir(output_dir):
            caminho_completo = os.path.join(output_dir, arquivo)
            if os.path.isfile(caminho_completo) and arquivo.endswith(".xlsx"):
                os.remove(caminho_completo)
                print("🧹 Planilha removida do diretório.")
    except Exception as e:
        print(f"❌ Erro ao limpar diretório:{e}")

    


    print("\n🏁 Processo finalizado.")