import requests
import os
import time
from datas import Datas
from dados import filiais_ativas, filiais_nomes
from centros_de_custos import centros_de_custos_por_filial


usuario = "anderson.carvalho"
senha = "And14050"

# definindo variaveis datas de acordo com o return em datas.py
data_inicio, data_fim = Datas.get()

#  URLS PENTAHO
host = "http://ec2-56-124-62-174.sa-east-1.compute.amazonaws.com:8080"
login_url = f"{host}/pentaho/j_spring_security_check"
report_url = f"{host}/pentaho/api/repos/%3Ahome%3Afiscal%3Areport%3Ain%3ARelat%C3%B3rio%20de%20Conferencia%20de%20Notas.prpt/generatedContent"

# login de acordo com as variaveis definidas acima
session = requests.Session()
login_payload = {
    "j_username": usuario,
    "j_password": senha
}
login_response = session.post(login_url, data=login_payload)

if login_response.status_code != 200 or "Login" in login_response.text:
    print("❌ Falha ao autenticar no Pentaho.")
    exit()

# preparar os centros de custos importados de acordo com as filiais definidas em dados e os centors de custos trazidos de centros_de_custos.py
filial_centros = {}
for filial_id in filiais_ativas:
    chave = str(filial_id).zfill(2)
    if chave in centros_de_custos_por_filial:
        centros = centros_de_custos_por_filial[chave]
        filial_centros[filial_id] = centros
    else:
        print(f"⚠️ Filial {chave} ({filiais_nomes.get(filial_id, 'Nome desconhecido')}) não encontrada em centros_de_custos")

# armazenamento dos relatórios
output_dir = "relatorios_Entradas"
os.makedirs(output_dir, exist_ok=True)

# definição de como será montado o payload (Layout do "formulario" que é transmitido para pentaho para que nos retorne o download do R.E)
def gerar_relatorio(filial_id, nome_filial, centros, output_file):
    payload = [
        ("ts", str(int(time.time() * 1000))),
        ("ds_filial", nome_filial),
        ("dt_inic", data_inicio),
        ("dt_fina", data_fim),
        ("output-target", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;page-mode=flow"),
        ("accepted-page", "0"),
        ("showParameters", "true"),
        ("renderMode", "REPORT"),
        ("htmlProportionalWidth", "false"),
        ("query-limit-ui-enabled", "true"),
        ("query-limit", "0"),
        ("maximum-query-limit", "0"),
        ("changedParameters", "dt_fina")
    ]

    for cc in centros:
        payload.append(("centro_custo", cc))

    response = session.post(report_url, data=payload)

    if response.status_code == 200:
        if b"<html" in response.content[:100].lower() and b"parameters" in response.content.lower():
            print(f"❌ Parâmetros inválidos para filial {filial_id}. Salvando debug HTML.")
            debug_path = os.path.join(output_dir, f"debug_filial_{filial_id:02d}.html")
            with open(debug_path, "wb") as f:
                f.write(response.content)
        else:
            with open(output_file, "wb") as f:
                f.write(response.content)
            print(f"✅ Relatório salvo: {output_file}")
    else:
        print(f"❌ Erro HTTP {response.status_code} para filial {filial_id}: {response.text}")


