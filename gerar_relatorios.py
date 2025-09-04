# -*- coding: utf-8 -*-
"""
Geração de relatórios Pentaho (v6)
- Obtém valores válidos de 'centro_custo' (XML/JSON) por múltiplos endpoints
- CLI (--inicio, --fim, --filiais, --saida-dir) para facilitar testes
- Loga e salva dumps de respostas para troubleshooting
- Casamento por CÓDIGO (mantém seu dicionário local intocado)
- Execução em /generatedContent com polling do job (XLSX)

Dependência: requests
"""

import os
import re
import time
import json
import argparse
import requests
from urllib.parse import quote
from xml.etree import ElementTree as ET
from datas import Datas
from dados import filiais_ativas, filiais_nomes,filiais_ref
from centros_de_custos import centros_de_custos_por_filial

# Credenciais (pode usar variáveis de ambiente)
USUARIO = os.environ.get("PENTAHO_USER", "anderson.carvalho")
SENHA   = os.environ.get("PENTAHO_PASS", "And14050")

# Host/relatório (pode sobrescrever via env se quiser)
HOST = os.environ.get("PENTAHO_HOST", "http://138.59.147.111:8080")
REPORT_PATH = os.environ.get("PENTAHO_REPORT_PATH", ":home:fiscal:report:in:Relatório de Conferencia de Notas.prpt")
REPORT_PATH_LEGACY = os.environ.get("PENTAHO_REPORT_PATH_LEGACY", "/home/fiscal/report/in/Relatório de Conferencia de Notas.prpt")

LOGIN_URL = f"{HOST}/pentaho/j_spring_security_check"
REPORT_URL = f"{HOST}/pentaho/api/repos/{quote(REPORT_PATH, safe='')}/generatedContent"
PARAMETER_UI_URL = f"{HOST}/pentaho/api/repos/{quote(REPORT_PATH, safe='')}/parameterUi"
PARAMETER_JSON_URL = f"{HOST}/pentaho/api/repos/{quote(REPORT_PATH, safe='')}/parameter"
PLUGIN_PARAMETER_URL = f"{HOST}/pentaho/plugin/reporting/api/parameter"
PLUGIN_PARAMETER_VALUES_URL = f"{HOST}/pentaho/plugin/reporting/api/parameter/values"

DEFAULT_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
    "X-Requested-With": "XMLHttpRequest",
}

def _split_code(label: str) -> str:
    parts = label.split(" - ", 1)
    return parts[0].strip() if parts else label.strip()

def _dedup(seq):
    seen, out = set(), []
    for x in seq:
        if x not in seen:
            out.append(x); seen.add(x)
    return out

def _parse_parameter_json(data):
    out = []
    if isinstance(data, dict):
        if "values" in data and isinstance(data["values"], list):
            for item in data["values"]:
                if isinstance(item, dict):
                    v = item.get("value") or item.get("label")
                    if v:
                        out.append(v)
        for container_key in ("parameter", "parameters"):
            params = data.get(container_key) or []
            for p in params:
                pname = p.get("name") or p.get("parameterName")
                if (pname or "").lower() == "centro_custo":
                    vals = p.get("values") or p.get("selections") or []
                    for item in vals:
                        if isinstance(item, dict):
                            v = item.get("value") or item.get("label")
                            if v:
                                out.append(v)
    return _dedup(out)

def _parse_parameter_xml(xml_text: str):
    out = []
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return out

    for param in root.iter():
        tag = (param.tag or "").lower()
        if tag.endswith("parameter"):
            name = (param.attrib.get("name") or param.attrib.get("parameterName") or "").lower()
            if name == "centro_custo":
                for node in param.iter():
                    ntag = (node.tag or "").lower()
                    if ntag.endswith("value"):
                        v = node.attrib.get("value")
                        if v:
                            out.append(v)
                        if node.text and node.text.strip():
                            out.append(node.text.strip())
                        for child in node:
                            if child.text and child.text.strip():
                                out.append(child.text.strip())

                for node in param.iter():
                    labels = [c.text.strip() for c in node if (c.tag.lower().endswith("label") and c.text and c.text.strip())]
                    values = [c.text.strip() for c in node if (c.tag.lower().endswith("value") and c.text and c.text.strip())]
                    for v in values or labels:
                        out.append(v)

    if not out:
        for node in root.iter():
            if (node.tag or "").lower().endswith("value"):
                if node.text and node.text.strip():
                    out.append(node.text.strip())
                v = node.attrib.get("value")
                if v:
                    out.append(v)
    return _dedup(out)

def _extract_values_by_content_type(resp, dump_base: str):
    # Salva SEMPRE o corpo bruto para facilitar debug
    ext = ".bin"
    try:
        ct = resp.headers.get("Content-Type", "")
        if "json" in ct.lower():
            ext = ".json"
        elif "xml" in ct.lower():
            ext = ".xml"
        elif "html" in ct.lower():
            ext = ".html"
    except Exception:
        pass
    try:
        with open(dump_base + ext, "wb") as f:
            f.write(resp.content)
    except Exception:
        pass

    ct = resp.headers.get("Content-Type", "").lower()
    if "application/json" in ct or "text/json" in ct or ct.endswith("/json"):
        try:
            data = resp.json()
            return _parse_parameter_json(data)
        except Exception:
            return []
    else:
        # trata como XML/HTML
        try:
            return _parse_parameter_xml(resp.text)
        except Exception:
            return []

def login(user: str, pwd: str) -> requests.Session:
    s = requests.Session()
    s.headers.update(DEFAULT_HEADERS)
    r = s.post(LOGIN_URL, data={"j_username": user, "j_password": pwd})
    r.raise_for_status()
    if "Login" in r.text and "pentaho" in r.text.lower():
        raise RuntimeError("Falha de autenticação no Pentaho.")
    return s

def get_allowed_centros(session: requests.Session, nome_filial: str, dt_inic: str, dt_fina: str, out_dir: str):
    form = {
        "renderMode": "PARAMETER",
        "showParameters": "true",
        "ds_filial": nome_filial,
        "dt_inic": dt_inic,
        "dt_fina": dt_fina,
        "htmlProportionalWidth": "false",
        "accepted-page": "0",
        "ts": str(int(time.time()*1000)),
    }

    # 1) /api/repos/.../parameter (POST)
    r = session.post(PARAMETER_JSON_URL, data=form, headers={**DEFAULT_HEADERS, "Accept": "*/*"})
    if r.ok:
        vals = _extract_values_by_content_type(r, os.path.join(out_dir, "debug_parameter_repos_post"))
        if vals:
            return vals

    # 2) /api/repos/.../parameter (GET)
    r = session.get(PARAMETER_JSON_URL, params=form, headers={**DEFAULT_HEADERS, "Accept": "*/*"})
    if r.ok:
        vals = _extract_values_by_content_type(r, os.path.join(out_dir, "debug_parameter_repos_get"))
        if vals:
            return vals

    # 3) /plugin/reporting/api/parameter (GET)
    params = {
        "path": REPORT_PATH_LEGACY,
        "renderMode": "PARAMETER",
        "showParameters": "true",
        "ds_filial": nome_filial,
        "dt_inic": dt_inic,
        "dt_fina": dt_fina,
        "ts": form["ts"],
    }
    r = session.get(PLUGIN_PARAMETER_URL, params=params, headers={**DEFAULT_HEADERS, "Accept": "*/*"})
    if r.ok:
        vals = _extract_values_by_content_type(r, os.path.join(out_dir, "debug_parameter_plugin_get"))
        if vals:
            return vals

    # 4) /plugin/reporting/api/parameter/values (GET)
    params = {
        "path": REPORT_PATH_LEGACY,
        "parameterName": "centro_custo",
        "renderMode": "PARAMETER",
        "ds_filial": nome_filial,
        "dt_inic": dt_inic,
        "dt_fina": dt_fina,
        "ts": form["ts"],
    }
    r = session.get(PLUGIN_PARAMETER_VALUES_URL, params=params, headers={**DEFAULT_HEADERS, "Accept": "*/*"})
    if r.ok:
        vals = _extract_values_by_content_type(r, os.path.join(out_dir, "debug_parameter_values_plugin_get"))
        if vals:
            return vals

    # 5) /api/repos/.../parameterUi (POST)
    r = session.post(PARAMETER_UI_URL, data=form, headers=DEFAULT_HEADERS)
    if r.ok:
        vals = _extract_values_by_content_type(r, os.path.join(out_dir, "debug_parameterUi_html"))
        if vals:
            return vals

    print("⚠️ Não consegui extrair opções de 'centro_custo'. Consulte os dumps em", out_dir)
    return []

def casar_por_codigo(payload_canonico, allowed_values):
    map_allowed = {}
    for v in allowed_values:
        code = _split_code(v)
        if code not in map_allowed:
            map_allowed[code] = v
    postar, faltando = [], []
    for label in payload_canonico:
        code = _split_code(label)
        if code in map_allowed:
            postar.append(map_allowed[code])
        else:
            faltando.append(label)
    return postar, faltando, map_allowed

def baixar_via_job(session: requests.Session, first_response: requests.Response, output_file: str, timeout_s: int = 240) -> bool:
    if first_response.status_code in (301, 302) and "Location" in first_response.headers:
        loc = first_response.headers["Location"]
        status_url = loc if loc.startswith("http") else HOST + loc

        start = time.time()
        while True:
            if time.time() - start > timeout_s:
                raise TimeoutError("Timeout aguardando término do job de relatório.")
            s = session.get(status_url, headers=DEFAULT_HEADERS)
            s.raise_for_status()
            try:
                info = s.json()
            except Exception:
                info = {}
            state = (info.get("state") or "").upper()
            if state == "FINISHED":
                m = re.search(r"/jobs/([^/]+)/", status_url)
                job_id = info.get("jobId") or (m.group(1) if m else None)
                if not job_id:
                    raise RuntimeError("Não foi possível inferir jobId.")
                break
            if state in ("ERROR", "FAILED"):
                raise RuntimeError(f"Falha no job: {info}")
            time.sleep(2)

        content_url = f"{HOST}/pentaho/plugin/reporting/api/jobs/{job_id}/content?output-target=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;page-mode=flow"
        file = session.get(content_url, headers=DEFAULT_HEADERS)
        file.raise_for_status()
        with open(output_file, "wb") as f:
            f.write(file.content)
        print(f"✅ Relatório salvo: {output_file}")
        return True
    return False

def gerar_relatorio(session: requests.Session, filial_id: int, nome_filial: str, centros: list, output_file: str, dt_inic: str, dt_fina: str, out_dir: str):
    print(f"\n=== Filial {filial_id:02d} | {nome_filial} ===")
    allowed = get_allowed_centros(session, nome_filial, dt_inic, dt_fina, out_dir)
    print(f"Valores 'centro_custo' disponíveis no período: {len(allowed)}")
    if allowed:
        # Loga alguns exemplos
        for ex in allowed[:5]:
            print("  •", ex)

    centros_post, faltando, map_allowed = casar_por_codigo(centros, allowed)

    # Log de casamento por código
    print(f"Casados por código: {len(centros_post)} | Faltando no período: {len(faltando)}")
    if faltando:
        for x in faltando[:10]:
            print("  - Não aceito neste período:", x)
        if len(faltando) > 10:
            print(f"  ... e mais {len(faltando)-10}")

    if not centros_post:
        print(f"❌ Nenhum centro aceito para filial {filial_id}. Verifique dumps em {out_dir}.")
        return

    payload = [
        ("ds_filial", nome_filial),
        ("dt_inic", dt_inic),
        ("dt_fina", dt_fina),
        ("renderMode", "REPORT"),
        ("output-target", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;page-mode=flow"),
        ("download", "true"),
    ] + [("centro_custo", cc) for cc in centros_post]

    resp = session.post(REPORT_URL, data=payload, allow_redirects=False)

    if baixar_via_job(session, resp, output_file):
        return

    if resp.status_code == 200:
        cd = resp.headers.get("Content-Disposition", "")
        if "filename=" in cd or resp.headers.get("Content-Type", "").startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            with open(output_file, "wb") as f:
                f.write(resp.content)
            print(f"✅ Relatório salvo: {output_file}")
            return

    dbg = os.path.join(out_dir, f"debug_exec_filial_{filial_id:02d}.html")
    with open(dbg, "wb") as f:
        f.write(resp.content)
    print(f"❌ Resposta inesperada para filial {filial_id}. HTML salvo em {dbg}")

def main():
    out_dir, args = getOut_dir()

    if args.inicio and args.fim:
        dt_inic, dt_fina = args.inicio, args.fim
    else:
        dt_inic, dt_fina = Datas.get()

    # Login
    session = login(USUARIO, SENHA)

    # Filiais alvo
    if args.filiais:
        alvo = [int(x.strip()) for x in args.filiais.split(",") if x.strip().isdigit()]
    else:
        alvo = list(filiais_ativas)

    # Mapeia filiais ativas para seus centros
    for filial_id in alvo:
        chave = str(filial_id).zfill(2)
        centros = centros_de_custos_por_filial.get(chave, [])
        if not centros:
            print(f"⚠️ Filial {chave} ({filiais_nomes.get(filial_id, 'Nome desconhecido')}) sem centros configurados.")
            continue
        nome_filial = filiais_nomes.get(filial_id, f"{filial_id}")
        output_file = os.path.join(out_dir, f"{filiais_ref[filial_id]}.xlsx")
        gerar_relatorio(session, filial_id, nome_filial, centros, output_file, dt_inic, dt_fina, out_dir)
    
    return out_dir


def getOut_dir():
    parser = argparse.ArgumentParser(description="Gerador de relatórios Pentaho (v6)")
    parser.add_argument("--inicio", help="Data inicial (YYYY-MM-DDTHH:MM:SS.mmm). Default: Datas.get()", default=None)
    parser.add_argument("--fim", help="Data final (YYYY-MM-DDTHH:MM:SS.mmm). Default: Datas.get()", default=None)
    parser.add_argument("--filiais", help="IDs de filiais separados por vírgula. Default: filiais_ativas", default=None)
    parser.add_argument("--saida-dir", help="Diretório de saída. Default: relatorios_Entradas", default="relatorios_Entradas")
    
    args = parser.parse_args()

    out_dir = args.saida_dir
    os.makedirs(out_dir, exist_ok=True)

    return out_dir, args

