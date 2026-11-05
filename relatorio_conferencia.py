"""Relatorio de Conferencia de Notas (RELATORIO DE REGISTRO DE ENTRADAS) por filial.

Replica o relatorio exportado do 3Bi: mesmo layout (modelos/modelo_conferencia_notas.xlsx),
uma linha por movimento x centro de custo x produto da view v_EntradaNotas, filtrando pela
data de criacao do movimento (dt_criacao) dentro do periodo.

Uso:
    python relatorio_conferencia.py --inicio 2026-01-01 --fim 2026-01-19 --filiais 45
    python relatorio_conferencia.py                      # mes corrente, filiais ativas
"""
import argparse
import os
import re
import sys
from collections import Counter
from copy import copy
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import openpyxl
import pyodbc
from dotenv import load_dotenv
from openpyxl.styles.cell_style import StyleArray

from dados import filiais_ativas
from datas import Datas

PROJ_DIR = Path(__file__).resolve().parent
load_dotenv(PROJ_DIR / ".env")

MODELO = PROJ_DIR / "modelos" / "modelo_conferencia_notas.xlsx"
SAIDA_PADRAO = PROJ_DIR / "relatorios_Entradas" / "conferencia"

LINHA_DADOS = 17          # primeira linha de dados no modelo (tambem e a linha de estilo)
ALTURA_LINHA = 12.6
ULTIMA_COL = 39           # AM
MESCLAS_LINHA = ("A{0}:B{0}", "F{0}:G{0}", "L{0}:M{0}", "AJ{0}:AK{0}")
COLUNAS_MESCLADAS = {"A": "B", "F": "G", "L": "M", "AJ": "AK"}   # inicio da mescla -> 2a coluna

# coluna do relatorio -> coluna da view (colunas B, G, M, AK ficam dentro das mesclas)
COLUNAS = {
    "A": "ds_ccusto", "C": "fluig", "D": "usuario", "E": "dt_criacao", "F": "tipo_mov",
    "H": "mov", "I": "id_mov", "J": "nota", "K": "dt_emissao", "L": "cnpj",
    "N": "ds_cliente", "O": "municipio", "P": "opt_simples", "Q": "natrendimento_csrf",
    "R": "natrendimento_irrf", "S": "cd_produto", "T": "ds_produto", "U": "dt_venc",
    "V": "valor_bruto", "W": "base_iss", "X": "aliq_iss", "Y": "iss", "Z": "base_inss",
    "AA": "aliq_inss", "AB": "inss", "AC": "aliq_csrf", "AD": "csrf", "AE": "valor_irrf",
    "AF": "valor_liquido", "AG": "iss_calculado", "AH": "Conf_iss", "AI": "inss_calculado",
    "AJ": "Conf_inss", "AL": "csrf_calculado", "AM": "Conf_csrf",
}
CAMPOS = ["cd_filial", "ds_filial", "cd_depart", *dict.fromkeys(COLUNAS.values())]


def conectar() -> pyodbc.Connection:
    host = os.environ.get("ODS_DB_HOST")
    database = os.environ.get("ODS_DB_NAME")
    user = os.environ.get("ODS_DB_USER")
    password = os.environ.get("ODS_DB_PASSWORD")
    if not all([host, database, user, password]):
        raise SystemExit(
            "ERROR: defina ODS_DB_HOST, ODS_DB_NAME, ODS_DB_USER e ODS_DB_PASSWORD em .env "
            "(veja .env.example)."
        )
    conn_str = f"DRIVER={{SQL Server}};SERVER={host};DATABASE={database};UID={user};PWD={password};"
    conn = pyodbc.connect(conn_str, timeout=15)
    # Colunas varchar usam collation Latin1_General; sem isso acentos viram '?'.
    conn.setdecoding(pyodbc.SQL_CHAR, encoding="cp1252")
    return conn


def buscar_lancamentos(conn: pyodbc.Connection, inicio: str, fim: str, filiais: list):
    placeholders = ",".join("?" for _ in filiais)
    query = (
        f"SELECT {', '.join(CAMPOS)} FROM v_EntradaNotas "
        f"WHERE TRY_CONVERT(date, dt_criacao, 105) BETWEEN ? AND ? AND cd_filial IN ({placeholders})"
    )
    cur = conn.cursor()
    cur.execute(query, [inicio, fim, *filiais])
    return cur.fetchall()


def _valor(v):
    """Como no 3Bi: numeros como numero (inteiro quando nao tem centavos), texto vazio = celula vazia."""
    if isinstance(v, Decimal):
        f = float(v)
        return int(f) if f == int(f) else f
    if isinstance(v, str):
        return v if v.strip() else None
    return v


def _soma(rows, campo) -> float:
    return round(sum(float(getattr(r, campo) or 0) for r in rows), 2)


def gerar_planilha(rows, inicio: date, fim: date, saida_dir: Path) -> Path:
    wb = openpyxl.load_workbook(MODELO)
    ws = wb.active

    # Estilos da linha de dados e do rodape (linhas 18 e 19 do modelo), antes de mexer nelas.
    # colunas vazias na linha do modelo (ex.: Valor Liquido) usam o estilo comum dos dados (Valor Bruto)
    padrao = ws[f"V{LINHA_DADOS}"]._style
    estilo_dados = {c: copy(ws.cell(LINHA_DADOS, c)._style if ws.cell(LINHA_DADOS, c).has_style else padrao)
                    for c in range(1, ULTIMA_COL + 1)}
    rodape = [{c: copy(ws.cell(r, c)._style) for c in range(1, ULTIMA_COL + 1)} for r in (18, 19)]
    texto_rodape = ws["A19"].value
    for m in [m for m in ws.merged_cells.ranges if m.min_row >= LINHA_DADOS]:
        ws.unmerge_cells(str(m))
    for r in (LINHA_DADOS, 18, 19):
        for c in range(1, ULTIMA_COL + 1):
            ws.cell(r, c).value = None
            ws.cell(r, c)._style = StyleArray()

    # Cabecalho
    primeira = rows[0]
    ws["I1"] = datetime.combine(date.today(), datetime.min.time())
    ws["C5"] = primeira.ds_filial
    ws["I5"] = str(primeira.cd_filial)
    ws["C6"] = Counter(r.cd_depart for r in rows).most_common(1)[0][0]
    ws["C8"] = datetime.combine(inicio, datetime.min.time())
    ws["E8"] = datetime.combine(fim, datetime.min.time())
    ws["W14"] = _soma(rows, "iss")
    ws["X14"] = _soma(rows, "inss")
    ws["Y14"] = _soma(rows, "valor_irrf")

    # Dados
    for i, row in enumerate(rows):
        n = LINHA_DADOS + i
        for col, campo in COLUNAS.items():
            valor = _valor(getattr(row, campo))
            # como no 3Bi: valor nulo nao gera celula (fica sem formatacao), exceto nas mesclas
            if valor is None and col not in COLUNAS_MESCLADAS:
                continue
            cell = ws[f"{col}{n}"]
            cell.value = valor
            cell._style = copy(estilo_dados[cell.column])
        for col in COLUNAS_MESCLADAS.values():
            ws[f"{col}{n}"]._style = copy(estilo_dados[openpyxl.utils.column_index_from_string(col)])
        ws.row_dimensions[n].height = ALTURA_LINHA
        for mescla in MESCLAS_LINHA:
            ws.merge_cells(mescla.format(n))
    ultima = LINHA_DADOS + len(rows) - 1

    # Rodape: linha em branco + "Powered by 3Bi"
    for k, (estilos, altura) in enumerate(zip(rodape, (21.0, 7.9)), start=1):
        n = ultima + k
        for c, estilo in estilos.items():
            ws.cell(n, c)._style = copy(estilo)
        ws.row_dimensions[n].height = altura
    ws.cell(ultima + 2, 1).value = texto_rodape
    ws.merge_cells(f"A{ultima + 2}:AI{ultima + 2}")

    ws.auto_filter.ref = f"A16:AM{ultima}"

    saida_dir.mkdir(parents=True, exist_ok=True)
    nome = f"Relatório de Conferencia de Notas {inicio:%d.%m.%Y} À {fim:%d.%m.%Y}.xlsx"
    pasta = re.sub(r'[\\/:*?"<>|]', "-", f"{primeira.cd_filial} - {str(primeira.ds_filial).strip()}")
    out = saida_dir / pasta / nome
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def main() -> int:
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="Relatorio de Conferencia de Notas (Registro de Entradas) via ODS")
    parser.add_argument("--inicio", default=None, help="YYYY-MM-DD (default: dia 01 do mes corrente)")
    parser.add_argument("--fim", default=None, help="YYYY-MM-DD (default: ultimo dia do mes corrente)")
    parser.add_argument("--filiais", default=None, help="codigos de filial separados por virgula (default: filiais ativas)")
    parser.add_argument("--saida", default=str(SAIDA_PADRAO), help="pasta de saida (uma subpasta por filial)")
    args = parser.parse_args()

    inicio_padrao, fim_padrao = Datas.get()
    inicio = date.fromisoformat(args.inicio or inicio_padrao)
    fim = date.fromisoformat(args.fim or fim_padrao)
    filiais = [int(f) for f in args.filiais.split(",")] if args.filiais else filiais_ativas

    with conectar() as conn:
        rows = buscar_lancamentos(conn, inicio.isoformat(), fim.isoformat(), filiais)

    por_filial: dict[int, list] = {}
    for r in rows:
        por_filial.setdefault(int(r.cd_filial), []).append(r)

    if not por_filial:
        print(f"Nenhum lancamento entre {inicio:%d/%m/%Y} e {fim:%d/%m/%Y}.")
        return 0
    for filial in sorted(por_filial):
        out = gerar_planilha(por_filial[filial], inicio, fim, Path(args.saida))
        print(f"Filial {filial}: {len(por_filial[filial])} linhas -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
