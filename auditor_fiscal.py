import argparse
import os
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from functools import lru_cache
from html import escape
from pathlib import Path

import pyodbc
from dotenv import load_dotenv

from centros_de_custos import centros_de_custos_por_filial
from dados import filiais_ativas, filiais_ref
from datas import Datas
from enviar_email import remetente, destinatario, senha_app, enviar_email_com_erro_no_corpo

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    # Evita UnicodeEncodeError no console do Windows (cp1252) com os
    # caracteres usados no relatorio (ex.: ❌, ✅, 🏥).
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

PROJ_DIR = Path(__file__).resolve().parent
load_dotenv(PROJ_DIR / ".env")

TOLERANCIA = 0.02  # diferencas menores que isso sao arredondamento, nao divergencia
SEPARADOR = "=-" * 20
MOV_SIMPLES_NACIONAL = "1.2.14"  # "Entrada NF Simples Nacional"
FILIAIS_SEM_VALIDACAO_VENCIMENTO = {35}  # nao se aplica a regra de vencimento < data de criacao

FILIAL_ENVIO_OCULTO = {31, 45}  # quando esta filial tiver divergencia, envia relatorio isolado (oculto) abaixo
EMAIL_DIVERGENCIA_FILIAL_OCULTA = "validacaonfhb@ints.org.br"

COLUNAS = [
    "cd_filial", "ds_filial", "id_mov", "fluig", "nota", "dt_criacao", "dt_venc", "ds_cliente", "cnpj",
    "cd_ccusto", "ds_ccusto", "tipo_mov", "opt_simples",
    "valor_bruto",
    "base_iss", "aliq_iss", "iss",
    "base_inss", "aliq_inss", "inss",
    "base_csrf", "aliq_csrf", "csrf",
]

IMPOSTOS = (
    ("ISS", "base_iss", "aliq_iss", "iss"),
    ("INSS", "base_inss", "aliq_inss", "inss"),
    ("CSRF", "base_csrf", "aliq_csrf", "csrf"),
)

CATEGORIAS = ("ISS", "INSS", "CSRF", "Centro de Custo", "Simples Nacional", "Data de Vencimento")

MENSAGENS_OK = {
    "ISS": "✅ Todos os lançamentos de ISS estão consistentes.",
    "INSS": "✅ Todos os lançamentos de INSS estão consistentes.",
    "CSRF": "✅ Todos os lançamentos de CSRF estão consistentes.",
    "Centro de Custo": "✅ Todos os centros de custo estão cadastrados corretamente.",
    "Simples Nacional": f"✅ Nenhuma divergência de Simples Nacional (mov {MOV_SIMPLES_NACIONAL}) encontrada.",
    "Data de Vencimento": "✅ Nenhuma data de vencimento anterior à data de criação.",
}

ROTULOS_ERRO = {
    "Simples Nacional": f"Erros Simples Nacional (mov {MOV_SIMPLES_NACIONAL})",
}


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
    return pyodbc.connect(conn_str, timeout=15)


def buscar_lancamentos(conn: pyodbc.Connection, inicio: str, fim: str, filiais: list):
    placeholders = ",".join("?" for _ in filiais)
    query = (
        f"SELECT {', '.join(COLUNAS)} FROM v_EntradaNotas "
        f"WHERE TRY_CONVERT(date, dt_criacao, 105) BETWEEN ? AND ? AND cd_filial IN ({placeholders}) "
        f"ORDER BY cd_filial, dt_criacao"
    )
    cur = conn.cursor()
    cur.execute(query, [inicio, fim, *filiais])
    return cur.fetchall()


def _num(v) -> float:
    return float(v) if v is not None else 0.0


def divergencia(base, aliq, valor):
    """Retorna o valor calculated (base * aliquota) se ele divergir do valor
    lancado em mais de TOLERANCIA; caso contrario, retorna None (sem
    divergencia)."""
    calculado = _num(base) * _num(aliq)
    if abs(_num(valor) - calculado) >= TOLERANCIA:
        return calculado
    return None


@lru_cache(maxsize=None)
def centros_permitidos(filial_id) -> frozenset | None:
    """Codigos de centro de custo cadastrados para a filial em
    centros_de_custos.py. Retorna None se a filial nao tem lista
    cadastrada (nesse caso nao ha o que validar)."""
    chave = str(filial_id).zfill(2)
    lista = centros_de_custos_por_filial.get(chave)
    if lista is None:
        return None
    return frozenset(item.split(" - ", 1)[0].strip() for item in lista)


def centro_custo_divergente(row, permitidos: frozenset) -> bool:
    return (row.cd_ccusto or "").strip() not in permitidos


def simples_nacional_divergente(row) -> bool:
    tipo_mov = (row.tipo_mov or "").strip()
    opt_simples = (row.opt_simples or "").strip().upper()
    return tipo_mov == MOV_SIMPLES_NACIONAL and opt_simples != "SIM"


def _parse_data_br(v) -> date | None:
    """Converte string dd-mm-yyyy (formato usado por dt_criacao/dt_venc na
    view) em date; retorna None se vazio ou nao parseavel."""
    if not v:
        return None
    try:
        return datetime.strptime(str(v).strip(), "%d-%m-%Y").date()
    except ValueError:
        return None


def vencimento_divergente(row) -> bool:
    dt_criacao = _parse_data_br(row.dt_criacao)
    dt_venc = _parse_data_br(row.dt_venc)
    if dt_criacao is None or dt_venc is None:
        return False
    return dt_venc < dt_criacao


def _evento(categoria: str, row, detalhe: str) -> dict:
    return {
        "categoria": categoria,
        "filial_id": row.cd_filial,
        "fluig": row.fluig or "(vazio)",
        "fornecedor": (row.ds_cliente or "").strip(),
        "cnpj": row.cnpj or "",
        "idmov": row.id_mov,
        "data": row.dt_criacao,
        "detalhe": detalhe,
    }


def coletar_eventos(rows) -> list[dict]:
    """Varre os lancamentos uma unica vez e retorna um evento por
    divergencia encontrada (um lancamento pode gerar mais de um evento).
    E a fonte unica de verdade usada pelo relatorio em texto, pelo e-mail
    em HTML e pela planilha exportada."""
    eventos = []
    for row in rows:
        for imposto, campo_base, campo_aliq, campo_valor in IMPOSTOS:
            calc = divergencia(getattr(row, campo_base), getattr(row, campo_aliq), getattr(row, campo_valor))
            if calc is not None:
                valor = getattr(row, campo_valor)
                detalhe = f"Valor {imposto} R$ {_num(valor):.2f} ≠ valor calculado R$ {calc:.2f}"
                eventos.append(_evento(imposto, row, detalhe))

        permitidos = centros_permitidos(row.cd_filial)
        if permitidos is not None and centro_custo_divergente(row, permitidos):
            cc = (row.cd_ccusto or "").strip() or "(vazio)"
            ds = (row.ds_ccusto or "").strip()
            detalhe = f"Centro de custo {cc} - {ds} não está cadastrado no de-para"
            eventos.append(_evento("Centro de Custo", row, detalhe))

        if simples_nacional_divergente(row):
            detalhe = (
                f"Lançamento no movimento {MOV_SIMPLES_NACIONAL} (Entrada NF Simples Nacional) "
                f"para empresa NÃO optante pelo Simples (opt_simples={row.opt_simples!r})"
            )
            eventos.append(_evento("Simples Nacional", row, detalhe))

        if row.cd_filial not in FILIAIS_SEM_VALIDACAO_VENCIMENTO and vencimento_divergente(row):
            detalhe = f"Data de vencimento {row.dt_venc} é anterior à data de criação {row.dt_criacao}"
            eventos.append(_evento("Data de Vencimento", row, detalhe))

    return eventos


def formatar_evento_texto(ev: dict) -> str:
    partes = [f"Fluig: {ev['fluig']}", f"Fornecedor: {ev['fornecedor']}"]
    if ev["categoria"] == "Simples Nacional":
        partes.append(f"CNPJ: {ev['cnpj']}")
    partes.append(f"IDMOV: {ev['idmov']}")
    if ev["categoria"] != "Data de Vencimento":
        partes.append(f"Data de lançamento: {ev['data']}")
    return f" ❌ {ev['detalhe']} → " + " | ".join(partes)


def montar_relatorio(rows, filiais_alvo: list) -> tuple[str, list[dict]]:
    """Monta o relatorio em texto no mesmo formato do fluxo antigo: um
    bloco por filial, com checklist por categoria. Retorna o texto e a
    lista de eventos de divergencia encontrados."""
    por_filial = defaultdict(list)
    for row in rows:
        por_filial[row.cd_filial].append(row)

    eventos = coletar_eventos(rows)
    eventos_por_filial_categoria = defaultdict(list)
    for ev in eventos:
        eventos_por_filial_categoria[(ev["filial_id"], ev["categoria"])].append(ev)

    linhas = ["🚨 INÍCIO DAS VALIDAÇÕES\n"]

    for filial_id in filiais_alvo:
        nome = filiais_ref.get(filial_id, str(filial_id))
        linhas.append(f"🏥 {nome}")

        filial_rows = por_filial.get(filial_id, [])
        if not filial_rows:
            linhas.append("⚠️  Filial sem movimento no periodo selecionado.")
            linhas.append(SEPARADOR)
            linhas.append("")
            continue

        for categoria in CATEGORIAS:
            if categoria == "Centro de Custo" and centros_permitidos(filial_id) is None:
                linhas.append("⚠️  Filial sem centros de custo cadastrados no de-para.")
                linhas.append(SEPARADOR)
                continue

            if categoria == "Data de Vencimento" and filial_id in FILIAIS_SEM_VALIDACAO_VENCIMENTO:
                linhas.append("⚠️  Regra de data de vencimento não se aplica a esta filial.")
                linhas.append(SEPARADOR)
                continue

            erros = eventos_por_filial_categoria.get((filial_id, categoria), [])
            if erros:
                rotulo = ROTULOS_ERRO.get(categoria, f"Erros {categoria}")
                linhas.append(f"{rotulo}: {len(erros)}")
                linhas.extend(formatar_evento_texto(ev) for ev in erros)
            else:
                linhas.append(MENSAGENS_OK[categoria])
            linhas.append(SEPARADOR)

        linhas.append("")

    linhas.append("\n✅ FIM DAS VALIDAÇÕES")
    return "\n".join(linhas), eventos


def montar_email_html(eventos: list[dict], filiais_alvo: list, inicio: str, fim: str) -> str:
    """Monta um e-mail HTML enxuto: um resumo por categoria no topo e,
    abaixo, uma tabela por filial - somente para filiais com divergencia.
    Filiais sem problema entram apenas na contagem do resumo, mas a lista
    completa de filiais fica sempre visivel (nao usamos <details>/<summary>
    porque varios clientes de e-mail ignoram o estado fechado e renderizam
    tudo aberto, o que ficava mais feio do que simplesmente mostrar direto)."""
    total_filiais = len(filiais_alvo)
    por_categoria = Counter(ev["categoria"] for ev in eventos)
    filiais_afetadas = defaultdict(list)
    for ev in eventos:
        filiais_afetadas[ev["filial_id"]].append(ev)

    chips = []
    for filial_id in sorted(filiais_alvo):
        n = len(filiais_afetadas.get(filial_id, []))
        if n:
            estilo = "background:#fef2f2;border:1px solid #fecaca;color:#991b1b;"
        else:
            estilo = "background:#f0fdf4;border:1px solid #bbf7d0;color:#166534;"
        rotulo = str(filial_id)
        chips.append(
            f"<span style='display:inline-block;{estilo}border-radius:999px;padding:4px 10px;"
            f"margin:3px;font-size:12px;'>{rotulo}</span>"
        )

    bloco_filiais_lista = f"""
      <div style="margin:8px 0 16px;">
        <div style="font-size:13px;color:#2f6690;font-weight:600;margin-bottom:8px;">
          Todas as {total_filiais} filiais incluídas nesta auditoria
        </div>
        <div>{''.join(chips)}</div>
      </div>
    """

    linhas_resumo = "".join(
        f"<tr><td style='padding:6px 12px;border-bottom:1px solid #e5e7eb;'>{escape(cat)}</td>"
        f"<td style='padding:6px 12px;border-bottom:1px solid #e5e7eb;text-align:center;'>{n}</td></tr>"
        for cat, n in sorted(por_categoria.items(), key=lambda item: -item[1])
    )

    blocos_filiais = []
    for filial_id in sorted(filiais_afetadas, key=lambda fid: -len(filiais_afetadas[fid])):
        evs = filiais_afetadas[filial_id]
        nome = escape(filiais_ref.get(filial_id, str(filial_id)))
        linhas_tabela = "".join(
            "<tr>"
            f"<td style='padding:6px 8px;border-bottom:1px solid #e5e7eb;'>{escape(ev['categoria'])}</td>"
            f"<td style='padding:6px 8px;border-bottom:1px solid #e5e7eb;'>{escape(str(ev['fluig']))}</td>"
            f"<td style='padding:6px 8px;border-bottom:1px solid #e5e7eb;'>{escape(ev['fornecedor'])}</td>"
            f"<td style='padding:6px 8px;border-bottom:1px solid #e5e7eb;'>{escape(str(ev['idmov']))}</td>"
            f"<td style='padding:6px 8px;border-bottom:1px solid #e5e7eb;'>{escape(str(ev['data']))}</td>"
            f"<td style='padding:6px 8px;border-bottom:1px solid #e5e7eb;'>{escape(ev['detalhe'])}</td>"
            "</tr>"
            for ev in evs
        )
        blocos_filiais.append(f"""
        <h3 style="margin:24px 0 8px;font-size:15px;color:#1f2933;">
          🏥 {nome} <span style="color:#b91c1c;font-weight:normal;">({len(evs)} divergência{'s' if len(evs) != 1 else ''})</span>
        </h3>
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <thead>
            <tr style="background:#f3f4f6;text-align:left;">
              <th style="padding:6px 8px;border-bottom:2px solid #d1d5db;">Categoria</th>
              <th style="padding:6px 8px;border-bottom:2px solid #d1d5db;">Fluig</th>
              <th style="padding:6px 8px;border-bottom:2px solid #d1d5db;">Fornecedor</th>
              <th style="padding:6px 8px;border-bottom:2px solid #d1d5db;">IDMOV</th>
              <th style="padding:6px 8px;border-bottom:2px solid #d1d5db;">Data</th>
              <th style="padding:6px 8px;border-bottom:2px solid #d1d5db;">Detalhe</th>
            </tr>
          </thead>
          <tbody>{linhas_tabela}</tbody>
        </table>
        """)

    filiais_ok = total_filiais - len(filiais_afetadas)

    return f"""
    <div style="font-family:'Segoe UI',Arial,sans-serif;color:#1f2933;max-width:900px;">
      <h2 style="margin-bottom:4px;">📋 Auditoria Fiscal — Relatório de Divergências</h2>
      <p style="color:#52606d;margin-top:0;">Período: {escape(inicio)} a {escape(fim)}</p>

      <table style="border-collapse:collapse;margin:16px 0;">
        <thead>
          <tr>
            <th style="text-align:left;padding:6px 12px;background:#f3f4f6;border-bottom:2px solid #d1d5db;">Categoria</th>
            <th style="padding:6px 12px;background:#f3f4f6;border-bottom:2px solid #d1d5db;">Ocorrências</th>
          </tr>
        </thead>
        <tbody>
          {linhas_resumo}
          <tr>
            <td style="padding:6px 12px;font-weight:bold;">Total</td>
            <td style="padding:6px 12px;text-align:center;font-weight:bold;">{len(eventos)}</td>
          </tr>
        </tbody>
      </table>

      <p style="color:#52606d;">
        {len(filiais_afetadas)} filial(is) com divergência · {filiais_ok} filial(is) sem problema (não detalhadas abaixo).
      </p>

      {bloco_filiais_lista}

      {''.join(blocos_filiais)}
    </div>
    """


def gerar_planilha(rows, saida_dir: Path) -> Path:
    import pandas as pd

    saida_dir.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame.from_records([tuple(r) for r in rows], columns=COLUNAS)
    out = saida_dir / f"auditoria_fiscal_{date.today().isoformat()}.xlsx"
    df.to_excel(out, index=False)
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Auditor Fiscal - divergencias ISS/INSS/CSRF via ODS")
    parser.add_argument("--inicio", default=None, help="YYYY-MM-DD (default: dia 01 do mes corrente)")
    parser.add_argument("--fim", default=None, help="YYYY-MM-DD (default: ultimo dia do mes corrente)")
    parser.add_argument("--filiais", default=None, help="codigos de filial separados por virgula (default: filiais ativas)")
    parser.add_argument("--saida-dir", default="relatorios_Entradas/auditorias")
    parser.add_argument(
        "--sem-email", action="store_true",
        help=f"nao envia o e-mail para {destinatario} mesmo se houver divergencia"
    )
    args = parser.parse_args()

    inicio_mes, fim_mes = Datas.get()
    inicio = args.inicio or inicio_mes
    fim = args.fim or fim_mes
    filiais_alvo = (
        [int(x.strip()) for x in args.filiais.split(",") if x.strip().isdigit()]
        if args.filiais
        else list(filiais_ativas)
    )

    print(f"INFO: consultando lancamentos de {inicio} a {fim} ({len(filiais_alvo)} filial(is)), tolerancia R$ {TOLERANCIA:.2f}")

    conn = conectar()
    try:
        rows = buscar_lancamentos(conn, inicio, fim, filiais_alvo)
    finally:
        conn.close()

    print(f"INFO: {len(rows)} lancamento(s) no periodo, montando relatorio...")

    relatorio, eventos = montar_relatorio(rows, filiais_alvo)
    print(relatorio)

    total_divergencias = len(eventos)
    print(f"INFO: {total_divergencias} divergencia(s) encontrada(s) no total")

    if total_divergencias == 0:
        print("SUCCESS: nenhuma divergencia encontrada no periodo")
        return 0

    idmovs_divergentes = {ev["idmov"] for ev in eventos}
    linhas_divergentes = [row for row in rows if row.id_mov in idmovs_divergentes]
    saida = gerar_planilha(linhas_divergentes, Path(args.saida_dir))
    print(f"INFO: planilha gerada em {saida}")

    if args.sem_email:
        print("INFO: --sem-email: e-mail nao enviado.")
    else:
        try:
            corpo_html = montar_email_html(eventos, filiais_alvo, inicio, fim)
            enviar_email_com_erro_no_corpo(relatorio, remetente, senha_app, destinatario, corpo_html=corpo_html)
            print(f"INFO: e-mail enviado para {destinatario}")
        except Exception as exc:
            print(f"ERROR: falha ao enviar e-mail: {exc}")

        for filial_oculta in FILIAL_ENVIO_OCULTO:
            eventos_filial_oculta = [ev for ev in eventos if ev["filial_id"] == filial_oculta]
            if eventos_filial_oculta:
                try:
                    rows_filial_oculta = [row for row in rows if row.cd_filial == filial_oculta]
                    relatorio_oculto, _ = montar_relatorio(rows_filial_oculta, [filial_oculta])
                    corpo_html_oculto = montar_email_html(eventos_filial_oculta, [filial_oculta], inicio, fim)
                    enviar_email_com_erro_no_corpo(
                        relatorio_oculto, remetente, senha_app, EMAIL_DIVERGENCIA_FILIAL_OCULTA,
                        corpo_html=corpo_html_oculto,
                        assunto=f"📋 Relatório de Erros - Validações (Filial {filial_oculta})",
                    )
                    print(
                        f"INFO: e-mail oculto de divergencias da filial {filial_oculta} "
                        f"enviado para {EMAIL_DIVERGENCIA_FILIAL_OCULTA}"
                    )
                except Exception as exc:
                    print(f"ERROR: falha ao enviar e-mail oculto da filial {filial_oculta}: {exc}")

    print(f"SUCCESS: auditoria concluida - {total_divergencias} divergencia(s) encontrada(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())