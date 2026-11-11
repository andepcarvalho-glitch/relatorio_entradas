"""Identifica movimentos de entrada cujos lancamentos financeiros ja foram baixados.

A baixa fica no financeiro do RM (tabela FLAN, a mesma que alimenta a v_Pagamentos):
FLAN.STATUSLAN = 1 -> baixado (equivale a v_Pagamentos.DATABAIXA preenchida).
Um movimento so e considerado baixado quando TODOS os seus lancamentos (parcelas)
estao baixados; se houver alguma parcela em aberto, ele continua no relatorio.
"""
STATUSLAN_BAIXADO = 1
_LOTE = 500


def movimentos_baixados(conn, id_movs) -> set:
    ids = sorted({int(i) for i in id_movs if i is not None})
    baixados = set()
    cur = conn.cursor()
    for i in range(0, len(ids), _LOTE):
        lote = ids[i:i + _LOTE]
        placeholders = ",".join("?" for _ in lote)
        cur.execute(
            f"SELECT IDMOV FROM FLAN WHERE CODCOLIGADA = 1 AND IDMOV IN ({placeholders}) "
            f"GROUP BY IDMOV HAVING MIN(CASE WHEN STATUSLAN = ? THEN 1 ELSE 0 END) = 1",
            [*lote, STATUSLAN_BAIXADO],
        )
        baixados.update(r[0] for r in cur.fetchall())
    return baixados


def remover_baixados(conn, rows) -> tuple[list, int]:
    """Retorna (linhas sem os movimentos baixados, quantidade de movimentos removidos)."""
    baixados = movimentos_baixados(conn, (r.id_mov for r in rows))
    return [r for r in rows if r.id_mov not in baixados], len(baixados)
