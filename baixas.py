"""Identifica movimentos de entrada que ja foram quitados (baixados) no financeiro.

Usa o status do movimento no RM: TMOV.STATUS = 'Q' (Quitado), que o RM marca quando
todos os lancamentos financeiros do movimento estao baixados.

Nao usar FLAN.STATUSLAN: na replica (ODS) ha lancamentos com baixa registrada na
FLANBAIXA e o movimento ja quitado, mas STATUSLAN ainda 0 (em aberto). Ex.: IDMOV 666377.
"""
STATUS_QUITADO = "Q"
_LOTE = 500


def movimentos_baixados(conn, id_movs) -> set:
    ids = sorted({int(i) for i in id_movs if i is not None})
    baixados = set()
    cur = conn.cursor()
    for i in range(0, len(ids), _LOTE):
        lote = ids[i:i + _LOTE]
        placeholders = ",".join("?" for _ in lote)
        cur.execute(
            f"SELECT IDMOV FROM TMOV WHERE CODCOLIGADA = 1 AND STATUS = ? AND IDMOV IN ({placeholders})",
            [STATUS_QUITADO, *lote],
        )
        baixados.update(r[0] for r in cur.fetchall())
    return baixados


def remover_baixados(conn, rows) -> tuple[list, int]:
    """Retorna (linhas sem os movimentos baixados, quantidade de movimentos removidos)."""
    baixados = movimentos_baixados(conn, (r.id_mov for r in rows))
    return [r for r in rows if r.id_mov not in baixados], len(baixados)
