"""Laudo: um resultado por checagem, com três estados (PASS/FAIL/NAO_VERIFICAVEL).

Regra central: `NAO_VERIFICAVEL` não é sucesso. Só o limite citável em
`normas.json` autoriza PASS/FAIL; sem limite, o veredito é não verificável.
"""

from __future__ import annotations

import json
import textwrap
from collections import Counter
from dataclasses import asdict, dataclass, field
from typing import Any, Iterable, Sequence

PASS = "PASS"
FAIL = "FAIL"
NAO_VERIFICAVEL = "NAO_VERIFICAVEL"
STATUS_VALIDOS = (PASS, FAIL, NAO_VERIFICAVEL)

_RANK = {PASS: 0, NAO_VERIFICAVEL: 1, FAIL: 2}

# códigos de saída dos CLIs
COD_OK = 0
COD_FAIL = 1
COD_USO = 2
COD_NAO_VERIFICAVEL_ESTRITO = 3


def pior(*status: str) -> str:
    """Combina sub-vereditos: FAIL > NAO_VERIFICAVEL > PASS."""
    if not status:
        return PASS
    return max(status, key=lambda s: _RANK[s])


@dataclass(frozen=True)
class Resultado:
    id: str
    titulo: str
    status: str
    mensagem: str
    detalhes: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if self.status not in STATUS_VALIDOS:
            raise ValueError(f"status inválido: {self.status!r} (use {STATUS_VALIDOS})")

    def como_dict(self) -> dict[str, Any]:
        return asdict(self)


def contar(resultados: Iterable[Resultado]) -> Counter:
    return Counter(r.status for r in resultados)


def codigo_saida(resultados: Sequence[Resultado], estrito: bool = False) -> int:
    """0 sem FAIL; 1 com FAIL; 3 se estrito e houver NAO_VERIFICAVEL."""
    if any(r.status == FAIL for r in resultados):
        return COD_FAIL
    if estrito and any(r.status == NAO_VERIFICAVEL for r in resultados):
        return COD_NAO_VERIFICAVEL_ESTRITO
    return COD_OK


def como_json(*, contexto: dict[str, Any], resultados: Sequence[Resultado]) -> str:
    return json.dumps(
        {
            "contexto": contexto,
            "resumo": dict(contar(resultados)),
            "checagens": [r.como_dict() for r in resultados],
        },
        ensure_ascii=False,
        indent=2,
    )


def _pad(texto: str, largura: int) -> str:
    texto = texto if len(texto) <= largura else texto[: largura - 1] + "…"
    return texto.ljust(largura)


def tabela(
    cabecalho: Sequence[str],
    linhas: Sequence[Sequence[str]],
    maximos: Sequence[int] | None = None,
) -> list[str]:
    """Tabela alinhada em texto puro (sem dependência externa)."""
    larguras = [len(c) for c in cabecalho]
    for linha in linhas:
        for i, celula in enumerate(linha):
            larguras[i] = max(larguras[i], len(str(celula)))
    tetos = list(maximos) if maximos else [24] * len(cabecalho)
    larguras = [min(w, teto) for w, teto in zip(larguras, tetos)]
    saida = ["  ".join(_pad(c, w) for c, w in zip(cabecalho, larguras)).rstrip()]
    saida.append("  ".join("─" * w for w in larguras))
    for linha in linhas:
        saida.append("  ".join(_pad(str(c), w) for c, w in zip(linha, larguras)).rstrip())
    return saida


REGUA = "═" * 78
REGUA_FINA = "─" * 78


def imprimir_laudo(
    *,
    titulo: str,
    campos: Sequence[tuple[str, str]],
    resultados: Sequence[Resultado],
    rodape: Sequence[str] = (),
    largura_mensagem: int = 60,
) -> None:
    """Imprime o laudo humano: cabeçalho, tabela de checagens, resumo."""
    print(REGUA)
    print(f" {titulo}")
    print(REGUA_FINA)
    for rotulo, valor in campos:
        print(f" {rotulo:<10} {valor}")
    print(REGUA_FINA)

    linhas: list[list[str]] = []
    for r in resultados:
        pedacos = textwrap.wrap(r.mensagem, largura_mensagem) or [""]
        linhas.append([r.id, r.status, pedacos[0]])
        for extra in pedacos[1:]:
            linhas.append(["", "", extra])
    print("\n".join(tabela(["CHECAGEM", "STATUS", "MENSAGEM"], linhas, maximos=[18, 15, largura_mensagem])))

    print(REGUA_FINA)
    contagem = contar(resultados)
    total = len(resultados)
    citaveis = sum(1 for r in resultados if r.status != NAO_VERIFICAVEL)
    print(
        f" {total} checagens: {contagem.get(PASS, 0)} PASS · "
        f"{contagem.get(FAIL, 0)} FAIL · {contagem.get(NAO_VERIFICAVEL, 0)} NAO_VERIFICAVEL"
    )
    print(f" com veredito (não NAO_VERIFICAVEL): {citaveis}/{total}")
    for linha in rodape:
        print(f" {linha}")
    print(REGUA)
