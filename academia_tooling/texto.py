"""Normalização de texto compartilhada pelos verificadores."""

from __future__ import annotations

import re
import unicodedata

# Partículas que não são sobrenome em referências lusófonas/estrangeiras.
PARTICULAS = frozenset(
    "de da do das dos e van von del della di la le los las du des da's o a os as".split()
)


def sem_acento(texto: str) -> str:
    """Remove diacríticos (NFKD), mantendo o resto do texto."""
    decomposto = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in decomposto if not unicodedata.combining(c))


def chave(texto: str) -> str:
    """Chave de comparação de nomes: sem acento, minúscula, tokens colapsados."""
    t = re.sub(r"[^a-z0-9]+", " ", sem_acento(texto or "").casefold())
    return re.sub(r"\s+", " ", t).strip()


def limpar_latex(texto: str) -> str:
    """Remove comandos LaTeX/BibTeX preservando a letra-base do acento.

    `Opera{\\c c}{\\~o}es` -> `Operacoes`: sem isso, o mesmo título escrito com e
    sem escape não casa na deduplicação.
    """
    return re.sub(r"\\[A-Za-z]+|\\.", "", texto or "")


def chave_titulo(texto: str) -> str:
    """Chave de comparação de títulos: só letras/dígitos, sem acento, minúscula.

    Serve para casar 'Vigiar e Punir.' com 'VIGIAR E PUNIR' e com
    'Vigiar e punir' escapado em LaTeX.
    """
    return re.sub(r"[^a-z0-9]", "", sem_acento(limpar_latex(texto or "")).casefold())


def numero(valor: float, casas: int = 2) -> str:
    """Número com vírgula decimal, sem zeros à direita (pt-BR, para o laudo)."""
    texto = f"{float(valor):.{casas}f}".rstrip("0").rstrip(".")
    return (texto or "0").replace(".", ",")


def contar_palavras(texto: str) -> int:
    return len(re.findall(r"\S+", texto or ""))


def achatar(texto: str, limite: int = 160) -> str:
    """Uma linha só, truncada com reticências — para caber no laudo."""
    t = re.sub(r"\s+", " ", (texto or "").strip())
    return t if len(t) <= limite else t[: limite - 1].rstrip() + "…"
