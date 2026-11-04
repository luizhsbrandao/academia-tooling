"""Parser BibTeX tolerante (sem dependência externa).

Objetivo: ler `.bib` real — com chaves repetidas, campos quebrados em várias
linhas, valores entre `{}` ou `""`, `@string`, comentários — e devolver as
entradas mais os problemas de sintaxe com a linha onde ocorreram.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

_NOME_CAMPO = re.compile(r"[A-Za-z_][A-Za-z0-9_\-:.]*")
_TIPOS_IGNORADOS = frozenset({"comment", "preamble", "string"})


@dataclass(frozen=True)
class Entrada:
    tipo: str
    chave: str
    campos: dict[str, str] = field(default_factory=dict)
    linha: int = 0

    def campo(self, *nomes: str) -> str:
        for nome in nomes:
            valor = self.campos.get(nome)
            if valor:
                return valor
        return ""

    def tem(self, *nomes: str) -> bool:
        return bool(self.campo(*nomes))

    def ano(self) -> int | None:
        bruto = self.campo("year", "date")
        casamento = re.search(r"\b(1[0-9]\d{2}|20\d{2})\b", bruto)
        return int(casamento.group(1)) if casamento else None

    def doi(self) -> str:
        return normalizar_doi(self.campo("doi"))

    def identificador(self) -> str:
        return self.doi() or self.campo("url").strip()


@dataclass(frozen=True)
class ProblemaSintaxe:
    linha: int
    mensagem: str


def normalizar_doi(valor: str) -> str:
    doi = (valor or "").strip()
    doi = re.sub(r"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", doi, flags=re.I)
    return doi.rstrip(" .,;").lower()


def _ler_corpo(texto: str, pos: int) -> tuple[str | None, int, str | None]:
    """Lê o interior de `{...}` ou `(...)` a partir do delimitador em `pos`."""
    abre = texto[pos]
    fecha = "}" if abre == "{" else ")"
    profundidade = 0
    i = pos
    em_aspas = False
    while i < len(texto):
        c = texto[i]
        if c == "\\":
            i += 2
            continue
        if abre == "{":
            if c == "{":
                profundidade += 1
            elif c == "}":
                profundidade -= 1
                if profundidade == 0:
                    return texto[pos + 1 : i], i + 1, None
        else:
            if c == '"':
                em_aspas = not em_aspas
            elif c == "(" and not em_aspas:
                profundidade += 1
            elif c == ")" and not em_aspas:
                profundidade -= 1
                if profundidade == 0:
                    return texto[pos + 1 : i], i + 1, None
        i += 1
    return None, len(texto), "entrada não fechada (faltou delimitador final)"


def _desembrulhar(valor: str) -> str:
    """Tira as camadas externas de {} / "" e colapsa espaços/quebras de linha."""
    v = valor.strip()
    while len(v) >= 2:
        if v[0] == "{" and v[-1] == "}":
            interno = v[1:-1]
            if interno.count("{") == interno.count("}"):
                v = interno.strip()
                continue
        if v[0] == '"' and v[-1] == '"' and v.count('"') == 2:
            v = v[1:-1].strip()
            continue
        break
    return re.sub(r"\s+", " ", v).strip()


def _ler_valor(texto: str, pos: int, macros: dict[str, str]) -> tuple[str, int]:
    """Lê um valor (com `#` de concatenação) até a vírgula de nível 0."""
    partes: list[str] = []
    i = pos
    profundidade = 0
    em_aspas = False
    atual = []
    while i < len(texto):
        c = texto[i]
        if c == "\\":
            atual.append(texto[i : i + 2])
            i += 2
            continue
        if c == "{":
            profundidade += 1
        elif c == "}":
            profundidade -= 1
        elif c == '"' and profundidade == 0:
            em_aspas = not em_aspas
        elif c == "," and profundidade == 0 and not em_aspas:
            break
        elif c == "#" and profundidade == 0 and not em_aspas:
            partes.append("".join(atual))
            atual = []
            i += 1
            continue
        atual.append(c)
        i += 1
    partes.append("".join(atual))

    limpos: list[str] = []
    for parte in partes:
        bruto = _desembrulhar(parte)
        if bruto.casefold() in macros:
            limpos.append(macros[bruto.casefold()])
        else:
            limpos.append(bruto)
    return "".join(limpos).strip(), i


def _parsear_campos(corpo: str, linha: int, macros: dict[str, str]) -> tuple[dict[str, str], list[ProblemaSintaxe]]:
    campos: dict[str, str] = {}
    problemas: list[ProblemaSintaxe] = []
    i = 0
    n = len(corpo)
    while i < n:
        while i < n and corpo[i] in " \t\r\n,":
            i += 1
        if i >= n:
            break
        casamento = _NOME_CAMPO.match(corpo, i)
        if not casamento:
            problemas.append(ProblemaSintaxe(linha, f"campo inválido perto de {corpo[i:i + 30]!r}"))
            break
        nome = casamento.group(0).casefold()
        i = casamento.end()
        while i < n and corpo[i] in " \t\r\n":
            i += 1
        if i >= n or corpo[i] != "=":
            problemas.append(ProblemaSintaxe(linha, f"campo {nome!r} sem '='"))
            break
        i += 1
        while i < n and corpo[i] in " \t\r\n":
            i += 1
        valor, i = _ler_valor(corpo, i, macros)
        if nome in campos:
            problemas.append(ProblemaSintaxe(linha, f"campo {nome!r} repetido na mesma entrada (fica o último)"))
        campos[nome] = valor
    return campos, problemas


def parsear(texto: str) -> tuple[list[Entrada], list[ProblemaSintaxe]]:
    entradas: list[Entrada] = []
    problemas: list[ProblemaSintaxe] = []
    macros: dict[str, str] = {}
    i = 0
    n = len(texto)
    while i < n:
        arroba = texto.find("@", i)
        if arroba < 0:
            break
        linha = texto.count("\n", 0, arroba) + 1
        j = arroba + 1
        while j < n and (texto[j].isalpha() or texto[j] in "_-"):
            j += 1
        tipo = texto[arroba + 1 : j].strip().casefold()
        while j < n and texto[j].isspace():
            j += 1
        if not tipo or j >= n or texto[j] not in "{(":
            if tipo:
                problemas.append(ProblemaSintaxe(linha, f"entrada @{tipo} sem corpo entre chaves"))
            i = max(j, arroba + 1)
            continue
        corpo, fim, erro = _ler_corpo(texto, j)
        if erro:
            problemas.append(ProblemaSintaxe(linha, f"@{tipo}: {erro}"))
            i = fim
            continue
        assert corpo is not None
        if tipo in _TIPOS_IGNORADOS:
            if tipo == "string":
                partes = corpo.split("=", 1)
                if len(partes) == 2:
                    nome = partes[0].strip().casefold()
                    valor, _ = _ler_valor(partes[1], 0, macros)
                    macros[nome] = valor
            i = fim
            continue

        chave, _, resto = corpo.partition(",")
        chave = chave.strip()
        if not chave:
            problemas.append(ProblemaSintaxe(linha, f"@{tipo} sem chave de citação"))
            i = fim
            continue
        campos, probs = _parsear_campos(resto, linha, macros)
        problemas.extend(probs)
        entradas.append(Entrada(tipo=tipo, chave=chave, campos=campos, linha=linha))
        i = fim
    return entradas, problemas


def ler(caminho: str | Path) -> tuple[list[Entrada], list[ProblemaSintaxe]]:
    return parsear(Path(caminho).read_text(encoding="utf-8", errors="replace"))
