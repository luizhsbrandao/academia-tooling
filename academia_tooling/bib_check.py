"""Auditor de bibliografia .bib.

Severidades: ERRO (quebra o pacote de submissão) e AVISO (higiene bibliográfica).
Só ERRO muda o código de saída, salvo `--estrito`.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from .bib_parse import Entrada, ProblemaSintaxe, ler, normalizar_doi, parsear
from .texto import achatar, chave, chave_titulo

ERRO = "ERRO"
AVISO = "AVISO"

ANO_MIN_PADRAO = 1500
ANO_MAX_PADRAO = date.today().year + 1

# Cada slot é o conjunto de campos que pode satisfazer a exigência.
REQUERIDOS: dict[str, tuple[tuple[str, ...], ...]] = {
    "article": (("author", "editor"), ("title",), ("journal", "journaltitle"), ("year", "date")),
    "book": (("author", "editor"), ("title",), ("publisher",), ("year", "date")),
    "booklet": (("title",), ("year", "date")),
    "inbook": (("author", "editor"), ("title",), ("publisher",), ("year", "date")),
    "incollection": (("author", "editor"), ("title",), ("booktitle",), ("publisher",), ("year", "date")),
    "inproceedings": (("author",), ("title",), ("booktitle",), ("year", "date")),
    "conference": (("author",), ("title",), ("booktitle",), ("year", "date")),
    "phdthesis": (("author",), ("title",), ("school", "institution"), ("year", "date")),
    "mastersthesis": (("author",), ("title",), ("school", "institution"), ("year", "date")),
    "techreport": (("author", "institution"), ("title",), ("institution", "school"), ("year", "date")),
    "manual": (("title",),),
    "misc": (("title",), ("year", "date")),
    "online": (("title",), ("url",), ("year", "date")),
    "unpublished": (("author",), ("title",), ("note",), ("year", "date")),
}

_RE_CITE = re.compile(
    r"\\(?:cite|citep|citet|autocite|parencite|textcite|footcite)\*?(?:\[[^\]]*\])*\{([^}]*)\}"
    r"|\[@([A-Za-z0-9_:.\-]+)",
)
_RE_CHAVE_LITERAL = re.compile(r"\b([A-Za-z][A-Za-z0-9_:\-]*\d{4}[a-z]?)\b")


@dataclass(frozen=True)
class Achado:
    severidade: str
    categoria: str
    chaves: tuple[str, ...]
    mensagem: str


@dataclass
class Relatorio:
    entradas: int = 0
    achados: list[Achado] = field(default_factory=list)
    notas: list[str] = field(default_factory=list)
    checagens_nao_executadas: list[str] = field(default_factory=list)

    @property
    def erros(self) -> list[Achado]:
        return [a for a in self.achados if a.severidade == ERRO]

    @property
    def avisos(self) -> list[Achado]:
        return [a for a in self.achados if a.severidade == AVISO]


def _chave_citada_no_texto(texto: str) -> set[str]:
    """Chaves citadas em .tex/.md (`\\cite{a,b}`, `[@a]`) ou literais `autor2020`."""
    chaves: set[str] = set()
    for grupo, pandoc in _RE_CITE.findall(texto):
        if pandoc:
            chaves.add(pandoc.strip())
            continue
        for parte in grupo.split(","):
            chave_limpa = parte.split(",")[0].strip()
            if chave_limpa:
                chaves.add(chave_limpa)
    if not chaves:
        chaves.update(_RE_CHAVE_LITERAL.findall(texto))
    return {c for c in chaves if c}


def auditar(
    entradas: list[Entrada],
    *,
    problemas: list[ProblemaSintaxe] | None = None,
    texto: str | None = None,
    ano_min: int = ANO_MIN_PADRAO,
    ano_max: int = ANO_MAX_PADRAO,
    estrito: bool = False,
) -> Relatorio:
    relatorio = Relatorio(entradas=len(entradas))

    for problema in problemas or []:
        relatorio.achados.append(
            Achado(ERRO, "sintaxe", (), f"linha {problema.linha}: {problema.mensagem}")
        )

    # chave duplicada
    por_chave: dict[str, list[Entrada]] = defaultdict(list)
    for entrada in entradas:
        por_chave[entrada.chave].append(entrada)
    for chave_citacao, repetidas in sorted(por_chave.items()):
        if len(repetidas) > 1:
            linhas = ", ".join(str(e.linha) for e in repetidas)
            relatorio.achados.append(
                Achado(ERRO, "chave_duplicada", (chave_citacao,), f"{len(repetidas)}× (linhas {linhas})")
            )

    # DOI duplicado
    por_doi: dict[str, list[Entrada]] = defaultdict(list)
    for entrada in entradas:
        doi = entrada.doi()
        if doi:
            por_doi[doi].append(entrada)
    for doi, repetidas in sorted(por_doi.items()):
        if len(repetidas) > 1:
            relatorio.achados.append(
                Achado(
                    ERRO,
                    "doi_duplicado",
                    tuple(e.chave for e in repetidas),
                    f"mesmo DOI {doi}" + (" (títulos diferentes)" if len({chave_titulo(e.campo('title')) for e in repetidas}) > 1 else ""),
                )
            )

    # título normalizado duplicado (quando não já pego pelo DOI)
    por_titulo: dict[str, list[Entrada]] = defaultdict(list)
    for entrada in entradas:
        titulo = chave_titulo(entrada.campo("title"))
        if len(titulo) >= 8:
            por_titulo[titulo].append(entrada)
    for _, repetidas in sorted(por_titulo.items()):
        if len(repetidas) > 1 and len({e.doi() for e in repetidas}) > 1:
            relatorio.achados.append(
                Achado(
                    ERRO,
                    "titulo_duplicado",
                    tuple(e.chave for e in repetidas),
                    f"mesmo título normalizado: {achatar(repetidas[0].campo('title'), 70)}",
                )
            )

    # campos obrigatórios por tipo, ano e identificador
    for entrada in entradas:
        regra = REQUERIDOS.get(entrada.tipo)
        if regra is None:
            relatorio.notas.append(
                f"@{entrada.tipo} ({entrada.chave}): tipo sem regra de campos obrigatórios — não verificado"
            )
        else:
            faltando = [
                "|".join(slot)
                for slot in regra
                if not any(entrada.campo(nome) for nome in slot)
            ]
            if faltando:
                relatorio.achados.append(
                    Achado(
                        ERRO,
                        "campo_obrigatorio",
                        (entrada.chave,),
                        f"@{entrada.tipo} sem {', '.join(faltando)}",
                    )
                )

        ano = entrada.ano()
        if ano is None:
            relatorio.achados.append(
                Achado(ERRO, "ano", (entrada.chave,), f"ano ausente ou ilegível ({entrada.campo('year', 'date')!r})")
            )
        elif not (ano_min <= ano <= ano_max):
            relatorio.achados.append(
                Achado(ERRO, "ano", (entrada.chave,), f"ano {ano} fora da faixa {ano_min}–{ano_max}")
            )

        if not entrada.doi() and not entrada.campo("url"):
            relatorio.achados.append(
                Achado(
                    AVISO if not estrito else ERRO,
                    "sem_identificador",
                    (entrada.chave,),
                    "sem doi nem url (obra antiga pode ser legítimo; conferir à mão)",
                )
            )

    # cruzamento com o texto
    if texto is None:
        relatorio.checagens_nao_executadas.extend(["chaves_nao_citadas", "citacoes_sem_entrada"])
        relatorio.notas.append("sem --texto/--docx: cruzamento citação × bibliografia não executado")
    else:
        chaves_texto = _chave_citada_no_texto(texto)
        chaves_bib = {e.chave for e in entradas}
        modo_chave = bool(_RE_CITE.search(texto))
        if modo_chave:
            for chave_citacao in sorted(chaves_texto - chaves_bib):
                relatorio.achados.append(
                    Achado(ERRO, "citacao_sem_entrada", (chave_citacao,), "citada no texto, ausente no .bib")
                )
            for chave_bib in sorted(chaves_bib - chaves_texto):
                relatorio.achados.append(
                    Achado(
                        ERRO if estrito else AVISO,
                        "chave_nao_citada",
                        (chave_bib,),
                        "entrada nunca citada no texto",
                    )
                )
            relatorio.notas.append(f"cruzamento por chave de citação ({len(chaves_texto)} chaves no texto)")
        else:
            # Modo autor-ano: só dá para afirmar que a entrada aparece ou não.
            corpo = chave(texto)
            for entrada in entradas:
                sobrenome = _sobrenome(entrada)
                ano = entrada.ano()
                if sobrenome and ano and f"{sobrenome} {ano}" in corpo:
                    continue
                relatorio.achados.append(
                    Achado(
                        AVISO,
                        "entrada_nao_encontrada_no_texto",
                        (entrada.chave,),
                        f"'{sobrenome or '?'}, {ano or '?'}' não aparece no texto (busca literal)",
                    )
                )
            relatorio.notas.append(
                "sem chaves de citação no texto: cruzamento em modo autor-ano (heurístico e só AVISO); "
                "'citação sem entrada' não é verificável nesse modo"
            )

    ordem = {ERRO: 0, AVISO: 1}
    relatorio.achados.sort(key=lambda a: (ordem.get(a.severidade, 9), a.categoria, a.chaves))
    return relatorio


def _sobrenome(entrada: Entrada) -> str:
    autor = entrada.campo("author", "editor").split(" and ")[0]
    if "," in autor:
        autor = autor.split(",", 1)[0]
    return chave(autor.split()[-1] if autor.split() else "")


def auditar_arquivo(
    caminho: str | Path,
    *,
    texto: str | None = None,
    ano_min: int = ANO_MIN_PADRAO,
    ano_max: int = ANO_MAX_PADRAO,
    estrito: bool = False,
) -> Relatorio:
    entradas, problemas = ler(caminho)
    return auditar(
        entradas,
        problemas=problemas,
        texto=texto,
        ano_min=ano_min,
        ano_max=ano_max,
        estrito=estrito,
    )


def codigo_saida(relatorio: Relatorio) -> int:
    return 1 if relatorio.erros else 0
