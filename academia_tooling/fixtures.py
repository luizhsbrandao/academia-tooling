"""Constrói as fixtures .docx sintéticas usadas pelos testes.

Nenhum arquivo de vault é lido aqui: o texto é inventado neste módulo. Os .docx
saem em `tests/fixtures/gerados/` (git-ignored) — o binário não entra no git, o
gerador sim, então qualquer checkout reproduz as fixtures com `make fixtures`.
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt

GERADOS_PADRAO = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "gerados"

CONFORME = {
    "titulo": "Guerra cognitiva e a construção do inimigo interno",
    "resumo": (
        "O artigo examina como a noção de guerra cognitiva foi apropriada por "
        "atores estatais brasileiros entre 2018 e 2022. A partir de um corpus de "
        "documentos oficiais e de material de campanha, argumenta-se que a "
        "categoria de inimigo interno foi recomposta por operações de informação "
        "que deslocaram o vocabulário securitário para o cotidiano digital. "
        "O percurso analítico combina revisão bibliográfica e análise de conteúdo."
    ),
    "palavras_chave": "Palavras-chave: guerra cognitiva; securitização; inimigo interno",
    "abstract": (
        "This article examines how the notion of cognitive warfare was appropriated "
        "by Brazilian state actors between 2018 and 2022. Drawing on official "
        "documents and campaign material, it argues that the internal enemy "
        "category was recomposed by information operations that displaced the "
        "security vocabulary into everyday digital life."
    ),
    "corpo": [
        "Introdução",
        (
            "A literatura sobre securitização descreve como elites políticas convertem "
            "temas comuns em questões existenciais. Segundo FOUCAULT (1979), a disciplina "
            "produz corpos dóceis e organiza o que pode ser dito. A leitura simbólica do "
            "Estado foi sistematizada por BOURDIEU (1989), cuja noção de campo explica a "
            "disputa pela legitimidade do discurso securitário no espaço público brasileiro."
        ),
        (
            "Nas duas últimas décadas, a difusão de plataformas digitais alterou o custo de "
            "produção e de circulação de narrativas. O deslocamento do vocabulário "
            "securitário para o cotidiano não elimina a centralidade do Estado, mas a "
            "redistribui (FOUCAULT, 1979). O resultado é um inimigo interno difuso, cuja "
            "identificação depende de operações de informação contínuas."
        ),
    ],
    "referencias": [
        "FOUCAULT, Michel. Vigiar e punir: nascimento da prisão. Petrópolis: Vozes, 1979.",
        "BOURDIEU, Pierre. O poder simbólico. Rio de Janeiro: Bertrand Brasil, 1989.",
    ],
}

FORA_DE_NORMA = {
    "titulo": "Notas sobre propaganda computacional",
    "corpo": [
        (
            "A propaganda computacional combina automação, contas falsas e curadoria "
            "de narrativas. O fenômeno não é novo, mas sua escala mudou o modo como "
            "campanhas eleitorais disputam atenção (FOUCAULT, 1979). Autores ligados "
            "aos estudos de ciência argumentam que a fronteira entre fato e opinião "
            "se desloca quando a circulação é medida em cliques (CHALMERS, 1993)."
        ),
        "Referências",
        "FOUCAULT, Michel. Vigiar e punir: nascimento da prisão. Petrópolis: Vozes, 1979.",
    ],
}

FRAGMENTO = [
    (
        "Este trecho foi extraído no meio de um capítulo e não traz título, resumo, "
        "abstract nem lista de referências, de modo que o verificador não tem como "
        "decidir se o manuscrito cumpre qualquer exigência estrutural da norma, e "
        "por isso deve responder que a checagem não é verificável."
    ),
    (
        "O parágrafo seguinte continua o raciocínio sem nenhuma marca de seção e sem "
        "citação no formato autor-data, apenas prosa corrida sobre fontes primárias "
        "consultadas em arquivo."
    ),
]


def _documento(
    *,
    fonte: str | None = None,
    tamanho: float | None = None,
    espaco: float | None = None,
    margens_cm: tuple[float, float, float, float] | None = None,
    autor: str = "",
    editor: str = "",
) -> Document:
    doc = Document()
    normal = doc.styles["Normal"]
    if fonte:
        normal.font.name = fonte
    if tamanho:
        normal.font.size = Pt(tamanho)
    if espaco:
        normal.paragraph_format.line_spacing = espaco
    if margens_cm:
        superior, inferior, esquerda, direita = margens_cm
        secao = doc.sections[0]
        secao.top_margin = Cm(superior)
        secao.bottom_margin = Cm(inferior)
        secao.left_margin = Cm(esquerda)
        secao.right_margin = Cm(direita)
    doc.core_properties.author = autor
    doc.core_properties.last_modified_by = editor
    return doc


def _apagar_declaracoes(doc: Document) -> None:
    """Apaga fonte/tamanho/espaçamento do docDefaults (fixture 'fragmento').

    Serve para provar o caminho NAO_VERIFICAVEL: o .docx não declara nada e
    python-docx não resolve o tema.
    """
    elemento = doc.styles.element
    for espaco in elemento.xpath("./w:docDefaults/w:pPrDefault/w:pPr/w:spacing"):
        espaco.getparent().remove(espaco)
    for rpr in elemento.xpath("./w:docDefaults/w:rPrDefault/w:rPr"):
        for filho in list(rpr):
            if filho.tag.rsplit("}", 1)[-1] in {"sz", "szCs", "rFonts"}:
                rpr.remove(filho)


def _salvar(doc: Document, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(destino))
    return destino


def _escrever_conforme(doc: Document) -> None:
    """Conteúdo do artigo de teste (título, resumo, corpo, referências casando)."""
    doc.add_paragraph(CONFORME["titulo"])
    doc.add_paragraph("Resumo")
    doc.add_paragraph(CONFORME["resumo"])
    doc.add_paragraph(CONFORME["palavras_chave"])
    doc.add_paragraph("Abstract")
    doc.add_paragraph(CONFORME["abstract"])
    for trecho in CONFORME["corpo"]:
        doc.add_paragraph(trecho)
    doc.add_paragraph("Referências")
    for entrada in CONFORME["referencias"]:
        doc.add_paragraph(entrada)


def conforme(destino: Path) -> Path:
    """ABNT conforme: Times 12, 1,5, margens 3/2/3/2, resumo, referências casando."""
    doc = _documento(
        fonte="Times New Roman",
        tamanho=12,
        espaco=1.5,
        margens_cm=(3.0, 2.0, 3.0, 2.0),
        autor="",
        editor="",
    )
    _escrever_conforme(doc)
    return _salvar(doc, destino)


def direto_10pt(destino: Path) -> Path:
    """Estilo `Normal` 12 pt e TODO o texto em 10 pt Arial por formatação direta.

    É o caso que reprovava o verificador (t_7026bcd5): o laudo lia "12 pt (estilo
    Normal)" e dava PASS enquanto o Word mostrava 10 pt em todos os runs. A
    formatação direta vence o estilo — o veredito tem de ser FAIL.
    """
    doc = _documento(
        fonte="Times New Roman",
        tamanho=12,
        espaco=1.5,
        margens_cm=(3.0, 2.0, 3.0, 2.0),
    )
    _escrever_conforme(doc)
    for paragrafo in doc.paragraphs:
        for run in paragrafo.runs:
            run.font.name = "Arial"
            run.font.size = Pt(10)
    return _salvar(doc, destino)


def espaco_exato(destino: Path) -> Path:
    """Estilo `Normal` 1,5 mas todos os parágrafos com espaçamento exato de 18 pt.

    `w:lineRule="exact"` não é múltiplo comparável com 1,5 linha: a checagem
    `espacamento` precisa responder NAO_VERIFICAVEL (motivo documento), nunca PASS.
    """
    doc = _documento(
        fonte="Times New Roman",
        tamanho=12,
        espaco=1.5,
        margens_cm=(3.0, 2.0, 3.0, 2.0),
    )
    _escrever_conforme(doc)
    for paragrafo in doc.paragraphs:
        paragrafo.paragraph_format.line_spacing = Pt(18)
    return _salvar(doc, destino)


def fora_de_norma(destino: Path) -> Path:
    """Fora de norma: Arial 10, espaço simples, margem 1 cm, sem resumo, autor exposto,
    citação sem entrada na bibliografia."""
    doc = _documento(
        fonte="Arial",
        tamanho=10,
        espaco=1.0,
        margens_cm=(1.0, 1.0, 1.0, 1.0),
        autor="Fulano de Tal",
        editor="Beltrano de Tal",
    )
    doc.add_paragraph(FORA_DE_NORMA["titulo"])
    for trecho in FORA_DE_NORMA["corpo"]:
        doc.add_paragraph(trecho)
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(destino))
    return destino


def fragmento(destino: Path) -> Path:
    """Fragmento: sem estrutura, sem citação, sem declaração de fonte/espaçamento."""
    doc = _documento(autor="", editor="")
    _apagar_declaracoes(doc)
    for trecho in FRAGMENTO:
        doc.add_paragraph(trecho)
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(destino))
    return destino


def montar(
    destino: str | Path,
    linhas: list[str],
    *,
    fonte: str | None = "Times New Roman",
    tamanho: float | None = 12,
    espaco: float | None = 1.5,
    margens_cm: tuple[float, float, float, float] = (3.0, 2.0, 3.0, 2.0),
    autor: str = "",
    editor: str = "",
) -> Path:
    """Grava um .docx arbitrário com o estilo pedido (testes pontuais)."""
    doc = _documento(
        fonte=fonte,
        tamanho=tamanho,
        espaco=espaco,
        margens_cm=margens_cm,
        autor=autor,
        editor=editor,
    )
    for linha in linhas:
        doc.add_paragraph(linha)
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(destino))
    return destino


def construir(destino: str | Path = GERADOS_PADRAO) -> dict[str, Path]:
    destino = Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    return {
        "conforme": conforme(destino / "conforme.docx"),
        "fora_de_norma": fora_de_norma(destino / "fora_de_norma.docx"),
        "fragmento": fragmento(destino / "fragmento.docx"),
        "direto_10pt": direto_10pt(destino / "direto_10pt.docx"),
        "espaco_exato": espaco_exato(destino / "espaco_exato.docx"),
    }


if __name__ == "__main__":  # pragma: no cover
    alvo = Path(sys.argv[1]) if len(sys.argv) > 1 else GERADOS_PADRAO
    for nome, caminho in construir(alvo).items():
        print(f"{nome}: {caminho}")
