"""Checagens do verificador .docx: PASS/FAIL/NAO_VERIFICAVEL de cada uma.

A tabela CASOS é o contrato: toda checagem em `CHECAGENS` precisa aparecer nos
três estados, e o caso FAIL de cada uma vem de um .docx real (não de mock). O
teste `test_toda_checagem_tem_os_tres_estados` reprova se alguém adicionar uma
checagem nova sem caso negativo.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from academia_tooling import fixtures
from academia_tooling.docx_check import (
    CHECAGENS,
    auditar,
    checar_autor,
    checar_espacamento,
    checar_fonte_familia,
    checar_fonte_tamanho,
    checar_margens,
    checar_palavras,
    checar_referencias,
    checar_secoes,
    extrair_citacoes,
)
from academia_tooling.laudo import FAIL, NAO_VERIFICAVEL, PASS, Resultado

MARGENS_ABNT = {"superior": 3.0, "inferior": 2.0, "esquerda": 3.0, "direita": 2.0}
SECOES_ABNT = ["titulo", "resumo", "referencias"]

# (id, função, limites da norma, fixture, estrito, status esperado)
CASOS: list[tuple[str, object, dict, str, bool, str]] = [
    # ---- família da fonte
    ("fonte_familia", checar_fonte_familia, {"familia_fonte": ["Times New Roman"]}, "conforme", False, PASS),
    ("fonte_familia", checar_fonte_familia, {"familia_fonte": ["Arial"]}, "conforme", False, FAIL),
    ("fonte_familia", checar_fonte_familia, {"familia_fonte": None}, "conforme", False, NAO_VERIFICAVEL),
    # ---- tamanho da fonte
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": 12}, "conforme", False, PASS),
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": 12}, "fora_de_norma", False, FAIL),
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": None}, "conforme", False, NAO_VERIFICAVEL),
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": 12}, "fragmento", False, NAO_VERIFICAVEL),
    # ---- espaçamento
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "conforme", False, PASS),
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "fora_de_norma", False, FAIL),
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "fragmento", False, NAO_VERIFICAVEL),
    ("espacamento", checar_espacamento, {"espacamento_linhas": None}, "conforme", False, NAO_VERIFICAVEL),
    # ---- margens
    ("margens", checar_margens, {"margens_cm": MARGENS_ABNT}, "conforme", False, PASS),
    ("margens", checar_margens, {"margens_cm": MARGENS_ABNT}, "fora_de_norma", False, FAIL),
    ("margens", checar_margens, {"margens_cm": None}, "conforme", False, NAO_VERIFICAVEL),
    # ---- seções obrigatórias
    ("secoes", checar_secoes, {"secoes_obrigatorias": SECOES_ABNT}, "conforme", False, PASS),
    ("secoes", checar_secoes, {"secoes_obrigatorias": SECOES_ABNT}, "fora_de_norma", False, FAIL),
    ("secoes", checar_secoes, {"secoes_obrigatorias": SECOES_ABNT}, "fragmento", False, NAO_VERIFICAVEL),
    ("secoes", checar_secoes, {"secoes_obrigatorias": None}, "conforme", False, NAO_VERIFICAVEL),
    # ---- contagem de palavras
    (
        "palavras",
        checar_palavras,
        {"palavras": {"min": 100, "max": 1000}, "contagem_exclui_referencias": False},
        "conforme",
        False,
        PASS,
    ),
    (
        "palavras",
        checar_palavras,
        {"palavras": {"min": 7000, "max": 9000}, "contagem_exclui_referencias": True},
        "conforme",
        False,
        FAIL,
    ),
    ("palavras", checar_palavras, {"palavras": None}, "conforme", False, NAO_VERIFICAVEL),
    # ---- metadado de autor
    ("autor_metadado", checar_autor, {"revisao_cega": True}, "conforme", False, PASS),
    ("autor_metadado", checar_autor, {"revisao_cega": True}, "fora_de_norma", False, FAIL),
    ("autor_metadado", checar_autor, {"revisao_cega": None}, "conforme", False, NAO_VERIFICAVEL),
    # ---- referências × citações
    ("referencias", checar_referencias, {"secoes_obrigatorias": SECOES_ABNT}, "conforme", False, PASS),
    ("referencias", checar_referencias, {"secoes_obrigatorias": SECOES_ABNT}, "fora_de_norma", False, FAIL),
    ("referencias", checar_referencias, {"secoes_obrigatorias": SECOES_ABNT}, "fragmento", False, NAO_VERIFICAVEL),
]


@pytest.mark.parametrize(
    "checagem,funcao,limites,fixture,estrito,esperado",
    CASOS,
    ids=[f"{c[0]}-{c[5]}-{c[3]}" for c in CASOS],
)
def test_estados_da_checagem(checagem, funcao, limites, fixture, estrito, esperado, docx, norma_factory, documento):
    norma = norma_factory(**limites)
    resultado = funcao(documento(docx[fixture]), norma, estrito)
    assert isinstance(resultado, Resultado)
    assert resultado.status == esperado, f"{checagem}/{fixture}: {resultado.status} ≠ {esperado} — {resultado.mensagem}"
    assert resultado.mensagem, "toda checagem precisa explicar o veredito"


def test_toda_checagem_tem_os_tres_estados():
    """Sem caso negativo, a checagem não é checagem."""
    for checagem, _, _ in CHECAGENS:
        estados = {c[5] for c in CASOS if c[0] == checagem}
        assert estados == {PASS, FAIL, NAO_VERIFICAVEL}, f"{checagem} só cobre {sorted(estados)}"


def test_auditar_devolve_uma_linha_por_checagem(docx, norma_factory):
    norma = norma_factory(
        margens_cm=MARGENS_ABNT,
        secoes_obrigatorias=SECOES_ABNT,
        tamanho_fonte_pt=12,
        espacamento_linhas=1.5,
    )
    resultados = auditar(docx["conforme"], norma)
    assert [r.id for r in resultados] == [c[0] for c in CHECAGENS]
    assert len({r.id for r in resultados}) == len(resultados)


def test_laudo_conforme_nao_reprova(docx, norma_factory):
    norma = norma_factory(
        margens_cm=MARGENS_ABNT,
        secoes_obrigatorias=SECOES_ABNT,
        tamanho_fonte_pt=12,
        espacamento_linhas=1.5,
        familia_fonte=["Times New Roman"],
        revisao_cega=True,
        palavras={"min": 100, "max": 1000},
        contagem_exclui_referencias=True,
    )
    resultados = auditar(docx["conforme"], norma)
    assert [r.status for r in resultados] == [PASS] * len(resultados), [
        (r.id, r.status, r.mensagem) for r in resultados if r.status != PASS
    ]


def test_palavras_por_secao_e_exclusao_das_referencias(docx, norma_factory):
    norma = norma_factory(palavras={"min": 1, "max": 10000}, contagem_exclui_referencias=True)
    from docx import Document

    resultado = checar_palavras(Document(str(docx["conforme"])), norma)
    por_secao = resultado.detalhes["por_secao"]
    assert any(nome.startswith("resumo") for nome in por_secao), por_secao
    assert any(nome.startswith("referencias") for nome in por_secao), por_secao
    assert resultado.detalhes["contado"] < resultado.detalhes["total"]
    assert resultado.detalhes["referencias"] > 0


def test_extrair_citacoes_reconhece_autor_data_e_ignora_ano_solto():
    texto = (
        "Segundo FOUCAULT (1979), a disciplina produz corpos. "
        "O campo foi descrito em (BOURDIEU, 1989) e em (SILVA; SOUZA, 2020). "
        "A coleta ocorreu em 2019, sem citação."
    )
    assert extrair_citacoes(texto) == {
        ("foucault", 1979),
        ("bourdieu", 1989),
        ("silva", 2020),
        ("souza", 2020),
    }


def test_referencias_estrito_reprova_entrada_nao_citada(tmp_path, norma_factory):
    caminho = fixtures.montar(
        tmp_path / "uma_citacao_duas_entradas.docx",
        [
            "Título do manuscrito",
            "Resumo",
            "Resumo curto do artigo.",
            "Corpo com citação (SILVA, 2020) apenas.",
            "Referências",
            "SILVA, João. Um artigo. Revista, 2020.",
            "SOUZA, Maria. Outro artigo. Revista, 2021.",
        ],
    )
    norma = norma_factory(secoes_obrigatorias=SECOES_ABNT)
    from docx import Document

    brando = checar_referencias(Document(str(caminho)), norma)
    assert brando.status == PASS
    assert brando.detalhes["entradas_nao_citadas"] == ["souza"]

    estrito = checar_referencias(Document(str(caminho)), norma, estrito=True)
    assert estrito.status == FAIL
    assert "não citada" in estrito.mensagem


def test_ano_divergente_entre_citacao_e_entrada(tmp_path, norma_factory):
    caminho = fixtures.montar(
        tmp_path / "ano_divergente.docx",
        [
            "Título do manuscrito",
            "Resumo",
            "Resumo curto.",
            "Corpo cita (SILVA, 2021).",
            "Referências",
            "SILVA, João. Um artigo. Revista, 2020.",
        ],
    )
    norma = norma_factory(secoes_obrigatorias=SECOES_ABNT)
    from docx import Document

    resultado = checar_referencias(Document(str(caminho)), norma)
    assert resultado.status == FAIL
    assert resultado.detalhes["anos_divergentes"] == ["silva, 2021"]


def test_manuscrito_em_nota_de_rodape_nao_e_reprovado(tmp_path, norma_factory):
    """História usa nota de rodapé: sem citação autor-data o veredito é NAO_VERIFICAVEL."""
    caminho = fixtures.montar(
        tmp_path / "notas.docx",
        [
            "Título do manuscrito",
            "Resumo",
            "Resumo curto.",
            "Corpo sem citação autor-data, apenas notas.",
            "Referências",
            "SILVA, João. Um artigo. Revista, 2020.",
        ],
    )
    norma = norma_factory(secoes_obrigatorias=SECOES_ABNT)
    from docx import Document

    resultado = checar_referencias(Document(str(caminho)), norma)
    assert resultado.status == NAO_VERIFICAVEL
    assert "nota de rodapé" in resultado.mensagem
