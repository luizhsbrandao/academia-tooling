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
from academia_tooling.laudo import FAIL, MOTIVO_DOCUMENTO, NAO_VERIFICAVEL, PASS, Resultado

MARGENS_ABNT = {"superior": 3.0, "inferior": 2.0, "esquerda": 3.0, "direita": 2.0}
SECOES_ABNT = ["titulo", "resumo", "referencias"]

# (id, função, limites da norma, fixture, estrito, status esperado)
CASOS: list[tuple[str, object, dict, str, bool, str]] = [
    # ---- família da fonte
    ("fonte_familia", checar_fonte_familia, {"familia_fonte": ["Times New Roman"]}, "conforme", False, PASS),
    ("fonte_familia", checar_fonte_familia, {"familia_fonte": ["Arial"]}, "conforme", False, FAIL),
    ("fonte_familia", checar_fonte_familia, {"familia_fonte": None}, "conforme", False, NAO_VERIFICAVEL),
    # estilo Normal em Times, corpo todo em Arial por formatação direta
    ("fonte_familia", checar_fonte_familia, {"familia_fonte": ["Times New Roman"]}, "direto_10pt", False, FAIL),
    # ---- tamanho da fonte
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": 12}, "conforme", False, PASS),
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": 12}, "fora_de_norma", False, FAIL),
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": None}, "conforme", False, NAO_VERIFICAVEL),
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": 12}, "fragmento", False, NAO_VERIFICAVEL),
    # Normal 12 pt, corpo todo em 10 pt: PASS aqui seria "PASS por estatística"
    ("fonte_tamanho", checar_fonte_tamanho, {"tamanho_fonte_pt": 12}, "direto_10pt", False, FAIL),
    # ---- espaçamento
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "conforme", False, PASS),
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "fora_de_norma", False, FAIL),
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "fragmento", False, NAO_VERIFICAVEL),
    ("espacamento", checar_espacamento, {"espacamento_linhas": None}, "conforme", False, NAO_VERIFICAVEL),
    # Normal 1,5 mas parágrafos com espaçamento exato (18 pt): não é múltiplo comparável
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "espaco_exato", False, NAO_VERIFICAVEL),
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
    # a norma exige referências e o documento não tem nenhuma entrada nem citação
    ("referencias", checar_referencias, {"secoes_obrigatorias": SECOES_ABNT}, "fragmento", False, FAIL),
    ("referencias", checar_referencias, {"secoes_obrigatorias": None}, "conforme", False, NAO_VERIFICAVEL),
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
    assert resultado.motivo_nao_verificavel == MOTIVO_DOCUMENTO
    assert "nota de rodapé" in resultado.mensagem


# ------------------------------------------------- t_7026bcd5: PASS por estatística


def test_fonte_tamanho_reprova_formato_direto_em_todo_o_corpo(docx, norma_factory):
    """'PASS 12 pt (estilo Normal); 8/8 runs com override' não pode existir.

    O Word aplica formatação direta sobre o estilo: o texto está em 10 pt.
    """
    from docx import Document

    norma = norma_factory(tamanho_fonte_pt=12)
    resultado = checar_fonte_tamanho(Document(str(docx["direto_10pt"])), norma)
    assert resultado.status == FAIL, resultado.mensagem
    total = resultado.detalhes["runs"]
    assert total > 0
    assert resultado.detalhes["runs_fora_do_padrao"] == total
    assert f"{total}/{total} runs fora do padrão" in resultado.mensagem
    assert "10 pt" in resultado.mensagem


def test_fonte_familia_reprova_formato_direto_em_todo_o_corpo(docx, norma_factory):
    from docx import Document

    norma = norma_factory(familia_fonte=["Times New Roman"])
    resultado = checar_fonte_familia(Document(str(docx["direto_10pt"])), norma)
    assert resultado.status == FAIL, resultado.mensagem
    assert "Arial" in resultado.mensagem
    assert resultado.detalhes["runs_fora_do_padrao"] == resultado.detalhes["runs"]


def test_fonte_tamanho_tolera_desvio_pontual_abaixo_do_limiar(tmp_path, norma_factory):
    """1 run em 20 diverge (~5% < LIMIAR_DESVIO): o padrão ainda descreve o corpo."""
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(12)
    for indice in range(10):
        paragrafo = doc.add_paragraph(f"Parágrafo {indice} do corpo em duas partes")
        paragrafo.add_run(" e a continuação do mesmo parágrafo")
    doc.paragraphs[0].runs[0].font.size = Pt(10)  # um único desvio
    caminho = tmp_path / "um_desvio.docx"
    doc.save(str(caminho))

    resultado = checar_fonte_tamanho(Document(str(caminho)), norma_factory(tamanho_fonte_pt=12))
    assert resultado.detalhes["runs"] == 20
    assert resultado.detalhes["runs_fora_do_padrao"] == 1
    assert resultado.status == PASS, resultado.mensagem
    assert "1/20 runs fora do padrão" in resultado.mensagem


def test_espacamento_com_espaco_exato_nao_vira_pass(docx, norma_factory):
    """`_multiplo` → None (Length) não pode cair no padrão do estilo e dar PASS."""
    from docx import Document

    norma = norma_factory(espacamento_linhas=1.5)
    resultado = checar_espacamento(Document(str(docx["espaco_exato"])), norma)
    assert resultado.status == NAO_VERIFICAVEL, resultado.mensagem
    assert resultado.motivo_nao_verificavel == MOTIVO_DOCUMENTO
    assert "espaçamento exato" in resultado.mensagem
    assert resultado.detalhes["paragrafos_com_espaco_fixo"] > 0


def test_espacamento_nao_aprova_por_moda_quando_o_desvio_e_material(tmp_path, norma_factory):
    """"Na maioria" não é conformidade: 2 de 10 parágrafos em 1,0 reprovam."""
    from docx import Document

    doc = Document()
    doc.styles["Normal"].paragraph_format.line_spacing = 1.5
    for indice in range(10):
        paragrafo = doc.add_paragraph(f"Parágrafo {indice} do corpo")
        if indice >= 8:
            paragrafo.paragraph_format.line_spacing = 1.0
    caminho = tmp_path / "maioria_ok.docx"
    doc.save(str(caminho))

    resultado = checar_espacamento(Document(str(caminho)), norma_factory(espacamento_linhas=1.5))
    assert resultado.detalhes["moda"] == 1.5  # é exatamente o caso que dava PASS
    assert resultado.status == FAIL, resultado.mensagem
    assert "2/10" in resultado.mensagem


def test_referencias_sem_entrada_nem_citacao_reprova(tmp_path, norma_factory):
    """A norma exige referências: ausência total do elemento é FAIL, não "provável fragmento"."""
    from docx import Document

    caminho = fixtures.montar(
        tmp_path / "sem_referencias.docx",
        [
            "Título do manuscrito",
            "Resumo",
            "Resumo curto.",
            "Corpo em prosa corrida, sem citação no formato autor-data.",
        ],
    )
    resultado = checar_referencias(Document(str(caminho)), norma_factory(secoes_obrigatorias=SECOES_ABNT))
    assert resultado.status == FAIL, resultado.mensagem
    assert resultado.detalhes["provavel_fragmento"] is True
    assert resultado.detalhes["entradas"] == 0 and resultado.detalhes["citacoes"] == 0
