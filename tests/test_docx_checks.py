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
    extrair_chamadas,
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
    # nada declarado em nível nenhum: o efetivo é o default da aplicação (1,0),
    # que não atende a 1,5 — reprova pelo default, não fica "não verificável"
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "fragmento", False, FAIL),
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "sem_setup", False, FAIL),
    ("espacamento", checar_espacamento, {"espacamento_linhas": None}, "conforme", False, NAO_VERIFICAVEL),
    # Normal 1,5 mas parágrafos com espaçamento exato (18 pt): não é múltiplo comparável
    ("espacamento", checar_espacamento, {"espacamento_linhas": 1.5}, "espaco_exato", False, NAO_VERIFICAVEL),
    # ---- margens
    ("margens", checar_margens, {"margens_cm": MARGENS_ABNT}, "conforme", False, PASS),
    ("margens", checar_margens, {"margens_cm": MARGENS_ABNT}, "fora_de_norma", False, FAIL),
    ("margens", checar_margens, {"margens_cm": None}, "conforme", False, NAO_VERIFICAVEL),
    # .docx sem w:sectPr: não declara margem alguma → default de 2,54 cm reprova
    ("margens", checar_margens, {"margens_cm": MARGENS_ABNT}, "sem_setup", False, FAIL),
    # ---- seções obrigatórias
    ("secoes", checar_secoes, {"secoes_obrigatorias": SECOES_ABNT}, "conforme", False, PASS),
    ("secoes", checar_secoes, {"secoes_obrigatorias": SECOES_ABNT}, "fora_de_norma", False, FAIL),
    ("secoes", checar_secoes, {"secoes_obrigatorias": SECOES_ABNT}, "fragmento", False, NAO_VERIFICAVEL),
    # resumo existe como rótulo em negrito, sem título de seção: forma reprova,
    # existência não pode ser negada
    ("secoes", checar_secoes, {"secoes_obrigatorias": SECOES_ABNT}, "resumo_rotulo", False, FAIL),
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
    # chamada de dois autores (narrativa, parentética e só o segundo): a entrada
    # existe e nenhum dos dois pode virar "não citado"
    (
        "referencias",
        checar_referencias,
        {"secoes_obrigatorias": SECOES_ABNT},
        "citacao_dois_autores",
        False,
        PASS,
    ),
    # ---- referências: os falsos positivos medidos no card t_7c691dbc, um por
    # causa, em fixtures sintéticas. Todos têm de ser PASS.
    # data de acesso ("Acesso em: … 2026") lida como ano de publicação
    (
        "referencias",
        checar_referencias,
        {"secoes_obrigatorias": SECOES_ABNT},
        "referencias_acesso",
        False,
        PASS,
    ),
    # dígitos do DOI ("10.14763/2026.1.2074") lidos como ano
    (
        "referencias",
        checar_referencias,
        {"secoes_obrigatorias": SECOES_ABNT},
        "referencias_doi",
        False,
        PASS,
    ),
    # apud: a fonte não consultada não entra nas referências
    (
        "referencias",
        checar_referencias,
        {"secoes_obrigatorias": SECOES_ABNT},
        "referencias_apud",
        False,
        PASS,
    ),
    # título de obra em prosa ("Em Propaganda (1928)") não é chamada autor-data
    (
        "referencias",
        checar_referencias,
        {"secoes_obrigatorias": SECOES_ABNT},
        "referencias_titulo_prosa",
        False,
        PASS,
    ),
    # duas entradas do mesmo sobrenome: a segunda também é citada
    (
        "referencias",
        checar_referencias,
        {"secoes_obrigatorias": SECOES_ABNT},
        "referencias_autor_repetido",
        False,
        PASS,
    ),
    # trava do verdadeiro positivo: chamada autor-data sem entrada reprova
    (
        "referencias",
        checar_referencias,
        {"secoes_obrigatorias": SECOES_ABNT},
        "referencias_sem_entrada",
        False,
        FAIL,
    ),
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


# --------------------------------- t_7c691dbc: 10 FAIL falsos em manuscrito ABNT
# Cada teste abaixo isola uma causa medida no manuscrito real. As fixtures são
# sintéticas (nada de vault): o que se reproduz é o *padrão* da entrada/chamada.


def _referencias(docx, norma_factory, fixture: str, estrito: bool = False) -> Resultado:
    from docx import Document

    norma = norma_factory(secoes_obrigatorias=SECOES_ABNT)
    return checar_referencias(Document(str(docx[fixture])), norma, estrito)


def test_ano_de_acesso_nao_vira_ano_de_publicacao(docx, norma_factory):
    """F1: em referência online o último ano é o do `Acesso em:`, não o da obra.

    Cinco acusações falsas de "ano divergente" vinham de `anos[-1]`.
    """
    resultado = _referencias(docx, norma_factory, "referencias_acesso")
    assert resultado.status == PASS, resultado.mensagem
    # o ano lido de cada entrada é o de publicação, não o do acesso (2026)
    assert resultado.detalhes["anos_por_entrada"] == [[2022], [2017], [2025], [2018], [2009]]
    assert resultado.detalhes["anos_divergentes"] == []
    assert resultado.detalhes["citacoes_sem_entrada"] == []
    assert resultado.detalhes["entradas_nao_citadas"] == []


def test_digitos_do_doi_nao_viram_ano(docx, norma_factory):
    """F2: `DOI: 10.14763/2026.1.2074` produzia ano 2074 contra a chamada de 2026."""
    resultado = _referencias(docx, norma_factory, "referencias_doi")
    assert resultado.status == PASS, resultado.mensagem
    assert resultado.detalhes["anos_por_entrada"] == [[2026]]
    assert resultado.detalhes["anos_divergentes"] == []


def test_apud_so_exige_a_obra_consultada(docx, norma_factory):
    """F3: NBR 10520 — a fonte não consultada citada por *apud* não é entrada."""
    resultado = _referencias(docx, norma_factory, "referencias_apud")
    assert resultado.status == PASS, resultado.mensagem
    assert resultado.detalhes["citacoes_sem_entrada"] == []
    assert resultado.detalhes["sobrenomes_citados"] == ["boswinkel"]
    assert "rid" not in resultado.detalhes["sobrenomes_na_bibliografia"]


def test_titulo_de_obra_em_prosa_nao_vira_chamada(docx, norma_factory):
    """F4/F6: "Em Propaganda (1928)" é menção a título, não chamada autor-data."""
    resultado = _referencias(docx, norma_factory, "referencias_titulo_prosa")
    assert resultado.status == PASS, resultado.mensagem
    assert resultado.detalhes["citacoes_sem_entrada"] == []
    assert resultado.detalhes["mencionadas_como_titulo"] == [
        "opinion (1922)",
        "propaganda (1928)",
    ]


def test_autor_com_duas_entradas_marca_as_duas(docx, norma_factory):
    """F5: `dono.setdefault` deixava a segunda obra do mesmo autor "não citada"."""
    resultado = _referencias(docx, norma_factory, "referencias_autor_repetido")
    assert resultado.status == PASS, resultado.mensagem
    assert resultado.detalhes["entradas_nao_citadas"] == []
    assert resultado.detalhes["citacoes_sem_entrada"] == []
    assert resultado.detalhes["anos_por_entrada"] == [[2001], [2014], [2017], [2022]]


def test_citacao_sem_entrada_continua_reprovando(docx, norma_factory):
    """Trava: o verdadeiro positivo não pode virar PASS na correção dos falsos.

    `(EGN/PPGEM, 2026)` é chamada autor-data sem entrada (o caso real que sobrava
    depois de descontadas as 10 acusações falsas); `(Klein, 2020)` cobre a
    chamada em caixa mista, que continua tendo de reprovar.
    """
    resultado = _referencias(docx, norma_factory, "referencias_sem_entrada")
    assert resultado.status == FAIL
    assert "egn (2026)" in resultado.detalhes["citacoes_sem_entrada"]
    assert "klein (2020)" in resultado.detalhes["citacoes_sem_entrada"]
    assert "foucault (1979)" not in resultado.detalhes["citacoes_sem_entrada"]


def test_chamada_em_caixa_baixa_continua_sendo_chamada():
    """F4 sem excesso: só a partícula de título (em/no/na) desqualifica a chamada."""
    chamadas = {
        chamada.rotulo(): chamada
        for chamada in extrair_chamadas(
            "Segundo Lippmann (1922), o estereótipo antecede o juízo. "
            "Em Propaganda (1928), Lippmann descrevera a fabricação do consentimento. "
            "A crítica está em (Klein, 2020) e o método em (SILVA; SOUZA, 2024)."
        )
    }
    assert chamadas["lippmann (1922)"].pode_acusar()
    assert chamadas["klein (2020)"].pode_acusar()
    assert chamadas["silva (2024)"].pode_acusar()
    assert chamadas["souza (2024)"].pode_acusar()
    assert not chamadas["propaganda (1928)"].pode_acusar()


def test_apud_extrai_apenas_a_obra_consultada():
    (chamada,) = extrair_chamadas("(RID, 2020, p. 13–14 apud BOSWINKEL et al., 2022, p. 5)")
    assert chamada.nomes == ("boswinkel",)
    assert chamada.ano == 2022


def test_anos_da_entrada_ignora_acesso_doi_e_issn():
    """Ano de entrada é conjunto e exclui cauda de acesso e identificadores."""
    from academia_tooling.docx_check import _anos_da_entrada

    assert _anos_da_entrada(
        "BOSWINKEL, H. et al. Coordinated behavior. Journal, abr. 2022. "
        "Disponível em: https://exemplo.invalid/artigo/2291. Acesso em: 11 maio 2026."
    ) == {2022}
    assert _anos_da_entrada(
        "CHEN, Y. Governing recommendation. Internet Policy Review, 2026. "
        "DOI: 10.14763/2026.1.2074. ISSN 2043-5432."
    ) == {2026}
    # duas datas legítimas (edição e reimpressão): o conjunto preserva as duas
    assert _anos_da_entrada("SILVA, J. Um artigo. Revista, 2019. 2. ed. 2021.") == {2019, 2021}


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


# -------------------- t_516daa8a: FAIL falso (resumo ausente, dois autores, setup)


def test_secoes_nao_diz_ausente_quando_o_resumo_e_rotulo_no_corpo(docx, norma_factory):
    """O resumo existe (rótulo em negrito); o que falta é o título de seção.

    Dizer "ausentes: resumo" manda reescrever o que está certo.
    """
    from docx import Document

    resultado = checar_secoes(
        Document(str(docx["resumo_rotulo"])), norma_factory(secoes_obrigatorias=SECOES_ABNT)
    )
    assert resultado.status == FAIL, resultado.mensagem
    assert resultado.detalhes["faltando"] == [], resultado.detalhes
    assert "resumo" in resultado.detalhes["rotulos_sem_titulo"], resultado.detalhes
    assert "ausente" not in resultado.mensagem, resultado.mensagem
    assert "sem título de seção" in resultado.mensagem
    assert "rótulo" in resultado.mensagem


def test_secoes_continua_dizendo_ausente_quando_o_elemento_nao_existe(docx, norma_factory):
    """A correção não pode esconder ausência real: sem resumo, "ausentes: resumo"."""
    from docx import Document

    resultado = checar_secoes(
        Document(str(docx["fora_de_norma"])), norma_factory(secoes_obrigatorias=SECOES_ABNT)
    )
    assert resultado.status == FAIL
    assert resultado.detalhes["faltando"] == ["resumo"], resultado.detalhes
    assert resultado.detalhes["rotulos_sem_titulo"] == {}
    assert "ausentes: resumo" in resultado.mensagem


def test_secoes_reconhece_rotulo_em_paragrafo_comum_sem_negrito(tmp_path, norma_factory):
    """"Resumo. <texto>" em parágrafo comum, sem negrito: é rótulo, não ausência."""
    from docx import Document

    caminho = fixtures.montar(
        tmp_path / "resumo_sem_negrito.docx",
        [
            "Título do manuscrito em parágrafo comum",
            "Resumo. O artigo examina um objeto qualquer e argumenta que ele importa.",
            "Corpo do artigo cita (SILVA, 2020).",
            "Referências",
            "SILVA, João. Um artigo. Revista, 2020.",
        ],
    )
    resultado = checar_secoes(Document(str(caminho)), norma_factory(secoes_obrigatorias=SECOES_ABNT))
    assert resultado.status == FAIL
    assert resultado.detalhes["faltando"] == []
    assert "resumo" in resultado.detalhes["rotulos_sem_titulo"]


def test_extrair_chamadas_preserva_todos_os_autores_de_cada_chamada():
    from academia_tooling.docx_check import extrair_chamadas

    texto = (
        "Silva e Souza (2024) discutem o tema. Também aparece em (SILVA; SOUZA, 2024), "
        "em (SILVA e SOUZA, 2023), em Silva, Souza e Lima (2023) e em Silva et al. (2022)."
    )
    nomes = [set(chamada.nomes) for chamada in extrair_chamadas(texto)]
    assert {"silva", "souza"} in nomes, nomes  # narrativa de dois autores
    assert {"silva", "souza"} in nomes
    assert {"silva", "souza", "lima"} in nomes, nomes  # três autores na narrativa
    assert {"silva"} in nomes, nomes  # et al. casa pelo primeiro autor


def test_referencias_casa_chamada_de_dois_autores_em_todas_as_formas(docx, norma_factory):
    """Chamada dupla não pode virar "citação sem entrada" + autor "não citado"."""
    from docx import Document

    resultado = checar_referencias(
        Document(str(docx["citacao_dois_autores"])), norma_factory(secoes_obrigatorias=SECOES_ABNT)
    )
    assert resultado.status == PASS, resultado.mensagem
    assert resultado.detalhes["citacoes_sem_entrada"] == [], resultado.detalhes
    assert resultado.detalhes["entradas_nao_citadas"] == [], resultado.detalhes
    assert resultado.detalhes["chamadas"] >= 3
    assert "souza" in resultado.detalhes["sobrenomes_na_bibliografia"]


def test_referencias_chamada_so_do_segundo_autor_encontra_a_entrada(tmp_path, norma_factory):
    from docx import Document

    caminho = fixtures.montar(
        tmp_path / "so_segundo_autor.docx",
        [
            "Título do manuscrito em parágrafo comum",
            "Resumo",
            "Resumo curto.",
            "O argumento é de Souza (2024), que assina o artigo com Silva.",
            "Referências",
            "SILVA, Antônio H.; SOUZA, Nélia A. A categoria elástica. Revista, 2024.",
        ],
    )
    resultado = checar_referencias(Document(str(caminho)), norma_factory(secoes_obrigatorias=SECOES_ABNT))
    assert resultado.status == PASS, resultado.mensagem
    assert resultado.detalhes["citacoes_sem_entrada"] == [], resultado.detalhes
    assert resultado.detalhes["entradas_nao_citadas"] == [], resultado.detalhes


def test_referencias_ainda_reprova_citacao_realmente_sem_entrada(tmp_path, norma_factory):
    """A tolerância a dois autores não pode cegar o cruzamento."""
    from docx import Document

    caminho = fixtures.montar(
        tmp_path / "citacao_orfa.docx",
        [
            "Título do manuscrito em parágrafo comum",
            "Resumo",
            "Resumo curto.",
            "Outro autor sustenta o ponto (LIMA, 2019).",
            "Referências",
            "SILVA, Antônio H.; SOUZA, Nélia A. A categoria elástica. Revista, 2024.",
        ],
    )
    resultado = checar_referencias(Document(str(caminho)), norma_factory(secoes_obrigatorias=SECOES_ABNT))
    assert resultado.status == FAIL
    assert any("lima" in item for item in resultado.detalhes["citacoes_sem_entrada"]), resultado.detalhes
    assert resultado.detalhes["entradas_nao_citadas"] == ["silva"], resultado.detalhes
    assert "sem entrada" in resultado.mensagem


def test_fixture_sem_setup_realmente_nao_declara_pagina(docx):
    """A fixture do D3 reproduz o arquivo medido: nem w:sectPr, nem w:spacing."""
    import zipfile

    with zipfile.ZipFile(docx["sem_setup"]) as pacote:
        xml = pacote.read("word/document.xml").decode("utf-8")
    assert "sectPr" not in xml
    assert "w:spacing" not in xml


def test_margens_sem_setup_de_pagina_diz_qual_e_o_efetivo(docx, norma_factory):
    """Sem pgMar o efetivo é o default da aplicação — não "não declarada" do verificador."""
    from docx import Document

    resultado = checar_margens(
        Document(str(docx["sem_setup"])), norma_factory(margens_cm=MARGENS_ABNT)
    )
    assert resultado.status == FAIL, resultado.mensagem
    assert resultado.detalhes["pagina_sem_setup"] is True
    assert resultado.detalhes["default_aplicacao_cm"] == 2.54
    assert "não declara margem alguma" in resultado.mensagem
    assert "default da aplicação" in resultado.mensagem
    assert "não declarada" not in resultado.mensagem


def test_espacamento_sem_declaracao_reprova_pelo_default_da_aplicacao(docx, norma_factory):
    from docx import Document

    resultado = checar_espacamento(
        Document(str(docx["sem_setup"])), norma_factory(espacamento_linhas=1.5)
    )
    assert resultado.status == FAIL, resultado.mensagem
    assert resultado.detalhes["default_aplicacao"] == 1.0
    assert "não declara espaçamento em nenhum nível" in resultado.mensagem
    assert "default da aplicação" in resultado.mensagem
    assert "não é legível" not in resultado.mensagem


def test_espacamento_sem_declaracao_nao_inventa_pass_quando_o_default_coincide(docx, norma_factory):
    """Default 1,0 e norma pedindo 1,0: nada declarado continua não verificado."""
    from docx import Document

    resultado = checar_espacamento(
        Document(str(docx["sem_setup"])), norma_factory(espacamento_linhas=1.0)
    )
    assert resultado.status == NAO_VERIFICAVEL, resultado.mensagem
    assert resultado.motivo_nao_verificavel == MOTIVO_DOCUMENTO
    assert "nada está declarado" in resultado.mensagem
