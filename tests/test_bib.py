"""Auditor .bib: parser, cada checagem e código de saída."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from academia_tooling.bib_check import AVISO, ERRO, auditar, auditar_arquivo, codigo_saida
from academia_tooling.bib_parse import Entrada, ler, normalizar_doi, parsear

FIX = Path(__file__).resolve().parent / "fixtures"
COM_PROBLEMAS = FIX / "bib_com_problemas.bib"
OK = FIX / "bib_ok.bib"
SEM_DOI = FIX / "bib_titulo_sem_doi.bib"
TEXTO_PROBLEMAS = FIX / "texto_problemas.tex"
TEXTO_OK = FIX / "texto_ok.tex"


def _categorias(relatorio) -> set[str]:
    return {a.categoria for a in relatorio.achados}


# ------------------------------------------------------------------ parser


def test_parser_le_aspas_braces_e_macro_string():
    texto = """
% comentário solto
@string{rev = "Revista X"}
@article{a1,
  author  = {Silva, Jo{\\~a}o},
  title   = "Um {T}\\'itulo",
  journal = rev,
  year    = 2020,
  doi     = {doi:10.1/ABC}
}
"""
    entradas, problemas = parsear(texto)
    assert problemas == []
    assert len(entradas) == 1
    entrada = entradas[0]
    assert entrada.tipo == "article"
    assert entrada.campos["journal"] == "Revista X"
    assert entrada.campos["title"] == "Um {T}\\'itulo"
    assert entrada.ano() == 2020
    assert entrada.doi() == "10.1/abc"


def test_parser_ignora_comment_e_preamble():
    entradas, problemas = parsear(
        '@comment{isto nao e entrada}\n@preamble{"\\newcommand{\\x}{y}"}\n@book{b, title={T}, year=1999}'
    )
    assert [e.chave for e in entradas] == ["b"]
    assert problemas == []


def test_parser_reporta_entrada_nao_fechada():
    entradas, problemas = parsear("@article{falta, title = {T}, year = 2020")
    assert entradas == []
    assert problemas and "não fechada" in problemas[0].mensagem


def test_parser_reporta_campo_sem_igual():
    _, problemas = parsear("@article{x, title {T}, year = 2020}")
    assert any("sem '='" in p.mensagem for p in problemas)


def test_parser_reporta_entrada_sem_chave():
    _, problemas = parsear("@article{, title = {T}}")
    assert any("sem chave" in p.mensagem for p in problemas)


def test_doi_normalizado_em_varias_formas():
    assert normalizar_doi("https://doi.org/10.1590/ABC") == "10.1590/abc"
    assert normalizar_doi("http://dx.doi.org/10.1590/ABC.") == "10.1590/abc"
    assert normalizar_doi("doi: 10.1590/ABC;") == "10.1590/abc"
    assert normalizar_doi("") == ""


def test_entrada_ajudantes():
    entrada = Entrada(tipo="article", chave="k", campos={"year": "2019", "doi": "10.1/X", "url": ""})
    assert entrada.ano() == 2019
    assert entrada.doi() == "10.1/x"
    assert entrada.identificador() == "10.1/x"


# ------------------------------------------------------------------ checagens


def test_bib_limpo_nao_tem_achado():
    relatorio = auditar_arquivo(OK, texto=TEXTO_OK.read_text(encoding="utf-8"))
    assert relatorio.entradas == 3
    assert relatorio.achados == [], [a.mensagem for a in relatorio.achados]
    assert codigo_saida(relatorio) == 0


def test_bib_com_problemas_detecta_cada_categoria():
    relatorio = auditar_arquivo(COM_PROBLEMAS, texto=TEXTO_PROBLEMAS.read_text(encoding="utf-8"))
    categorias = _categorias(relatorio)
    esperadas = {
        "doi_duplicado",
        "titulo_duplicado",
        "campo_obrigatorio",
        "ano",
        "citacao_sem_entrada",
        "chave_nao_citada",
        "sem_identificador",
    }
    assert esperadas <= categorias, sorted(esperadas - categorias)
    assert codigo_saida(relatorio) == 1


def test_duplicata_por_doi_aponta_as_duas_chaves():
    relatorio = auditar_arquivo(COM_PROBLEMAS, texto=TEXTO_PROBLEMAS.read_text(encoding="utf-8"))
    achado = next(a for a in relatorio.achados if a.categoria == "doi_duplicado")
    assert set(achado.chaves) == {"silva2020", "silva2020b"}
    assert "10.1590/123" in achado.mensagem


def test_duplicata_por_titulo_normalizado():
    relatorio = auditar_arquivo(COM_PROBLEMAS, texto=TEXTO_PROBLEMAS.read_text(encoding="utf-8"))
    achado = next(a for a in relatorio.achados if a.categoria == "titulo_duplicado")
    assert set(achado.chaves) == {"silva2020", "souza2021"}


def test_titulo_duplicado_sem_doi_nas_duas_entradas():
    """t_7026bcd5: duas entradas sem DOI montavam `{""}` (tamanho 1) e nunca disparavam."""
    relatorio = auditar_arquivo(SEM_DOI)
    achados = [a for a in relatorio.achados if a.categoria == "titulo_duplicado"]
    assert len(achados) == 1, [a.mensagem for a in relatorio.achados]
    assert achados[0].severidade == ERRO
    assert set(achados[0].chaves) == {"lima2018", "lima2018b"}
    assert codigo_saida(relatorio) == 1


def test_grupo_com_doi_compartilhado_nao_duplica_o_achado():
    """Mesmo título + mesmo DOI já é `doi_duplicado`: não se reporta duas vezes."""
    relatorio = auditar_arquivo(COM_PROBLEMAS)
    grupos_doi = [set(a.chaves) for a in relatorio.achados if a.categoria == "doi_duplicado"]
    grupos_titulo = [set(a.chaves) for a in relatorio.achados if a.categoria == "titulo_duplicado"]
    assert {"silva2020", "silva2020b"} in grupos_doi
    assert {"silva2020", "silva2020b"} not in grupos_titulo


def test_campo_obrigatorio_por_tipo():
    relatorio = auditar_arquivo(OK)
    assert "campo_obrigatorio" not in _categorias(relatorio)

    relatorio = auditar_arquivo(COM_PROBLEMAS, texto=TEXTO_PROBLEMAS.read_text(encoding="utf-8"))
    achado = next(a for a in relatorio.achados if a.categoria == "campo_obrigatorio" and a.chaves == ("livro_sem_editora",))
    assert "publisher" in achado.mensagem


def test_ano_fora_de_faixa_e_configuravel():
    relatorio = auditar_arquivo(OK, ano_min=1980, ano_max=2000)
    achado = next(a for a in relatorio.achados if a.categoria == "ano")
    assert achado.chaves == ("foucault1979",)
    assert "1979" in achado.mensagem


def test_sem_identificador_e_aviso_nao_reprova():
    relatorio = auditar_arquivo(COM_PROBLEMAS)
    achados = [a for a in relatorio.achados if a.categoria == "sem_identificador"]
    assert achados and all(a.severidade == AVISO for a in achados)


def test_entrada_sem_doi_nem_url_aparece_como_aviso(tmp_path):
    caminho = tmp_path / "so_titulo.bib"
    caminho.write_text(
        "@book{sem_id, author={Silva, Joao}, title={Titulo}, publisher={Editora}, year={2001}}\n",
        encoding="utf-8",
    )
    relatorio = auditar_arquivo(caminho)
    achados = [a for a in relatorio.achados if a.categoria == "sem_identificador"]
    assert len(achados) == 1
    assert achados[0].severidade == AVISO
    assert codigo_saida(relatorio) == 0


def test_estrito_promove_aviso_a_erro(tmp_path):
    caminho = tmp_path / "so_titulo.bib"
    caminho.write_text(
        "@book{sem_id, author={Silva, Joao}, title={Titulo}, publisher={Editora}, year={2001}}\n",
        encoding="utf-8",
    )
    relatorio = auditar_arquivo(caminho, estrito=True)
    assert all(a.severidade == ERRO for a in relatorio.achados)
    assert codigo_saida(relatorio) == 1


def test_sem_texto_as_checagens_de_citacao_nao_rodam():
    relatorio = auditar_arquivo(COM_PROBLEMAS)
    assert set(relatorio.checagens_nao_executadas) == {"chaves_nao_citadas", "citacoes_sem_entrada"}
    assert "citacao_sem_entrada" not in _categorias(relatorio)
    assert any("sem --texto/--docx" in nota for nota in relatorio.notas)


def test_citacao_sem_entrada_e_erro():
    relatorio = auditar_arquivo(COM_PROBLEMAS, texto=TEXTO_PROBLEMAS.read_text(encoding="utf-8"))
    achado = next(a for a in relatorio.achados if a.categoria == "citacao_sem_entrada")
    assert achado.chaves == ("fantasma2022",)
    assert achado.severidade == ERRO


def test_chave_nao_citada_lista_o_que_sobrou():
    relatorio = auditar_arquivo(COM_PROBLEMAS, texto=TEXTO_PROBLEMAS.read_text(encoding="utf-8"))
    nao_citadas = {a.chaves[0] for a in relatorio.achados if a.categoria == "chave_nao_citada"}
    assert {"silva2020b", "livro_sem_editora", "nota_solta", "ano_errado", "sem_ano"} <= nao_citadas


def test_modo_autor_ano_quando_nao_ha_chaves_de_citacao():
    relatorio = auditar_arquivo(OK, texto="O artigo dialoga com Foucault 1979 e Bourdieu 1989.")
    categorias = _categorias(relatorio)
    assert "citacao_sem_entrada" not in categorias
    assert any("autor-ano" in nota for nota in relatorio.notas)
    faltando = {a.chaves[0] for a in relatorio.achados if a.categoria == "entrada_nao_encontrada_no_texto"}
    assert faltando == {"chalmers1993"}


def test_relatorio_ordena_erros_antes_de_avisos():
    relatorio = auditar_arquivo(COM_PROBLEMAS, texto=TEXTO_PROBLEMAS.read_text(encoding="utf-8"))
    severidades = [a.severidade for a in relatorio.achados]
    assert severidades == sorted(severidades, key=lambda s: 0 if s == ERRO else 1)


def test_entradas_do_fixture_sao_lidas_inteiras():
    entradas, problemas = ler(COM_PROBLEMAS)
    assert problemas == []
    assert len(entradas) == 7
    assert {e.chave for e in entradas} >= {"silva2020", "souza2021", "livro_sem_editora"}
