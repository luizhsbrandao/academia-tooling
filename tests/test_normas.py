"""Contrato de `normas.json`: só limite com fonte citável autoriza PASS/FAIL."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from academia_tooling.normas import (
    CAMINHO_PADRAO,
    Limite,
    NormaInvalida,
    carregar,
    fontes,
    obter,
)


def _escrever(tmp_path: Path, limites: dict) -> Path:
    caminho = tmp_path / "normas.json"
    caminho.write_text(
        json.dumps({"normas": {"teste": {"nome": "Norma de teste", "limites": limites}}}),
        encoding="utf-8",
    )
    return caminho


def test_normas_do_repo_carregam():
    normas = carregar(CAMINHO_PADRAO)
    assert set(normas) == {"abnt", "rbcs", "regn"}
    assert all(n.nome for n in normas.values())


def test_abnt_fixa_margens_e_corpo():
    abnt = carregar(CAMINHO_PADRAO)["abnt"]
    assert abnt.limite("margens_cm").valor == {
        "superior": 3.0,
        "inferior": 2.0,
        "esquerda": 3.0,
        "direita": 2.0,
    }
    assert abnt.limite("tamanho_fonte_pt").valor == 12.0
    assert abnt.limite("espacamento_linhas").valor == 1.5


def test_rbcs_tem_faixa_de_palavras_citavel():
    rbcs = carregar(CAMINHO_PADRAO)["rbcs"]
    limite = rbcs.limite("palavras")
    assert limite.citavel
    assert limite.valor == {"min": 7000, "max": 9000}
    assert "RBCS" in limite.fonte


def test_abnt_nao_chuta_familia_nem_palavras():
    abnt = carregar(CAMINHO_PADRAO)["abnt"]
    for chave in ("familia_fonte", "palavras", "revisao_cega"):
        assert not abnt.citavel(chave), f"{chave} não é fixado pela NBR 14724 e deveria ser null"


def test_regn_nao_tem_nenhum_limite_citavel():
    regn = carregar(CAMINHO_PADRAO)["regn"]
    assert fontes(regn) == {}
    assert all(not regn.citavel(chave) for chave in regn.limites)


def test_normas_exemplo_carrega_e_cobre_as_oito_checagens():
    """O modelo de edital serve de cópia: se ele quebrar, a instrução do README quebra."""
    exemplo = carregar(CAMINHO_PADRAO.parent / "normas.exemplo.json")["meu-edital"]
    esperados = {
        "margens_cm",
        "tamanho_fonte_pt",
        "familia_fonte",
        "espacamento_linhas",
        "palavras",
        "contagem_exclui_referencias",
        "secoes_obrigatorias",
        "revisao_cega",
    }
    assert esperados <= set(exemplo.limites)
    assert all(exemplo.citavel(chave) for chave in esperados)


def test_limite_sem_fonte_e_rejeitado(tmp_path):
    caminho = _escrever(tmp_path, {"tamanho_fonte_pt": {"valor": 12}})
    with pytest.raises(NormaInvalida, match="fonte"):
        carregar(caminho)


def test_limite_com_fonte_vazia_e_rejeitado(tmp_path):
    caminho = _escrever(tmp_path, {"tamanho_fonte_pt": {"valor": 12, "fonte": "   "}})
    with pytest.raises(NormaInvalida):
        carregar(caminho)


def test_limite_null_vira_nao_citavel(tmp_path):
    caminho = _escrever(tmp_path, {"familia_fonte": None})
    norma = carregar(caminho)["teste"]
    assert norma.limite("familia_fonte") == Limite("familia_fonte")
    assert not norma.citavel("familia_fonte")


def test_palavras_sem_modo_de_contagem_e_rejeitada(tmp_path):
    caminho = _escrever(
        tmp_path,
        {"palavras": {"valor": {"min": 100, "max": 200}, "fonte": "edital X"}},
    )
    with pytest.raises(NormaInvalida, match="contagem_exclui_referencias"):
        carregar(caminho)


def test_margens_incompletas_rejeitadas(tmp_path):
    caminho = _escrever(
        tmp_path,
        {"margens_cm": {"valor": {"superior": 3, "inferior": 2}, "fonte": "edital X"}},
    )
    with pytest.raises(NormaInvalida, match="margens_cm"):
        carregar(caminho)


def test_secao_desconhecida_rejeitada(tmp_path):
    caminho = _escrever(
        tmp_path,
        {"secoes_obrigatorias": {"valor": ["titulo", "anexos"], "fonte": "edital X"}},
    )
    with pytest.raises(NormaInvalida, match="desconhecidas"):
        carregar(caminho)


def test_faixa_de_palavras_invertida_rejeitada(tmp_path):
    caminho = _escrever(
        tmp_path,
        {
            "palavras": {"valor": {"min": 9000, "max": 7000}, "fonte": "edital X"},
            "contagem_exclui_referencias": {"valor": True, "fonte": "edital X"},
        },
    )
    with pytest.raises(NormaInvalida, match="faixa inválida"):
        carregar(caminho)


def test_limite_desconhecido_rejeitado(tmp_path):
    caminho = _escrever(tmp_path, {"entrelinhas": {"valor": 2, "fonte": "edital X"}})
    with pytest.raises(NormaInvalida, match="desconhecido"):
        carregar(caminho)


def test_json_invalido_e_arquivo_ausente(tmp_path):
    quebrado = tmp_path / "quebrado.json"
    quebrado.write_text("{ isso não é json", encoding="utf-8")
    with pytest.raises(NormaInvalida, match="JSON inválido"):
        carregar(quebrado)
    with pytest.raises(NormaInvalida, match="não encontrado"):
        carregar(tmp_path / "nao_existe.json")


def test_norma_desconhecida(tmp_path):
    normas = carregar(CAMINHO_PADRAO)
    with pytest.raises(NormaInvalida, match="norma desconhecida"):
        obter(normas, "nature")
