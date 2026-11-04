"""CLIs: códigos de saída, laudo em texto, JSON e --normas próprio."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from academia_tooling.cli_bib import main as main_bib
from academia_tooling.cli_docx import main as main_docx
from academia_tooling.laudo import (
    COD_FAIL,
    COD_NAO_VERIFICAVEL,
    COD_OK,
    FAIL,
    MOTIVO_DOCUMENTO,
    MOTIVO_NORMA,
    NAO_VERIFICAVEL,
    PASS,
    Resultado,
    codigo_saida,
)

FIX = Path(__file__).resolve().parent / "fixtures"
REPO = Path(__file__).resolve().parent.parent


def _normas(tmp_path: Path, limites: dict) -> Path:
    caminho = tmp_path / "normas.json"
    caminho.write_text(
        json.dumps({"normas": {"teste": {"nome": "Edital de teste", "limites": limites}}}),
        encoding="utf-8",
    )
    return caminho


# ------------------------------------------------------------------ check_docx


def test_docx_conforme_abnt_exit_0(docx, capsys):
    codigo = main_docx(["--arquivo", str(docx["conforme"]), "--norma", "abnt"])
    saida = capsys.readouterr().out
    assert codigo == 0
    assert "0 FAIL" in saida
    assert "ABNT NBR 14724" in saida
    assert "ABNT NBR 14724:2011 — margens" in saida, "o laudo precisa mostrar a fonte do limite"


def test_docx_fora_de_norma_exit_1(docx, capsys):
    codigo = main_docx(["--arquivo", str(docx["fora_de_norma"]), "--norma", "abnt"])
    saida = capsys.readouterr().out
    assert codigo == 1
    for checagem in ("fonte_tamanho", "espacamento", "margens", "secoes", "referencias"):
        assert checagem in saida


def test_docx_estrito_reprova_nao_verificavel(docx, capsys):
    assert main_docx(["--arquivo", str(docx["conforme"]), "--norma", "abnt"]) == 0
    capsys.readouterr()
    assert main_docx(["--arquivo", str(docx["conforme"]), "--norma", "abnt", "--estrito"]) == 3


def test_docx_rbcs_reprova_por_faixa_de_palavras(docx, capsys):
    codigo = main_docx(["--arquivo", str(docx["conforme"]), "--norma", "rbcs"])
    saida = capsys.readouterr().out
    assert codigo == 1
    assert "7000" in saida or "7.000" in saida or "fora da faixa" in saida


def test_docx_regn_nao_verifica_nada_e_avisa(docx, capsys):
    codigo = main_docx(["--arquivo", str(docx["conforme"]), "--norma", "regn"])
    saida = capsys.readouterr().out
    assert codigo == 0
    assert "8 NAO_VERIFICAVEL" in saida
    assert "nada foi verificado" in saida


def test_docx_json_estruturado(docx, capsys):
    main_docx(["--arquivo", str(docx["fora_de_norma"]), "--norma", "abnt", "--json"])
    dados = json.loads(capsys.readouterr().out)
    assert dados["contexto"]["norma"] == "abnt"
    assert dados["resumo"]["FAIL"] >= 1
    assert {c["id"] for c in dados["checagens"]} >= {"margens", "referencias"}
    assert all(c["mensagem"] for c in dados["checagens"])


def test_docx_norma_desconhecida_exit_2(docx, capsys):
    assert main_docx(["--arquivo", str(docx["conforme"]), "--norma", "nature"]) == 2
    assert "norma desconhecida" in capsys.readouterr().err


def test_docx_arquivo_inexistente_exit_2(tmp_path, capsys):
    assert main_docx(["--arquivo", str(tmp_path / "nao_existe.docx"), "--norma", "abnt"]) == 2
    assert "não encontrado" in capsys.readouterr().err


def test_docx_arquivo_corrompido_exit_2(tmp_path, capsys):
    ruim = tmp_path / "quebrado.docx"
    ruim.write_text("isto não é um docx", encoding="utf-8")
    assert main_docx(["--arquivo", str(ruim), "--norma", "abnt"]) == 2
    assert "erro ao ler" in capsys.readouterr().err


def test_docx_quieto_nao_imprime(docx, capsys):
    assert main_docx(["--arquivo", str(docx["fora_de_norma"]), "--norma", "abnt", "--quieto"]) == 1
    assert capsys.readouterr().out == ""


# --------------------------------------------- exit code e laudo


def test_docx_10pt_em_todos_os_runs_reprova(docx, capsys):
    """Fixture permanente: Normal 12 pt, corpo inteiro em 10 pt Arial → FAIL e exit 1."""
    codigo = main_docx(["--arquivo", str(docx["direto_10pt"]), "--norma", "abnt"])
    saida = capsys.readouterr().out
    assert codigo == 1
    linha = next(linha for linha in saida.splitlines() if linha.startswith("fonte_tamanho"))
    assert "FAIL" in linha, linha
    assert "PASS" not in linha, linha
    achatada = " ".join(saida.split())  # a tabela quebra a mensagem em 60 colunas
    assert "runs fora do padrão (10 pt)" in achatada
    assert "exigido 12 pt" in achatada
    assert "exit 1" in saida


def test_docx_nao_verificavel_de_documento_exit_3(docx, capsys):
    """Espaçamento exato: a norma fixa 1,5, o documento não permite verificar → exit 3."""
    codigo = main_docx(["--arquivo", str(docx["espaco_exato"]), "--norma", "abnt"])
    saida = capsys.readouterr().out
    assert codigo == 3, saida
    linha = next(linha for linha in saida.splitlines() if linha.startswith("espacamento"))
    assert "NAO_VERIFICAVEL" in linha
    assert "documento não permitiu verificar: 1" in saida
    assert "exit 3" in saida


def test_docx_nao_verificavel_de_norma_continua_exit_0(docx, capsys):
    """REGN não fixa nada: NAO_VERIFICAVEL por falta de limite é a resposta certa, exit 0."""
    assert main_docx(["--arquivo", str(docx["conforme"]), "--norma", "regn"]) == 0
    assert "documento não permitiu verificar: 0" in capsys.readouterr().out


def test_codigo_saida_separa_norma_de_documento():
    def resultado(id_: str, status: str, motivo: str | None = None) -> Resultado:
        return Resultado(id_, id_, status, "mensagem", {}, motivo_nao_verificavel=motivo)

    passa = resultado("a", PASS)
    norma = resultado("b", NAO_VERIFICAVEL, MOTIVO_NORMA)
    documento = resultado("c", NAO_VERIFICAVEL, MOTIVO_DOCUMENTO)
    falha = resultado("d", FAIL)

    assert codigo_saida([passa, norma]) == COD_OK
    assert codigo_saida([passa, documento]) == COD_NAO_VERIFICAVEL  # exit 3 por default
    assert codigo_saida([passa, norma], estrito=True) == COD_NAO_VERIFICAVEL
    assert codigo_saida([documento, falha]) == COD_FAIL  # FAIL vence


def test_resultado_recusa_nao_verificavel_sem_motivo():
    with pytest.raises(ValueError):
        Resultado("x", "x", NAO_VERIFICAVEL, "sem motivo")
    with pytest.raises(ValueError):
        Resultado("x", "x", PASS, "pass com motivo", motivo_nao_verificavel=MOTIVO_NORMA)


def test_docx_normas_proprias(tmp_path, docx, capsys):
    caminho = _normas(
        tmp_path,
        {
            "revisao_cega": {"valor": True, "fonte": "edital próprio: revisão cega"},
            "tamanho_fonte_pt": {"valor": 10, "fonte": "edital próprio: corpo 10"},
        },
    )
    codigo = main_docx(
        ["--arquivo", str(docx["fora_de_norma"]), "--norma", "teste", "--normas", str(caminho)]
    )
    saida = capsys.readouterr().out
    assert codigo == 1
    assert "autor_metadado" in saida
    assert "revisão cega" in saida


# ------------------------------------------------------------------ check_bib


def test_bib_com_problemas_exit_1(capsys):
    codigo = main_bib(
        ["--bib", str(FIX / "bib_com_problemas.bib"), "--texto", str(FIX / "texto_problemas.tex")]
    )
    saida = capsys.readouterr().out
    assert codigo == 1
    assert "doi_duplicado" in saida and "citacao_sem_entrada" in saida
    assert "ERRO" in saida and "AVISO" in saida


def test_bib_limpo_exit_0(capsys):
    assert main_bib(["--bib", str(FIX / "bib_ok.bib"), "--texto", str(FIX / "texto_ok.tex")]) == 0
    assert "nenhum achado" in capsys.readouterr().out


def test_bib_json(capsys):
    main_bib(["--bib", str(FIX / "bib_com_problemas.bib"), "--texto", str(FIX / "texto_problemas.tex"), "--json"])
    dados = json.loads(capsys.readouterr().out)
    assert dados["entradas"] == 7
    assert dados["erros"] >= 1
    assert {a["categoria"] for a in dados["achados"]} >= {"doi_duplicado", "ano"}


def test_bib_sem_texto_lista_o_que_nao_rodou(capsys):
    main_bib(["--bib", str(FIX / "bib_ok.bib")])
    saida = capsys.readouterr().out
    assert "chaves_nao_citadas" in saida and "citacoes_sem_entrada" in saida


def test_bib_com_docx_usa_texto_do_docx(docx, capsys):
    codigo = main_bib(["--bib", str(FIX / "bib_ok.bib"), "--docx", str(docx["conforme"])])
    saida = capsys.readouterr().out
    assert codigo == 0
    assert "autor-ano" in saida


def test_bib_estrito_reprova_aviso(capsys):
    caminho = FIX / "bib_ok.bib"
    assert main_bib(["--bib", str(caminho)]) == 0
    capsys.readouterr()
    assert main_bib(["--bib", str(caminho), "--estrito"]) == 0  # sem texto, nada de AVISO novo


def test_bib_arquivo_inexistente_exit_2(tmp_path, capsys):
    assert main_bib(["--bib", str(tmp_path / "x.bib")]) == 2
    assert "não encontrado" in capsys.readouterr().err


def test_bib_texto_inexistente_exit_2(tmp_path, capsys):
    assert main_bib(["--bib", str(FIX / "bib_ok.bib"), "--texto", str(tmp_path / "x.tex")]) == 2


def test_bib_quieto(capsys):
    assert main_bib(["--bib", str(FIX / "bib_com_problemas.bib"), "--quieto"]) == 1
    assert capsys.readouterr().out == ""


# ------------------------------------------------------------------ scripts da raiz


@pytest.mark.parametrize("script", ["check_docx.py", "check_bib.py"])
def test_scripts_da_raiz_existem_e_apontam_para_o_pacote(script):
    texto = (REPO / script).read_text(encoding="utf-8")
    assert "raise SystemExit(main())" in texto
    assert (REPO / "normas.json").is_file()
