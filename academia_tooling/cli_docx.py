"""CLI do verificador de manuscrito: `check_docx.py --arquivo X.docx --norma abnt`."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from .docx_check import CHECAGENS, LIMITE_DA_CHECAGEM, auditar
from .laudo import COD_USO, codigo_saida, como_json, contar, imprimir_laudo
from .normas import CAMINHO_PADRAO, NormaInvalida, carregar, fontes, obter


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="check_docx.py",
        description="Laudo de conformidade de manuscrito .docx contra uma norma (ABNT/REGN/RBCS).",
        epilog=(
            "Códigos de saída: 0 = nenhum FAIL; 1 = há FAIL; 2 = erro de uso/norma; "
            "3 = --estrito e alguma checagem NAO_VERIFICAVEL."
        ),
    )
    parser.add_argument("--arquivo", required=True, type=Path, help="manuscrito .docx")
    parser.add_argument("--norma", required=True, help="slug da norma em normas.json (abnt, regn, rbcs, ...)")
    parser.add_argument(
        "--normas",
        type=Path,
        default=CAMINHO_PADRAO,
        help=f"JSON de limites (padrão: {CAMINHO_PADRAO.name} do repo)",
    )
    parser.add_argument("--json", action="store_true", help="saída em JSON em vez de tabela")
    parser.add_argument(
        "--estrito",
        action="store_true",
        help="trata NAO_VERIFICAVEL como reprovação (exit 3) e 'entradas não citadas' como FAIL",
    )
    parser.add_argument("--quieto", action="store_true", help="só o código de saída, sem laudo")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)
    try:
        normas = carregar(args.normas)
        norma = obter(normas, args.norma)
        if not args.arquivo.is_file():
            raise NormaInvalida(f"arquivo não encontrado: {args.arquivo}")
        resultados = auditar(args.arquivo, norma, estrito=args.estrito)
    except NormaInvalida as exc:
        print(f"erro: {exc}", file=sys.stderr)
        return COD_USO
    except Exception as exc:  # .docx corrompido, zip inválido, etc.
        print(f"erro ao ler {args.arquivo}: {type(exc).__name__}: {exc}", file=sys.stderr)
        return COD_USO

    codigo = codigo_saida(resultados, estrito=args.estrito)
    contextual = {
        "arquivo": str(args.arquivo),
        "norma": norma.slug,
        "norma_nome": norma.nome,
        "normas_arquivo": str(args.normas),
        "estrito": args.estrito,
        "gerado_em": datetime.now().astimezone().isoformat(timespec="seconds"),
    }

    if args.quieto:
        return codigo
    if args.json:
        print(como_json(contexto=contextual, resultados=resultados))
        return codigo

    contagem = contar(resultados)
    com_limite = sum(1 for cid, _, _ in CHECAGENS if norma.citavel(LIMITE_DA_CHECAGEM[cid]))
    rodape = [
        f"limites citáveis ({len(fontes(norma))}):",
        *[f"  · {chave}: {texto}" for chave, texto in fontes(norma).items()],
    ]
    if norma.nota:
        rodape += ["", f"nota da norma: {norma.nota}"]
    if com_limite == 0:
        rodape += ["", "ATENÇÃO: nenhuma checagem desta norma tem limite citável — nada foi verificado."]
    rodape.append(f"exit {codigo}")
    imprimir_laudo(
        titulo="verificador de manuscrito .docx",
        campos=[
            ("arquivo", str(args.arquivo)),
            ("norma", f"{norma.nome} [{norma.slug}]"),
            ("limites", f"{com_limite} de {len(CHECAGENS)} checagens com limite citável"),
        ],
        resultados=resultados,
        rodape=rodape,
    )
    return codigo


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
