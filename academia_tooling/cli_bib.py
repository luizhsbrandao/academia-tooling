"""CLI do auditor de bibliografia: `check_bib.py --bib X.bib [--texto t.tex|--docx m.docx]`."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from .bib_check import ANO_MAX_PADRAO, ANO_MIN_PADRAO, AVISO, ERRO, auditar_arquivo, codigo_saida
from .docx_check import texto_do_docx
from .laudo import REGUA, REGUA_FINA, tabela


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="check_bib.py",
        description="Auditor de bibliografia .bib: duplicatas, campos obrigatórios, ano, identificador e citações.",
        epilog="Códigos de saída: 0 = nenhum ERRO; 1 = há ERRO; 2 = erro de uso.",
    )
    parser.add_argument("--bib", required=True, type=Path, help="arquivo .bib")
    fonte = parser.add_mutually_exclusive_group()
    fonte.add_argument("--texto", type=Path, help=".tex/.md/.txt com as citações (\\cite{}, [@chave], chave literal)")
    fonte.add_argument("--docx", type=Path, help="manuscrito .docx (texto extraído para o cruzamento)")
    parser.add_argument("--ano-min", type=int, default=ANO_MIN_PADRAO)
    parser.add_argument("--ano-max", type=int, default=ANO_MAX_PADRAO)
    parser.add_argument("--estrito", action="store_true", help="AVISO também reprova (exit 1)")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--quieto", action="store_true", help="só o código de saída")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)

    if not args.bib.is_file():
        print(f"erro: .bib não encontrado: {args.bib}", file=sys.stderr)
        return 2
    texto = None
    if args.texto is not None:
        if not args.texto.is_file():
            print(f"erro: texto não encontrado: {args.texto}", file=sys.stderr)
            return 2
        texto = args.texto.read_text(encoding="utf-8", errors="replace")
    elif args.docx is not None:
        if not args.docx.is_file():
            print(f"erro: .docx não encontrado: {args.docx}", file=sys.stderr)
            return 2
        try:
            texto = texto_do_docx(args.docx)
        except Exception as exc:
            print(f"erro ao ler {args.docx}: {type(exc).__name__}: {exc}", file=sys.stderr)
            return 2

    try:
        relatorio = auditar_arquivo(
            args.bib,
            texto=texto,
            ano_min=args.ano_min,
            ano_max=args.ano_max,
            estrito=args.estrito,
        )
    except OSError as exc:
        print(f"erro ao ler {args.bib}: {exc}", file=sys.stderr)
        return 2

    codigo = codigo_saida(relatorio)
    if args.quieto:
        return codigo

    if args.json:
        print(
            json.dumps(
                {
                    "contexto": {
                        "bib": str(args.bib),
                        "texto": str(args.texto or args.docx or ""),
                        "estrito": args.estrito,
                        "gerado_em": datetime.now().astimezone().isoformat(timespec="seconds"),
                    },
                    "entradas": relatorio.entradas,
                    "erros": len(relatorio.erros),
                    "avisos": len(relatorio.avisos),
                    "achados": [
                        {
                            "severidade": a.severidade,
                            "categoria": a.categoria,
                            "chaves": list(a.chaves),
                            "mensagem": a.mensagem,
                        }
                        for a in relatorio.achados
                    ],
                    "notas": relatorio.notas,
                    "checagens_nao_executadas": relatorio.checagens_nao_executadas,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return codigo

    print(REGUA)
    print(" auditor de bibliografia .bib")
    print(REGUA_FINA)
    print(f" bib       {args.bib}")
    print(f" cruza     {args.texto or args.docx or '(nenhum: só o .bib)'}")
    print(f" faixa ano {args.ano_min}–{args.ano_max}" + ("  · modo estrito" if args.estrito else ""))
    print(REGUA_FINA)
    if relatorio.achados:
        linhas = [
            [a.severidade, a.categoria, ", ".join(a.chaves) or "—", a.mensagem] for a in relatorio.achados
        ]
        print("\n".join(tabela(["SEV", "CATEGORIA", "CHAVE(S)", "MENSAGEM"], linhas, maximos=[5, 24, 26, 60])))
    else:
        print(" nenhum achado")
    for nota in relatorio.notas:
        print(f" nota: {nota}")
    print(REGUA_FINA)
    print(
        f" {relatorio.entradas} entradas · {len(relatorio.erros)} ERRO · {len(relatorio.avisos)} AVISO"
        f" · não executadas: {', '.join(relatorio.checagens_nao_executadas) or 'nenhuma'}"
    )
    print(f" exit {codigo}")
    print(REGUA)
    return codigo


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
