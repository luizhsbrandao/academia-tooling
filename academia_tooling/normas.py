"""Carga e validação de `normas.json`.

Contrato: cada limite é `null` (norma não fixa / não citável => NAO_VERIFICAVEL)
ou `{"valor": ..., "fonte": "citação não vazia"}`. Um limite sem fonte é rejeitado
na carga — é assim que "nunca chutar" deixa de ser recomendação e vira código.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

RAIZ = Path(__file__).resolve().parent.parent
CAMINHO_PADRAO = RAIZ / "normas.json"

SECOES_CONHECIDAS = frozenset({"titulo", "resumo", "abstract", "referencias"})
LADOS_MARGEM = ("superior", "inferior", "esquerda", "direita")


class NormaInvalida(ValueError):
    """normas.json malformado, ou limite sem fonte citável."""


@dataclass(frozen=True)
class Limite:
    chave: str
    valor: Any = None
    fonte: str | None = None

    @property
    def citavel(self) -> bool:
        return self.valor is not None and bool(self.fonte)


@dataclass(frozen=True)
class Norma:
    slug: str
    nome: str
    limites: dict[str, Limite] = field(default_factory=dict)
    nota: str = ""

    def limite(self, chave: str) -> Limite:
        return self.limites.get(chave, Limite(chave))

    def citavel(self, chave: str) -> bool:
        return self.limite(chave).citavel


def _exigir_fonte(chave: str, dado: Any) -> str:
    if not isinstance(dado, dict):
        raise NormaInvalida(f"limite {chave!r}: esperado objeto com 'valor' e 'fonte', veio {type(dado).__name__}")
    fonte = dado.get("fonte")
    if not isinstance(fonte, str) or not fonte.strip():
        raise NormaInvalida(f"limite {chave!r}: sem 'fonte' citável — use null se a norma não fixa o parâmetro")
    if "valor" not in dado:
        raise NormaInvalida(f"limite {chave!r}: objeto sem 'valor'")
    return fonte.strip()


def _validar_valor(chave: str, valor: Any) -> Any:
    if chave == "margens_cm":
        if not isinstance(valor, dict) or set(valor) != set(LADOS_MARGEM):
            raise NormaInvalida(f"margens_cm: esperado objeto com {LADOS_MARGEM}, veio {valor!r}")
        for lado, cm in valor.items():
            if not isinstance(cm, (int, float)) or cm <= 0:
                raise NormaInvalida(f"margens_cm.{lado}: esperado número > 0, veio {cm!r}")
        return {lado: float(cm) for lado, cm in valor.items()}
    if chave in ("tamanho_fonte_pt", "espacamento_linhas"):
        if not isinstance(valor, (int, float)) or valor <= 0:
            raise NormaInvalida(f"{chave}: esperado número > 0, veio {valor!r}")
        return float(valor)
    if chave == "familia_fonte":
        fontes = [valor] if isinstance(valor, str) else valor
        if not isinstance(fontes, list) or not fontes or not all(isinstance(f, str) and f.strip() for f in fontes):
            raise NormaInvalida(f"familia_fonte: esperado string ou lista de strings, veio {valor!r}")
        return [f.strip() for f in fontes]
    if chave == "palavras":
        if not isinstance(valor, dict) or not {"min", "max"} <= set(valor):
            raise NormaInvalida(f"palavras: esperado {{'min': n, 'max': n}}, veio {valor!r}")
        minimo, maximo = valor["min"], valor["max"]
        if not all(isinstance(n, int) for n in (minimo, maximo)) or minimo < 0 or maximo <= minimo:
            raise NormaInvalida(f"palavras: faixa inválida {valor!r}")
        return {"min": int(minimo), "max": int(maximo)}
    if chave == "secoes_obrigatorias":
        if not isinstance(valor, list) or not valor:
            raise NormaInvalida(f"secoes_obrigatorias: esperado lista, veio {valor!r}")
        desconhecidas = [s for s in valor if s not in SECOES_CONHECIDAS]
        if desconhecidas:
            raise NormaInvalida(
                f"secoes_obrigatorias: seções desconhecidas {desconhecidas} (conhecidas: {sorted(SECOES_CONHECIDAS)})"
            )
        return list(valor)
    if chave == "revisao_cega":
        if not isinstance(valor, bool):
            raise NormaInvalida(f"revisao_cega: esperado booleano, veio {valor!r}")
        return valor
    if chave == "contagem_exclui_referencias":
        if not isinstance(valor, bool):
            raise NormaInvalida(f"contagem_exclui_referencias: esperado booleano, veio {valor!r}")
        return valor
    raise NormaInvalida(f"limite desconhecido: {chave!r}")


def carregar(caminho: str | Path = CAMINHO_PADRAO) -> dict[str, Norma]:
    """Lê e valida o JSON de normas. Retorna {slug: Norma}."""
    caminho = Path(caminho)
    try:
        bruto = json.loads(caminho.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise NormaInvalida(f"arquivo de normas não encontrado: {caminho}") from exc
    except json.JSONDecodeError as exc:
        raise NormaInvalida(f"{caminho}: JSON inválido ({exc})") from exc

    normas_brutas = bruto.get("normas")
    if not isinstance(normas_brutas, dict) or not normas_brutas:
        raise NormaInvalida(f"{caminho}: falta o objeto 'normas'")

    normas: dict[str, Norma] = {}
    for slug, dado in normas_brutas.items():
        if not isinstance(dado, dict):
            raise NormaInvalida(f"norma {slug!r}: esperado objeto")
        nome = dado.get("nome")
        if not isinstance(nome, str) or not nome.strip():
            raise NormaInvalida(f"norma {slug!r}: sem 'nome'")
        limites_brutos = dado.get("limites") or {}
        if not isinstance(limites_brutos, dict):
            raise NormaInvalida(f"norma {slug!r}: 'limites' deve ser objeto")

        limites: dict[str, Limite] = {}
        for chave, valor_bruto in limites_brutos.items():
            if valor_bruto is None:
                limites[chave] = Limite(chave)
                if chave == "contagem_exclui_referencias":
                    raise NormaInvalida(
                        f"norma {slug!r}: 'contagem_exclui_referencias' sem valor não faz sentido — remova a chave"
                    )
                continue
            fonte = _exigir_fonte(chave, valor_bruto)
            limites[chave] = Limite(chave, _validar_valor(chave, valor_bruto["valor"]), fonte)

        # Faixa de palavras e modo de contagem andam juntos: sem o modo, o
        # veredito dependeria de adivinhação sobre o que a revista conta.
        if limites.get("palavras") and limites["palavras"].citavel and not limites.get("contagem_exclui_referencias", Limite("x")).citavel:
            raise NormaInvalida(
                f"norma {slug!r}: 'palavras' exige 'contagem_exclui_referencias' com fonte "
                "(dizer se a faixa conta a lista de referências)"
            )

        normas[slug] = Norma(slug=slug, nome=nome.strip(), limites=limites, nota=dado.get("nota", ""))
    return normas


def obter(normas: dict[str, Norma], slug: str) -> Norma:
    try:
        return normas[slug]
    except KeyError as exc:
        raise NormaInvalida(
            f"norma desconhecida: {slug!r} (disponíveis: {', '.join(sorted(normas))})"
        ) from exc


def fontes(norma: Norma) -> dict[str, str]:
    """Só os limites citáveis, com a fonte — vai no rodapé do laudo/JSON."""
    return {chave: lim.fonte for chave, lim in norma.limites.items() if lim.citavel and lim.fonte}
