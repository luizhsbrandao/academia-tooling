"""Fixtures compartilhadas dos testes.

Os .docx são gerados (nunca vêm de vault); os .bib/.tex são texto versionado.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:  # roda sem instalar o pacote
    sys.path.insert(0, str(RAIZ))

from academia_tooling import fixtures as gerador  # noqa: E402
from academia_tooling.normas import Limite, Norma  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"


@pytest.fixture(scope="session")
def docx() -> dict[str, Path]:
    """Caminhos dos .docx sintéticos (conforme, fora_de_norma, fragmento)."""
    return gerador.construir(FIXTURES / "gerados")


@pytest.fixture
def norma_factory():
    """Fabricar Norma em memória, sem passar por JSON (casos pontuais)."""

    def fabricar(**limites) -> Norma:
        registrados = {}
        for chave, valor in limites.items():
            registrados[chave] = Limite(chave) if valor is None else Limite(chave, valor, "fonte de teste")
        return Norma(slug="teste", nome="norma de teste", limites=registrados)

    return fabricar


@pytest.fixture
def documento():
    """Abrir um .docx como objeto python-docx."""
    from docx import Document

    def abrir(caminho: Path):
        return Document(str(caminho))

    return abrir
