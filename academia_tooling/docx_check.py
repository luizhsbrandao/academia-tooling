"""Checagens objetivas de conformidade de manuscrito .docx.

Cada checagem devolve PASS, FAIL ou NAO_VERIFICAVEL:

* PASS/FAIL — existe limite citável em `normas.json` e o .docx declara o atributo;
* NAO_VERIFICAVEL — a norma não fixa o parâmetro (limite null/sem fonte) ou o
  .docx não declara o atributo de forma legível por `python-docx` (herança de
  tema, espaçamento exato, etc.). Nunca se converte isso em PASS.

Duas regras que o veredito precisa respeitar (senão o laudo mente):

* **PASS não pode conviver com desvio material.** A formatação direta vence o
  estilo no Word: se mais de `LIMIAR_DESVIO` dos runs/parágrafos do corpo
  divergem do padrão declarado — seja por tamanho, família ou espaçamento —, o
  status é FAIL, com a contagem no veredito.
* **"Não verificável" tem motivo.** `motivo_nao_verificavel` diz se foi a norma
  que não fixa o parâmetro (`MOTIVO_NORMA`, não reprova) ou o documento que não
  permite verificar (`MOTIVO_DOCUMENTO`, exit 3).
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Length, Pt

from .laudo import FAIL, MOTIVO_DOCUMENTO, MOTIVO_NORMA, NAO_VERIFICAVEL, PASS, Resultado
from .normas import Norma
from .texto import PARTICULAS, achatar, chave, chave_titulo, contar_palavras, numero

TOL_CM = 0.08  # 3 cm = 1701 twips ≈ 3,0002 cm; a folga cobre arredondamento do Word
TOL_PT = 0.05
TOL_ESPACO = 0.02

# Fração máxima de runs/parágrafos do corpo que pode divergir do padrão (por
# formatação direta) sem reprovar. Acima disso o atributo declarado no estilo não
# descreve o documento: é PASS por estatística, e o laudo não pode dar PASS.
LIMIAR_DESVIO = 0.10

MARGENS = (
    ("superior", "top_margin"),
    ("inferior", "bottom_margin"),
    ("esquerda", "left_margin"),
    ("direita", "right_margin"),
)

ANONIMOS = frozenset(
    {
        "autor",
        "autores",
        "autor a",
        "anonimo",
        "anonymous",
        "anônimo",
        "revisor",
        "sem autor",
        "blind",
        "cega",
        "cego",
    }
)

# ---------------------------------------------------------------- docDefaults


def _docdefaults_rpr(doc) -> tuple[str | None, float | None]:
    """(família, tamanho em pt) declarados em w:docDefaults/w:rPrDefault."""
    elemento = doc.styles.element
    nomes = elemento.xpath("./w:docDefaults/w:rPrDefault/w:rPr/w:rFonts/@w:ascii")
    tamanhos = elemento.xpath("./w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")
    pts = float(tamanhos[0]) / 2.0 if tamanhos else None  # w:sz é meio-ponto
    return (nomes[0] if nomes else None), pts


def _docdefaults_espaco(doc) -> tuple[float | None, bool]:
    """(múltiplo, é_espaçamento_fixo) de w:docDefaults/w:pPrDefault."""
    elemento = doc.styles.element
    linhas = elemento.xpath("./w:docDefaults/w:pPrDefault/w:pPr/w:spacing/@w:line")
    regras = elemento.xpath("./w:docDefaults/w:pPrDefault/w:pPr/w:spacing/@w:lineRule")
    if not linhas:
        return None, False
    regra = regras[0] if regras else "auto"
    if regra != "auto":
        return None, True  # exato/atLeast fixa altura de linha, não é múltiplo comparável
    return float(linhas[0]) / 240.0, False


def _fonte_normal(doc) -> tuple[str | None, float | None, str]:
    """Família/tamanho do estilo Normal, com fallback declarado para docDefaults."""
    nome = tamanho = None
    origem = "estilo Normal"
    try:
        estilo = doc.styles["Normal"]
    except KeyError:
        estilo = None
    if estilo is not None:
        nome = estilo.font.name
        if estilo.font.size is not None:
            tamanho = float(estilo.font.size.pt)
    if nome is None or tamanho is None:
        nome_padrao, tamanho_padrao = _docdefaults_rpr(doc)
        herdou = False
        if nome is None and nome_padrao:
            nome, herdou = nome_padrao, True
        if tamanho is None and tamanho_padrao:
            tamanho, herdou = tamanho_padrao, True
        if herdou:
            origem = "docDefaults (herdado; estilo Normal não declara)"
        elif nome is None and tamanho is None:
            origem = "nada declarado"
    return nome, tamanho, origem


def _multiplo(valor) -> float | None:
    """line_spacing -> múltiplo; Length (espaçamento fixo) não é comparável."""
    if valor is None or isinstance(valor, Length):
        return None
    return float(valor)


def _espaco_declarado(pPr) -> tuple[float | None, bool]:
    """(múltiplo, é_espaçamento_fixo) do `w:pPr` de um parágrafo ou estilo.

    `w:lineRule` ausente vale "auto" no OOXML (múltiplo de 240 twips, como o Word
    grava 1,5 linha em `<w:spacing w:line="360"/>`); "exact" e "atLeast" fixam
    altura de linha em pontos e **não** são múltiplos comparáveis com 1,5.
    """
    if pPr is None:
        return None, False
    linha = pPr.spacing_line
    if linha is None:
        return None, False
    regra = pPr.spacing_lineRule
    if regra is None or regra == WD_LINE_SPACING.MULTIPLE:
        return linha / Pt(12), False
    return None, True


def _espaco_do_estilo(doc, nome: str = "Normal") -> tuple[float | None, bool]:
    try:
        estilo = doc.styles[nome]
    except KeyError:
        return None, False
    return _espaco_declarado(estilo.element.pPr)


def _runs_do_corpo(doc) -> list:
    """Runs com texto dos parágrafos com texto — o que o leitor vê como corpo.

    Run vazio não renderiza nada e não entra na amostra (senão inflaria a fração
    de "não legível" e transformaria documento legível em NAO_VERIFICAVEL).
    """
    return [
        run
        for paragrafo in doc.paragraphs
        if paragrafo.text.strip()
        for run in paragrafo.runs
        if run.text.strip()
    ]


def _tamanho_do_run(run, herdado: float | None) -> float | None:
    """Tamanho efetivo: formatação direta do run vence o estilo Normal."""
    if run.font.size is not None:
        return float(run.font.size.pt)
    return herdado


def _familia_do_run(run, herdada: str | None) -> str | None:
    return run.font.name or herdada


def _fracao(parte: int, total: int) -> float:
    return parte / total if total else 0.0


# ---------------------------------------------------------------- estrutura


_PADROES_CABECALHO: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("resumo", re.compile(r"^(resumo|resumo expandido|summary)\b", re.I)),
    ("abstract", re.compile(r"^(abstract|summary in english)\b", re.I)),
    (
        "referencias",
        re.compile(
            r"^(refer[eê]ncias(\s+bibliogr[aá]ficas)?|bibliografia|obras consultadas|fontes|references|works cited)\b",
            re.I,
        ),
    ),
    ("palavras_chave", re.compile(r"^(palavras[-\s]chave|keywords?|descritores)\b", re.I)),
)

_PADRAO_SECAO = re.compile(
    r"^(\d+(\.\d+)*[\.\)]?\s+)?("
    r"introdu[cç][aã]o|apresenta[cç][aã]o|justificativa|objetivos?|referencial(\s+te[oó]rico)?|"
    r"metodologia|m[ée]todos?|materiais\s+e\s+m[ée]todos|resultados|discuss[ãa]o|"
    r"considera[cç][oõ]es\s+finais|conclus[ãa]o|agradecimentos?|anexos?|ap[êe]ndices?"
    r")\b",
    re.I,
)


def _categoria_cabecalho(linha: str) -> str | None:
    texto = linha.strip()
    if not texto or contar_palavras(texto) > 12:
        return None
    for nome, padrao in _PADROES_CABECALHO:
        if padrao.match(texto):
            return nome
    if _PADRAO_SECAO.match(texto):
        return "secao"
    return None


@dataclass
class Estrutura:
    """Leitura heurística do esqueleto do documento."""

    linhas: list[str]
    titulo: str | None = None
    indices: dict[str, int] = field(default_factory=dict)
    trechos: list[tuple[str, str]] = field(default_factory=list)  # (seção, texto)
    corpo: str = ""

    @property
    def alguma_estrutura(self) -> bool:
        return bool(self.indices) or self.titulo is not None


def _estrutura(doc) -> Estrutura:
    linhas = [p.text.strip() for p in doc.paragraphs]
    indices: dict[str, int] = {}
    for i, linha in enumerate(linhas):
        categoria = _categoria_cabecalho(linha)
        if categoria and categoria not in indices:
            indices[categoria] = i

    titulo = None
    for linha in linhas:
        if not linha:
            continue
        # só a primeira linha com texto é candidata a título: heurística declarada
        if (
            _categoria_cabecalho(linha) is None
            and 3 <= contar_palavras(linha) <= 25
            and not linha.endswith((".", ":", ";"))
        ):
            titulo = linha
        break

    # quebra por cabeçalhos reconhecidos
    trechos: list[tuple[str, list[str]]] = [("corpo", [])]
    for i, linha in enumerate(linhas):
        categoria = _categoria_cabecalho(linha)
        if categoria:
            trechos.append((f"{categoria}:{achatar(linha, 40)}", []))
            continue
        trechos[-1][1].append(linha)

    partes = [(nome, "\n".join(ls).strip()) for nome, ls in trechos]
    corpo = partes[0][1] if partes else ""
    if "referencias" in indices:
        corte = indices["referencias"]
        corpo = "\n".join(linhas[:corte])
    return Estrutura(linhas=linhas, titulo=titulo, indices=indices, trechos=partes, corpo=corpo)


def _bloco_referencias(estrutura: Estrutura) -> list[str]:
    idx = estrutura.indices.get("referencias")
    if idx is None:
        return []
    return [linha for linha in estrutura.linhas[idx + 1 :] if linha]


# ---------------------------------------------------------------- citações

_RE_PAREN = re.compile(r"\(([^()]{2,240}?)\)")
_RE_ANO = re.compile(r"\b(1[5-9]\d{2}|20\d{2})\b")
_RE_NARRATIVA = re.compile(r"([A-ZÀ-Ý][A-Za-zÀ-ÿ'\-]{2,})\s*\(\s*(1[5-9]\d{2}|20\d{2})\s*\)")


def _sobrenome_do_trecho(trecho: str) -> str | None:
    """Primeiro token 'de nome próprio' do trecho, ignorando partículas."""
    for token in re.findall(r"[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ'\-]+", trecho):
        if token.casefold() in PARTICULAS or len(token) < 3:
            continue
        if token[0].isupper():
            return chave(token)
    return None


def extrair_citacoes(texto: str) -> set[tuple[str, int | None]]:
    """Citações autor-data: (SOBRENOME, ano) e Sobrenome (ano).

    Heurística declarada: só pega o formato autor-data. Manuscrito de História em
    nota de rodapé não é reconhecido — e aí a checagem responde NAO_VERIFICAVEL.
    """
    achados: set[tuple[str, int | None]] = set()
    for grupo in _RE_PAREN.findall(texto):
        anos = [int(a) for a in _RE_ANO.findall(grupo)]
        if not anos:
            continue
        sobrenomes = []
        for segmento in grupo.split(";"):
            antes_da_virgula = segmento.split(",")[0]
            sobrenome = _sobrenome_do_trecho(antes_da_virgula)
            if sobrenome:
                sobrenomes.append(sobrenome)
        for i, sobrenome in enumerate(sobrenomes):
            achados.add((sobrenome, anos[min(i, len(anos) - 1)]))
    for nome, ano in _RE_NARRATIVA.findall(texto):
        if nome.casefold() in PARTICULAS:
            continue
        achados.add((chave(nome), int(ano)))
    return achados


def _sobrenome_da_entrada(entrada: str) -> str | None:
    linha = entrada.strip().lstrip("-•*0123456789. ")
    casamento = re.match(r"^([^,.;]{2,60}),", linha)
    trecho = casamento.group(1) if casamento else linha[:60]
    return _sobrenome_do_trecho(trecho)


def _ano_da_entrada(entrada: str) -> int | None:
    anos = _RE_ANO.findall(entrada)
    return int(anos[-1]) if anos else None


# ---------------------------------------------------------------- checagens


def checar_fonte_familia(doc, norma: Norma, estrito: bool = False) -> Resultado:
    limite = norma.limite("familia_fonte")
    titulo = "família da fonte do corpo"
    if not limite.citavel:
        return Resultado(
            "fonte_familia",
            titulo,
            NAO_VERIFICAVEL,
            "a norma não fixa família tipográfica (limite null em normas.json)",
            motivo_nao_verificavel=MOTIVO_NORMA,
        )
    nome, _, origem = _fonte_normal(doc)
    permitidas = [chave(f) for f in limite.valor]
    runs = _runs_do_corpo(doc)
    efetivos = [(run, _familia_do_run(run, nome)) for run in runs]
    fora = [(run, valor) for run, valor in efetivos if valor and chave(valor) not in permitidas]
    cegos = [run for run, valor in efetivos if valor is None]
    detalhes = {
        "declarada": nome,
        "origem": origem,
        "permitidas": limite.valor,
        "runs": len(runs),
        "runs_com_formato_direto": sum(1 for run in runs if run.font.name),
        "runs_fora_do_padrao": len(fora),
        "runs_sem_familia_legivel": len(cegos),
        "limiar_desvio": LIMIAR_DESVIO,
        "fonte": limite.fonte,
    }
    exigido = f"permitidas: {', '.join(limite.valor)}"
    if nome is None and all(valor is None for _, valor in efetivos):
        return Resultado(
            "fonte_familia",
            titulo,
            NAO_VERIFICAVEL,
            f"família não declarada no .docx ({origem}) e nenhum run traz a fonte "
            f"({exigido}); não dá para comparar",
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )

    declarada = f"{nome} ({origem})" if nome else "não declarada no estilo (só no run)"
    contagem = f"{len(fora)}/{len(runs)} runs fora do padrão"
    if _fracao(len(fora), len(runs)) > LIMIAR_DESVIO:
        familias = sorted({valor for _, valor in fora})
        return Resultado(
            "fonte_familia",
            titulo,
            FAIL,
            f"{declarada}; {contagem} ({', '.join(familias)}); {exigido}",
            detalhes,
        )
    if _fracao(len(cegos), len(runs)) > LIMIAR_DESVIO:
        return Resultado(
            "fonte_familia",
            titulo,
            NAO_VERIFICAVEL,
            f"{declarada}; {len(cegos)}/{len(runs)} runs sem família legível "
            f"(herança de tema) — não verificável",
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )
    return Resultado(
        "fonte_familia",
        titulo,
        PASS,
        f"{declarada}; {contagem}; {exigido}",
        detalhes,
    )


def checar_fonte_tamanho(doc, norma: Norma, estrito: bool = False) -> Resultado:
    limite = norma.limite("tamanho_fonte_pt")
    titulo = "tamanho da fonte do corpo"
    if not limite.citavel:
        return Resultado(
            "fonte_tamanho",
            titulo,
            NAO_VERIFICAVEL,
            "a norma não fixa corpo (limite null em normas.json)",
            motivo_nao_verificavel=MOTIVO_NORMA,
        )
    esperado = float(limite.valor)
    _, tamanho, origem = _fonte_normal(doc)
    runs = _runs_do_corpo(doc)
    efetivos = [(run, _tamanho_do_run(run, tamanho)) for run in runs]
    fora = [(run, valor) for run, valor in efetivos if valor is not None and abs(valor - esperado) > TOL_PT]
    cegos = [run for run, valor in efetivos if valor is None]
    detalhes = {
        "declarado_pt": tamanho,
        "origem": origem,
        "exigido_pt": esperado,
        "runs": len(runs),
        "runs_com_formato_direto": sum(1 for run in runs if run.font.size is not None),
        "runs_fora_do_padrao": len(fora),
        "runs_sem_tamanho_legivel": len(cegos),
        "limiar_desvio": LIMIAR_DESVIO,
        "fonte": limite.fonte,
    }
    if tamanho is None and all(valor is None for _, valor in efetivos):
        return Resultado(
            "fonte_tamanho",
            titulo,
            NAO_VERIFICAVEL,
            f"corpo não declarado no .docx ({origem}) e nenhum run traz tamanho; "
            f"não dá para comparar com {numero(esperado)} pt",
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )

    base = (
        f"{numero(tamanho)} pt ({origem})" if tamanho is not None else "só formatação direta"
    )
    if not runs:
        base = f"{base}, exigido {numero(esperado)} pt; documento sem runs de texto"
        fora_total = 1 if tamanho is not None and abs(tamanho - esperado) > TOL_PT else 0
        return Resultado(
            "fonte_tamanho",
            titulo,
            FAIL if fora_total else PASS,
            base,
            detalhes,
        )
    contagem = f"{len(fora)}/{len(runs)} runs fora do padrão"
    if _fracao(len(fora), len(runs)) > LIMIAR_DESVIO:
        tamanhos = sorted({numero(valor) for _, valor in fora})
        return Resultado(
            "fonte_tamanho",
            titulo,
            FAIL,
            f"{base}, exigido {numero(esperado)} pt; {contagem} ({', '.join(tamanhos)} pt)",
            detalhes,
        )
    if _fracao(len(cegos), len(runs)) > LIMIAR_DESVIO:
        return Resultado(
            "fonte_tamanho",
            titulo,
            NAO_VERIFICAVEL,
            f"{base}, exigido {numero(esperado)} pt; {len(cegos)}/{len(runs)} runs sem "
            "tamanho legível — não verificável",
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )
    return Resultado(
        "fonte_tamanho",
        titulo,
        PASS,
        f"{base}, exigido {numero(esperado)} pt; {contagem}",
        detalhes,
    )


def checar_espacamento(doc, norma: Norma, estrito: bool = False) -> Resultado:
    limite = norma.limite("espacamento_linhas")
    titulo = "espaçamento entre linhas do corpo"
    if not limite.citavel:
        return Resultado(
            "espacamento",
            titulo,
            NAO_VERIFICAVEL,
            "a norma não fixa espaçamento (limite null em normas.json)",
            motivo_nao_verificavel=MOTIVO_NORMA,
        )
    esperado = float(limite.valor)

    padrao, padrao_fixo = _espaco_do_estilo(doc)
    origem_padrao = "estilo Normal"
    if padrao is None and not padrao_fixo:
        padrao, padrao_fixo = _docdefaults_espaco(doc)
        origem_padrao = "docDefaults" if (padrao is not None or padrao_fixo) else "nada declarado"

    valores: list[float] = []
    fixos = 0
    for paragrafo in doc.paragraphs:
        if not paragrafo.text.strip():
            continue
        multiplo, fixo = _espaco_declarado(paragrafo._p.pPr)
        if fixo:
            # Espaçamento fixo (w:lineRule exact/atLeast) NÃO é múltiplo: não se
            # pode cair no padrão do estilo e inventar conformidade.
            fixos += 1
            continue
        efetivo = multiplo if multiplo is not None else padrao
        if efetivo is not None:
            valores.append(round(efetivo, 2))
        elif padrao_fixo:
            fixos += 1

    total = len(valores) + fixos
    nota = f" (padrão: {origem_padrao})" if origem_padrao != "estilo Normal" else ""
    if total == 0:
        return Resultado(
            "espacamento",
            titulo,
            NAO_VERIFICAVEL,
            "nenhum parágrafo/estilo declara espaçamento em múltiplo; o efetivo vem do tema e não é legível",
            {"origem_padrao": origem_padrao, "exigido": esperado, "fonte": limite.fonte},
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )

    contagem = Counter(valores)
    desvios = sum(1 for v in valores if abs(v - esperado) > TOL_ESPACO)
    detalhes = {
        "moda": contagem.most_common(1)[0][0] if valores else None,
        "distribuicao": dict(contagem),
        "paragrafos": len(valores),
        "paragrafos_fora": desvios,
        "paragrafos_com_espaco_fixo": fixos,
        "total_paragrafos": total,
        "origem_padrao": origem_padrao,
        "exigido": esperado,
        "limiar_desvio": LIMIAR_DESVIO,
        "fonte": limite.fonte,
    }

    if not valores:
        return Resultado(
            "espacamento",
            titulo,
            NAO_VERIFICAVEL,
            f"{fixos} parágrafo(s) usam espaçamento exato (w:lineRule exact/atLeast), que "
            f"não é múltiplo comparável com {numero(esperado)}{nota}",
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )

    moda, ocorrencias = contagem.most_common(1)[0]

    sobra_fixo = f"; {fixos} parágrafo(s) com espaçamento exato" if fixos else ""
    if _fracao(desvios, total) > LIMIAR_DESVIO:
        return Resultado(
            "espacamento",
            titulo,
            FAIL,
            f"moda {numero(moda)}, exigido {numero(esperado)}; "
            f"{desvios}/{total} parágrafos fora{sobra_fixo}{nota}",
            detalhes,
        )
    if _fracao(fixos, total) > LIMIAR_DESVIO:
        return Resultado(
            "espacamento",
            titulo,
            NAO_VERIFICAVEL,
            f"{fixos}/{total} parágrafos usam espaçamento exato (w:lineRule exact/atLeast); "
            f"não comparável com {numero(esperado)}{nota}",
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )
    return Resultado(
        "espacamento",
        titulo,
        PASS,
        f"{numero(moda)} em {ocorrencias}/{len(valores)} parágrafos com múltiplo; "
        f"{desvios} fora do padrão{sobra_fixo}{nota}",
        detalhes,
    )


def checar_margens(doc, norma: Norma, estrito: bool = False) -> Resultado:
    limite = norma.limite("margens_cm")
    if not limite.citavel:
        return Resultado(
            "margens",
            "margens da página",
            NAO_VERIFICAVEL,
            "a norma não fixa margens (limite null em normas.json)",
            motivo_nao_verificavel=MOTIVO_NORMA,
        )
    esperado = limite.valor
    secoes = list(doc.sections)
    if not secoes:
        return Resultado(
            "margens",
            "margens da página",
            NAO_VERIFICAVEL,
            "documento sem seção legível",
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )

    problemas: list[str] = []
    indefinidas: list[str] = []
    medidas: dict[str, list[float]] = defaultdict(list)
    for indice, secao in enumerate(secoes, 1):
        for lado, atributo in MARGENS:
            comprimento = getattr(secao, atributo, None)
            if comprimento is None:
                indefinidas.append(f"seção {indice}: {lado} não declarada")
                continue
            cm = round(float(comprimento.cm), 2)
            medidas[lado].append(cm)
            if abs(cm - esperado[lado]) > TOL_CM:
                problemas.append(f"seção {indice}: {lado} {numero(cm)} cm (exigido {numero(esperado[lado])})")

    detalhes = {
        "medidas_cm": {lado: sorted(set(v)) for lado, v in medidas.items()},
        "exigido_cm": esperado,
        "secoes": len(secoes),
        "fonte": limite.fonte,
    }
    resumo_medidas = " · ".join(f"{lado} {numero(sorted(set(v))[0])}" for lado, v in medidas.items() if v)
    exigido_txt = "/".join(numero(esperado[lado]) for lado, _ in MARGENS)
    if problemas:
        return Resultado(
            "margens",
            "margens da página",
            FAIL,
            "; ".join(problemas),
            detalhes,
        )
    if indefinidas:
        return Resultado(
            "margens",
            "margens da página",
            NAO_VERIFICAVEL,
            "; ".join(indefinidas),
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )
    return Resultado(
        "margens",
        "margens da página",
        PASS,
        f"{resumo_medidas} cm (exigido {exigido_txt})",
        detalhes,
    )


def checar_secoes(doc, norma: Norma, estrito: bool = False) -> Resultado:
    limite = norma.limite("secoes_obrigatorias")
    if not limite.citavel:
        return Resultado(
            "secoes",
            "título/abstract/resumo presentes",
            NAO_VERIFICAVEL,
            "a norma não fixa seções obrigatórias (limite null em normas.json)",
            motivo_nao_verificavel=MOTIVO_NORMA,
        )
    estrutura = _estrutura(doc)
    exigidas: list[str] = limite.valor
    encontradas: dict[str, str] = {}
    if estrutura.titulo:
        encontradas["titulo"] = estrutura.titulo
    for nome in ("resumo", "abstract", "referencias"):
        if nome in estrutura.indices:
            encontradas[nome] = estrutura.linhas[estrutura.indices[nome]]
    faltando = [s for s in exigidas if s not in encontradas]

    detalhes = {
        "exigidas": exigidas,
        "encontradas": encontradas,
        "faltando": faltando,
        "marcadores_de_estrutura": sorted(estrutura.indices),
        "fonte": limite.fonte,
    }
    if not estrutura.alguma_estrutura:
        return Resultado(
            "secoes",
            "título/abstract/resumo presentes",
            NAO_VERIFICAVEL,
            "nenhum cabeçalho reconhecível (título/resumo/abstract/referências): heurística não decide",
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )
    if faltando:
        return Resultado(
            "secoes",
            "título/abstract/resumo presentes",
            FAIL,
            f"ausentes: {', '.join(faltando)}; encontrados: {', '.join(sorted(encontradas)) or 'nenhum'}",
            detalhes,
        )
    return Resultado(
        "secoes",
        "título/abstract/resumo presentes",
        PASS,
        f"encontrados: {', '.join(sorted(encontradas))}",
        detalhes,
    )


def checar_palavras(doc, norma: Norma, estrito: bool = False) -> Resultado:
    limite = norma.limite("palavras")
    estrutura = _estrutura(doc)
    por_secao = {nome: contar_palavras(texto) for nome, texto in estrutura.trechos if texto}
    total = sum(contar_palavras(linha) for linha in estrutura.linhas)
    excluir = norma.limite("contagem_exclui_referencias")
    exclui_referencias = bool(excluir.valor) if excluir.citavel else False
    palavras_referencias = sum(
        contar_palavras(texto) for nome, texto in estrutura.trechos if nome.startswith("referencias")
    )
    contado = total - palavras_referencias if exclui_referencias else total

    detalhes = {
        "total": total,
        "referencias": palavras_referencias,
        "contado": contado,
        "contagem_exclui_referencias": exclui_referencias,
        "por_secao": por_secao,
        "fonte": limite.fonte if limite.citavel else None,
    }
    if exclui_referencias:
        nota = f" (corpo; total {total} − {palavras_referencias} de referências)"
    else:
        nota = f" (texto completo, inclui {palavras_referencias} palavras de referências)"
    if not limite.citavel:
        return Resultado(
            "palavras",
            "contagem de palavras",
            NAO_VERIFICAVEL,
            f"{contado} palavras{nota}; a norma não fixa faixa",
            detalhes,
            motivo_nao_verificavel=MOTIVO_NORMA,
        )
    minimo, maximo = limite.valor["min"], limite.valor["max"]
    detalhes.update({"min": minimo, "max": maximo})
    if minimo <= contado <= maximo:
        return Resultado(
            "palavras",
            "contagem de palavras",
            PASS,
            f"{contado} palavras{nota}; faixa {minimo}–{maximo}",
            detalhes,
        )
    return Resultado(
        "palavras",
        "contagem de palavras",
        FAIL,
        f"{contado} palavras{nota}; fora da faixa {minimo}–{maximo}",
        detalhes,
    )


def checar_autor(doc, norma: Norma, estrito: bool = False) -> Resultado:
    limite = norma.limite("revisao_cega")
    autor = (doc.core_properties.author or "").strip()
    editor = (doc.core_properties.last_modified_by or "").strip()
    detalhes = {
        "author": achatar(autor, 80) or "(vazio)",
        "last_modified_by": achatar(editor, 80) or "(vazio)",
        "fonte": limite.fonte if limite.citavel else None,
    }
    if not limite.citavel:
        return Resultado(
            "autor_metadado",
            'metadado de autor ("cego"?)',
            NAO_VERIFICAVEL,
            "a norma escolhida não declara revisão cega (limite null em normas.json)",
            detalhes,
            motivo_nao_verificavel=MOTIVO_NORMA,
        )
    if not limite.valor:
        return Resultado(
            "autor_metadado",
            'metadado de autor ("cego"?)',
            NAO_VERIFICAVEL,
            "a norma declara revisão não cega; metadado não é exigência",
            detalhes,
            motivo_nao_verificavel=MOTIVO_NORMA,
        )

    def anonimo(valor: str) -> bool:
        base = chave(valor)
        return not base or base in ANONIMOS or "anon" in base or "blind" in base

    expostos = [
        f"{campo}={valor!r}"
        for campo, valor in (("author", autor), ("last_modified_by", editor))
        if not anonimo(valor)
    ]
    if expostos:
        return Resultado(
            "autor_metadado",
            'metadado de autor ("cego"?)',
            FAIL,
            "revisão cega exigida e o metadado identifica o autor: " + "; ".join(expostos),
            detalhes,
        )
    return Resultado(
        "autor_metadado",
        'metadado de autor ("cego"?)',
        PASS,
        f"metadado sem identificação (author={detalhes['author']}, last_modified_by={detalhes['last_modified_by']})",
        detalhes,
    )


def checar_referencias(doc, norma: Norma, estrito: bool = False) -> Resultado:
    limite = norma.limite("secoes_obrigatorias")
    # A coerência citação × bibliografia só é cobrada quando a norma fixa que
    # existe bloco de referências. Sem isso, o veredito seria opinião do script.
    if not limite.citavel or "referencias" not in (limite.valor or []):
        return Resultado(
            "referencias",
            "referências: entradas × citações",
            NAO_VERIFICAVEL,
            "a norma não exige bloco de referências (limite null): coerência bibliográfica não verificada",
            {"fonte": limite.fonte if limite.citavel else None},
            motivo_nao_verificavel=MOTIVO_NORMA,
        )
    estrutura = _estrutura(doc)
    entradas = _bloco_referencias(estrutura)
    citacoes = extrair_citacoes(estrutura.corpo)

    mapa: dict[str, set[int]] = defaultdict(set)
    sem_sobrenome: list[str] = []
    for entrada in entradas:
        sobrenome = _sobrenome_da_entrada(entrada)
        if not sobrenome:
            sem_sobrenome.append(achatar(entrada, 60))
            continue
        ano = _ano_da_entrada(entrada)
        if ano:
            mapa[sobrenome].add(ano)
        else:
            mapa[sobrenome]  # entrada sem ano conta como existente

    citados = {sobrenome for sobrenome, _ in citacoes}
    sem_entrada = sorted({s for s, _ in citacoes if s not in mapa})
    anos_divergentes = sorted(
        {(s, a) for s, a in citacoes if a and s in mapa and mapa[s] and a not in mapa[s]}
    )
    nao_citadas = sorted(set(mapa) - citados)

    detalhes = {
        "entradas": len(entradas),
        "citacoes": len(citacoes),
        "sobrenomes_citados": sorted(citados),
        "sobrenomes_na_bibliografia": sorted(mapa),
        "citacoes_sem_entrada": sem_entrada,
        "anos_divergentes": [f"{s}, {a}" for s, a in anos_divergentes],
        "entradas_nao_citadas": nao_citadas,
        "entradas_sem_sobrenome": sem_sobrenome,
        "fonte": limite.fonte if limite.citavel else None,
    }

    if not citacoes and not entradas:
        # A norma exige referências e o documento não traz nem entrada nem citação:
        # não é "não verificável", é ausência do elemento exigido.
        detalhes["provavel_fragmento"] = True
        return Resultado(
            "referencias",
            "referências: entradas × citações",
            FAIL,
            "a norma exige bloco de referências e o documento não tem nenhuma entrada "
            "nem citação autor-data (provável fragmento, não artigo completo)",
            detalhes,
        )
    if citacoes and not entradas:
        return Resultado(
            "referencias",
            "referências: entradas × citações",
            FAIL,
            f"{len(citacoes)} citações autor-data e nenhum bloco de referências",
            detalhes,
        )
    if not citacoes and entradas:
        return Resultado(
            "referencias",
            "referências: entradas × citações",
            NAO_VERIFICAVEL,
            f"{len(entradas)} entradas, nenhuma citação autor-data no corpo "
            "(pode usar nota de rodapé: não verificável)",
            detalhes,
            motivo_nao_verificavel=MOTIVO_DOCUMENTO,
        )
    if sem_entrada or anos_divergentes:
        partes = []
        if sem_entrada:
            partes.append(f"{len(sem_entrada)} citação(ões) sem entrada ({', '.join(sem_entrada)})")
        if anos_divergentes:
            partes.append(f"{len(anos_divergentes)} ano(s) divergente(s)")
        if nao_citadas:
            partes.append(f"{len(nao_citadas)} entrada(s) não citada(s)")
        return Resultado(
            "referencias",
            "referências: entradas × citações",
            FAIL,
            f"{len(entradas)} entrada(s) × {len(citacoes)} citação(ões): " + "; ".join(partes),
            detalhes,
        )
    if nao_citadas and estrito:
        return Resultado(
            "referencias",
            "referências: entradas × citações",
            FAIL,
            f"{len(entradas)} entradas × {len(citacoes)} citações; "
            f"{len(nao_citadas)} entradas não citadas (estrito): {', '.join(nao_citadas)}",
            detalhes,
        )
    nota = f"; {len(nao_citadas)} entradas não citadas" if nao_citadas else ""
    return Resultado(
        "referencias",
        "referências: entradas × citações",
        PASS,
        f"{len(entradas)} entradas × {len(citacoes)} citações casam{nota}",
        detalhes,
    )


CHECAGENS: tuple[tuple[str, str, Callable[..., Resultado]], ...] = (
    ("fonte_familia", "família da fonte do corpo", checar_fonte_familia),
    ("fonte_tamanho", "tamanho da fonte do corpo", checar_fonte_tamanho),
    ("espacamento", "espaçamento entre linhas do corpo", checar_espacamento),
    ("margens", "margens da página", checar_margens),
    ("secoes", "título/abstract/resumo presentes", checar_secoes),
    ("palavras", "contagem de palavras", checar_palavras),
    ("autor_metadado", 'metadado de autor ("cego"?)', checar_autor),
    ("referencias", "referências: entradas × citações", checar_referencias),
)

# O id da checagem não é o nome do limite em normas.json: este mapa é a ponte.
LIMITE_DA_CHECAGEM: dict[str, str] = {
    "fonte_familia": "familia_fonte",
    "fonte_tamanho": "tamanho_fonte_pt",
    "espacamento": "espacamento_linhas",
    "margens": "margens_cm",
    "secoes": "secoes_obrigatorias",
    "palavras": "palavras",
    "autor_metadado": "revisao_cega",
    "referencias": "secoes_obrigatorias",
}


def auditar(caminho: str | Path, norma: Norma, estrito: bool = False) -> list[Resultado]:
    """Abre o .docx e roda todas as checagens. Erro de leitura sobe para o CLI."""
    documento = Document(str(caminho))
    return [funcao(documento, norma, estrito) for _, _, funcao in CHECAGENS]


def texto_do_docx(caminho: str | Path) -> str:
    """Texto plano do .docx (usado pelo check_bib para cruzar citações)."""
    documento = Document(str(caminho))
    return "\n".join(p.text for p in documento.paragraphs)
