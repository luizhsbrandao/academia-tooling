# academia-tooling

Verificador de conformidade de manuscrito `.docx` (ABNT / REGN / RBCS) e auditor de
bibliografia `.bib`, para pacotes de submissão da frente `academia`.

O princípio do repo é **não chutar**: cada limite numérico vem de `normas.json` com a
fonte citada; quando a norma não fixa o parâmetro (ou o `.docx` não declara o atributo),
a resposta é `NAO_VERIFICAVEL` — nunca `PASS`.

**Nenhum arquivo de vault é lido por este repo.** As fixtures `.docx` são sintéticas,
geradas por `academia_tooling/fixtures.py`.

## Instalação

```bash
make venv     # uv venv --python 3.13 .venv + uv pip install -e ".[dev]"
make fixtures # gera tests/fixtures/gerados/*.docx
```

Dependência de execução: `python-docx`. Testes: `pytest`.

## Uso

```bash
# conformidade do manuscrito
./check_docx.py --arquivo manuscrito.docx --norma abnt
./check_docx.py --arquivo manuscrito.docx --norma rbcs --json
./check_docx.py --arquivo manuscrito.docx --norma regn --normas meu_edital.json

# bibliografia (+ cruzamento com o texto, opcional)
./check_bib.py --bib referencias.bib
./check_bib.py --bib referencias.bib --texto artigo.tex
./check_bib.py --bib referencias.bib --docx manuscrito.docx --estrito
```

Tudo de uma vez (`pytest` + laudo dos fixtures de aceite):

```bash
make check
```

### Códigos de saída — `check_docx.py`

| código | significado |
|---|---|
| 0 | nenhum `FAIL` (pode haver `NAO_VERIFICAVEL`) |
| 1 | pelo menos um `FAIL` |
| 2 | erro de uso: `.docx` inexistente/corrompido, norma desconhecida, `normas.json` inválido |
| 3 | `--estrito` e alguma checagem `NAO_VERIFICAVEL` (úsil em CI: "não verificado" reprova) |

### Códigos de saída — `check_bib.py`

| código | significado |
|---|---|
| 0 | nenhum `ERRO` (pode haver `AVISO`) |
| 1 | pelo menos um `ERRO` (ou `AVISO` com `--estrito`) |
| 2 | erro de uso: arquivo inexistente, `.docx` ilegível |

## `normas.json` — o contrato

Cada limite é `null` **ou** `{"valor": ..., "fonte": "citação não vazia"}`. Um limite com
`valor` e sem `fonte` é **rejeitado na carga** (`NormaInvalida`): é assim que "nunca
chutar" deixa de ser recomendação e vira código.

| limite | usado por | forma |
|---|---|---|
| `margens_cm` | `margens` | `{superior, inferior, esquerda, direita}` em cm |
| `tamanho_fonte_pt` | `fonte_tamanho` | número (pt) |
| `familia_fonte` | `fonte_familia` | string ou lista de famílias aceitas |
| `espacamento_linhas` | `espacamento` | múltiplo (1.5 = 1,5 linha) |
| `palavras` | `palavras` | `{min, max}` — exige `contagem_exclui_referencias` |
| `contagem_exclui_referencias` | `palavras` | `true`/`false` |
| `secoes_obrigatorias` | `secoes` e `referencias` | subconjunto de `titulo, resumo, abstract, referencias` |
| `revisao_cega` | `autor_metadado` | `true`/`false` |

Normas embarcadas:

* **`abnt`** — NBR 14724:2011: margens 3/2 cm, corpo 12 pt, espaço 1,5, título/resumo/
  referências obrigatórios. Família tipográfica e limite de palavras **não** são fixados
  pela norma → `null` → `NAO_VERIFICAVEL`. Revisão cega: `null`.
* **`rbcs`** — só a faixa de 7.000–9.000 palavras (corpo, sem referências) está
  registrada. Regra tipográfica não é citável a partir daqui → `null`.
* **`regn`** — nenhum limite citável → **todas** as checagens respondem
  `NAO_VERIFICAVEL`, por desenho. Preencha com o edital da chamada
  (`normas.exemplo.json` é o modelo) e rode com `--normas`.

> `normas.json` é registro, não jurisprudência: confira a edição vigente da norma e o
> edital da chamada antes de tratar o laudo como definitivo.

## O que cada checagem faz

| checagem | como decide |
|---|---|
| `fonte_familia` | família do estilo `Normal` (fallback `docDefaults`), comparada sem acento/caixa |
| `fonte_tamanho` | tamanho do `Normal`; reporta também runs com `size` explícito fora do padrão |
| `espacamento` | `paragraph_format.line_spacing` dos parágrafos, com herança de estilo/`docDefaults`; decide pela **moda** e conta os fora |
| `margens` | `section.top/bottom/left/right_margin` em cm, tolerância 0,08 cm (arredondamento twips) |
| `secoes` | cabeçalhos reconhecidos por regex (Resumo/Abstract/Referências/seções numeradas); título = primeira linha curta sem ponto final |
| `palavras` | contagem por token, total e por seção; respeita `contagem_exclui_referencias` |
| `autor_metadado` | `core_properties.author` e `last_modified_by` contra marcadores de anonimização |
| `referencias` | extrai `(SOBRENOME, ano)` e `Sobrenome (ano)` do corpo e casa com os sobrenomes/anos do bloco final |

### O que **não** é coberto

* **Formato**: só `.docx`. Não lê `.doc`, `.odt`, `.pdf`, `.rtf`.
* **Estilo de citação**: não valida NBR 10520/6023 (recuo de citação longa, ordem dos
  elementos da referência, itálico, pontuação). `referencias` compara *quem* foi citado,
  não *como*.
* **Citações em nota de rodapé / sistema numérico**: não lidas. Manuscrito de História
  com notas cai em `NAO_VERIFICAVEL` na checagem `referencias` — de propósito.
* **Herança de tema**: `python-docx` não resolve `theme1.xml`. Fonte/espaçamento definidos
  só no tema aparecem como `nada declarado` → `NAO_VERIFICAVEL`.
* **Fonte por run / por parágrafo**: `fonte_familia` olha o estilo `Normal`; overrides
  individuais não são comparados (só `fonte_tamanho` os conta).
* **Texto fora de `doc.paragraphs`**: tabelas, cabeçalho/rodapé, notas de rodapé, caixas de
  texto e campos não entram na contagem de palavras nem na extração de citações.
* **Conteúdo**: nada de qualidade argumentativa, plágio, ortografia, adequação ao escopo.
* **`.bib`**: não valida se o DOI existe (sem rede), não checa páginas/edição/volume como
  exigência de estilo, não expande macros LaTeX complexas, e a duplicata por título é
  heurística (título normalizado: sem acento, sem pontuação, sem comando LaTeX).
* **Autoria "cega"**: só `author` e `last_modified_by`. Comentários, propriedades
  personalizadas, histórico de revisões e nomes no corpo do texto não são verificados.
* **Nada de indexação**: o repo não lê vault, não copia manuscrito e não guarda corpus.

## Fixtures e prova

`tests/fixtures/`:

* **geradas** (`gerados/`, git-ignored) por `academia_tooling/fixtures.py`:
  * `conforme.docx` — Times 12, 1,5, margens 3/2/3/2, Resumo/Abstract/Referências, citações
    casando com a bibliografia, metadado de autor vazio;
  * `fora_de_norma.docx` — Arial 10, espaço simples, margem 1 cm, sem Resumo, autor exposto,
    citação (`CHALMERS, 1993`) sem entrada;
  * `fragmento.docx` — trecho sem cabeçalho, sem citação e **sem** declaração de fonte/
    espaçamento (prova do caminho `NAO_VERIFICAVEL`).
* **versionadas**: `bib_ok.bib` + `texto_ok.tex`, `bib_com_problemas.bib` +
  `texto_problemas.tex` (duplicata por DOI, duplicata por título normalizado, `@book` sem
  `publisher`, ano 2099, entrada sem ano, sem DOI/URL, chave não citada, citação sem entrada).

A suíte (`pytest -q`, 94 testes) exige os **três estados para cada uma das 8 checagens**:
`test_toda_checagem_tem_os_tres_estados` reprova se alguém adicionar checagem sem caso
negativo — checagem que nunca falha não é checagem.

## Estrutura

```
check_docx.py, check_bib.py      CLIs finos (só chamam o pacote)
normas.json                      limites citáveis por norma
normas.exemplo.json              modelo para edital de chamada
academia_tooling/
  laudo.py                       Resultado/PASS-FAIL-NAO_VERIFICAVEL, tabela, exit codes
  normas.py                      carga e validação do JSON (limite sem fonte é rejeitado)
  docx_check.py                  as 8 checagens sobre python-docx
  bib_parse.py                   parser BibTeX tolerante (sem dependência externa)
  bib_check.py                   achados ERRO/AVISO + cruzamento citação × bibliografia
  fixtures.py                    gerador dos .docx sintéticos
tests/                           pytest (checagens, normas, parser, CLIs)
```
