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
| 0 | nenhum `FAIL` e nada ficou sem verificar (só `NAO_VERIFICAVEL` **por falta de limite na norma**, que é a resposta correta) |
| 1 | pelo menos um `FAIL` |
| 2 | erro de uso: `.docx` inexistente/corrompido, norma desconhecida, `normas.json` inválido |
| 3 | alguma checagem não pôde verificar **o documento** (nada declarado, espaçamento exato, heurística cega) — ou qualquer `NAO_VERIFICAVEL` com `--estrito` |

`NAO_VERIFICAVEL` sempre carrega `motivo_nao_verificavel ∈ {norma, documento}`:

* **`norma`** — a norma não fixa o parâmetro (`null` em `normas.json`): exit `0`, é a resposta certa;
* **`documento`** — a norma fixa, mas o `.docx` não permite verificar: exit `3` **por default**.
  Não verificado não é aprovado; o código de saída não trata os dois casos como iguais.

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

Toda checagem de forma (`fonte_familia`, `fonte_tamanho`, `espacamento`) compara o
**efetivo**: a formatação direta do run/parágrafo vence o estilo `Normal`. Se mais de
`LIMIAR_DESVIO` (**10 %**, em `docx_check.py`) do corpo divergir do padrão declarado, o
status é `FAIL` com a contagem no veredito — `PASS` não convive com desvio material.

| checagem | como decide |
|---|---|
| `fonte_familia` | família efetiva de cada run do corpo (direta ou herdada do `Normal`/`docDefaults`), comparada sem acento/caixa; > 10 % fora das permitidas → `FAIL` |
| `fonte_tamanho` | tamanho efetivo de cada run (direto ou herdado); > 10 % fora do corpo exigido → `FAIL` com "N/M runs fora do padrão" |
| `espacamento` | `w:spacing` de cada parágrafo (herança de estilo/`docDefaults`); > 10 % dos parágrafos fora do múltiplo exigido → `FAIL`. Espaçamento **exato** (`w:lineRule="exact"`/`atLeast`) não é múltiplo comparável → `NAO_VERIFICAVEL` (motivo `documento`), nunca `PASS`. Sem `w:spacing` em nível nenhum, o efetivo é o default da aplicação (espaço simples): atende → `NAO_VERIFICAVEL`, não atende → `FAIL` |
| `margens` | `section.top/bottom/left/right_margin` em cm, tolerância 0,08 cm (arredondamento twips). Sem `w:pgMar` (nem `w:sectPr`), o efetivo é o default da aplicação (2,54 cm nos quatro lados): atende → `NAO_VERIFICAVEL`, não atende → `FAIL` com o motivo escrito |
| `secoes` | cabeçalhos reconhecidos por regex (Resumo/Abstract/Referências/seções numeradas); título = primeira linha curta sem ponto final. Elemento escrito como **rótulo no corpo** ("Resumo. O artigo…", "Palavras-chave: …") conta como **presente sem título de seção** — `FAIL` de forma, nunca "ausente" |
| `palavras` | contagem por token, total e por seção; respeita `contagem_exclui_referencias` |
| `autor_metadado` | `core_properties.author` e `last_modified_by` contra marcadores de anonimização |
| `referencias` | extrai chamadas autor-data (`(SILVA; SOUZA, 2024)`, `Silva e Souza (2024)`, `Silva et al. (2024)`) e casa **qualquer** sobrenome da chamada com **qualquer** autor das entradas; "não citada" só quando nenhuma chamada casar. Sem **nenhuma** entrada nem citação, a norma que exige referências reprova (`FAIL`) |

Quatro decisões de `referencias` que evitam `FAIL` falso (cada uma com fixture própria em
`tests/`, card `t_7c691dbc`):

* **ano da entrada é conjunto, não o último do texto.** A cauda de acesso (`Disponível em:`/
  `Acesso em:`) e os identificadores (DOI, ISBN, ISSN, e-locator) saem antes da varredura —
  em referência online o último ano do texto é o do acesso, e `10.14763/2026.1.2074` tem
  `2074` no sufixo. Mais de um ano na mesma entrada (edição e reimpressão) é aceito.
* ***apud* não exige entrada da fonte não consultada.** NBR 10520 (citação de citação):
  só a obra consultada entra nas referências. Em `(RID, 2020 apud BOSWINKEL et al., 2022)`
  a única entrada exigida é a de BOSWINKEL, e o ano que vale é o dela.
* **menção a título de obra em prosa não é chamada.** `Em Propaganda (1928), Lippmann…`
  não pede entrada de "propaganda": chamada autor-data é sobrenome em caixa alta (NBR
  10520) ou nome próprio **não** precedido de partícula que introduz título (`em`, `no`,
  `na`, `nos`, `nas`). As menções reconhecidas aparecem em `mencionadas_como_titulo`, no
  JSON do laudo — o veredito diz que viu, não que ignorou.
* **o mesmo sobrenome pode ter várias entradas.** O mapa vai de sobrenome para o
  **conjunto** de índices: a chamada marca todas as entradas cujo conjunto de anos contém
  o ano citado (e todas as do sobrenome quando o ano não desambigua). Com uma única
  entrada por sobrenome, a segunda obra do mesmo autor ficava eternamente "não citada".

### O que **não** é coberto

* **Formato**: só `.docx`. Não lê `.doc`, `.odt`, `.pdf`, `.rtf`.
* **Estilo de citação**: não valida NBR 10520/6023 (recuo de citação longa, ordem dos
  elementos da referência, itálico, pontuação). `referencias` compara *quem* foi citado,
  não *como*.
* **Citações em nota de rodapé / sistema numérico**: não lidas. Manuscrito que tem
  referências mas nenhuma citação autor-data cai em `NAO_VERIFICAVEL` (motivo `documento`)
  na checagem `referencias` — de propósito; sem referências **e** sem citações, é `FAIL`.
* **Chamada em caixa mista precedida de `em/no/na`**: lida como título de obra em prosa e
  **não** reprova — `em Klein (2020)` sem entrada passa. É o preço de não acusar
  `Em Propaganda (1928)`; chamada em caixa alta (`(KLEIN, 2020)`, a forma da NBR 10520) e
  chamada parentética em caixa mista continuam reprovando.
* **Herança de tema**: `python-docx` não resolve `theme1.xml`. Fonte definida só no tema
  aparece como `nada declarado` → `NAO_VERIFICAVEL` (motivo `documento`, exit 3). Margem e
  espaçamento **não** são tema: quando o arquivo não declara nenhum dos dois, o efetivo é
  o default da aplicação (2,54 cm; espaço simples) e a checagem reprova se ele não atende.
* **Espaçamento misto**: se uma minoria de parágrafos (< 10 %) usa espaçamento exato, o
  veredito `PASS` aparece com a contagem no rodapé; o texto exato em si não é medido.
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
  * `direto_10pt.docx` — `Normal` 12 pt com o corpo inteiro em 10 pt Arial por formatação
    direta: `PASS` por estatística não é aceitável, o veredito é `FAIL`;
  * `espaco_exato.docx` — `Normal` 1,5 com todo parágrafo em espaçamento exato de 18 pt:
    `NAO_VERIFICAVEL` (não é múltiplo comparável), nunca `PASS`;
  * `fragmento.docx` — trecho sem cabeçalho, sem citação e **sem** declaração de fonte/
    espaçamento (prova do caminho `NAO_VERIFICAVEL` de fonte e do `FAIL` por default de
    espaçamento);
  * `resumo_rotulo.docx` — resumo/abstract como **rótulo em negrito** no corpo, sem título
    de seção: o elemento existe, `secoes` reprova a forma e não diz "ausente";
  * `citacao_dois_autores.docx` — `Silva e Souza (2024)`, `(SILVA; SOUZA, 2024)` e
    `Souza (2024)` contra uma entrada de dois autores: nada de "citação sem entrada" nem de
    "não citada";
  * `sem_setup.docx` — `.docx` sem `w:sectPr` e sem `w:spacing`: `margens` e `espacamento`
    dizem qual é o default da aplicação e reprovam o que não atende.
* **versionadas**: `bib_ok.bib` + `texto_ok.tex`, `bib_com_problemas.bib` +
  `texto_problemas.tex` (duplicata por DOI, duplicata por título normalizado, `@book` sem
  `publisher`, ano 2099, entrada sem ano, sem DOI/URL, chave não citada, citação sem entrada).

A suíte (`pytest -q`, 127 testes) exige os **três estados para cada uma das 8 checagens**:
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
