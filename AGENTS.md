# AGENTS.md — academia-tooling

Regras para qualquer agente (Claude Code, Antigravity, OpenCode) que trabalhe neste repositório.
Vale junto com `~/.agents/STATION-RULES.md`; onde as duas falarem, vale a mais restritiva.

## O que é

Verificador de conformidade de manuscrito `.docx` (ABNT NBR 14724 e conjuntos de regras de
periódico) e auditor de bibliografia `.bib`. Pacote Python MIT, público, uma dependência de runtime
(`python-docx`) e suíte de testes com pytest.

A regra de desenho é **não adivinhar**: todo limite numérico vem de `normas.json` ao lado de uma
fonte citável. Quando a norma não fixa um parâmetro — ou o documento não o declara — o veredito é
`NAO_VERIFICAVEL`, nunca `PASS`. `normas.json` rejeita limite sem `fonte` no carregamento
(`NormaInvalida`), e os códigos de saída separam os dois silêncios: `0` = conforme ou "a norma não
diz"; `3` = "a norma diz, mas este documento não permite verificar". Não afrouxe essa recusa para
fazer teste passar: é o produto.

## Comandos

```bash
.venv/bin/python -m pytest -q                                          # verificação canônica
.venv/bin/pip install -e ".[dev]"                                      # instala se .venv não existir
.venv/bin/python -m academia_tooling.fixtures tests/fixtures/gerados    # gera os .docx de teste
```

`make test` é atalho (`fixtures` + `pytest`); quem decide é o `pytest`. Não há lint configurado no
`pyproject.toml` — não invente regra que não existe. Não rode `npm`, `cargo` nem instalador de
terceiro.

## Dados

- Os `.docx` de teste são **gerados** (`academia_tooling/fixtures.py`), nunca copiados de vault.
- Nenhum manuscrito real, nome de autor ou dado de estudante entra neste repositório — nem em
  fixture, nem em nome de arquivo, nem em docstring de teste. Achado de conformidade sobre texto real
  pertence ao perfil dono do vault, não a este repo.
- Segredo não se lê (`.env`, `*.pem`, `*.key`, `.ssh`, `.gnupg`) e não se escreve em código. O
  `pre-commit` roda `gitleaks` com `.gitleaks.toml` explícito e **falha fechado**.

## Regras

- **Não edite** `*.egg-info/`, `dist/`, `build/` nem `tests/fixtures/gerados/` — são saída de build.
- **Não mude a superfície pública** (`check-docx`/`check-bib`, nomes de veredito `PASS`/`FAIL`/
  `NAO_VERIFICAVEL`, códigos de saída, chaves do `normas.json`) sem autorização explícita: há CI
  publicando e consumidor fora do repo.
- Teste novo entra só se provar comportamento — não teste que espelha a implementação.
- Diff pequeno, um assunto por commit. Mensagem em português, imperativa, com o porquê.
- **Nunca** `git push`, `git add -A`, `git reset --hard`, `git clean`. A árvore fica pronta; commit e
  push são do operador.
- **Nunca** `--dangerously-skip-permissions`, nunca `--auto`.

## Verificação

Quem roda teste é o operador ou o wrapper — não você. Se você tem shell, use **apenas**
`.venv/bin/python -m pytest`; qualquer outro comando é motivo de recusa. Relate o que verificou de
fato, com o comando e a saída; "deve funcionar" não é resultado.
