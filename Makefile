PY  := .venv/bin/python
FIX := tests/fixtures/gerados

.PHONY: venv fixtures test laudo check limpo

venv:
	uv venv --python 3.13 .venv
	uv pip install --python $(PY) -e ".[dev]"

fixtures:
	$(PY) -m academia_tooling.fixtures $(FIX)

test: fixtures
	$(PY) -m pytest -q

# Laudo dos fixtures de aceite: um conforme, um fora de norma, um fragmento, o caso
# de formatação direta (10 pt em todos os runs → FAIL), o espaçamento exato, os três
# padrões medidos em .docx real no t_516daa8a (resumo como rótulo, chamada de dois
# autores, .docx sem setup de página), as seis causas de FAIL falso em `referencias`
# medidas no t_7c691dbc (data de acesso, DOI, apud, título em prosa, autor repetido,
# e a trava do verdadeiro positivo) e o .bib.
laudo: fixtures
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/conforme.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/fora_de_norma.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/fragmento.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/direto_10pt.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/espaco_exato.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/resumo_rotulo.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/citacao_dois_autores.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/sem_setup.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/referencias_acesso.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/referencias_doi.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/referencias_apud.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/referencias_titulo_prosa.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/referencias_autor_repetido.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_docx.py --arquivo $(FIX)/referencias_sem_entrada.docx --norma abnt; echo "exit $$?"
	@echo
	$(PY) check_bib.py --bib tests/fixtures/bib_com_problemas.bib --texto tests/fixtures/texto_problemas.tex; echo "exit $$?"

check: test laudo

limpo:
	rm -rf $(FIX) .pytest_cache
