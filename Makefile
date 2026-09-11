PYTHON ?= $(if $(wildcard ../.venv/bin/python),../.venv/bin/python,python3)
SEED ?= 20260911

.PHONY: solve validate test sensitivity paper all

solve:
	$(PYTHON) solve_c.py --seed $(SEED)

validate:
	$(PYTHON) validate_results.py
	$(PYTHON) scripts/validate_input_xml.py

test:
	$(PYTHON) -m pytest -q

sensitivity:
	$(PYTHON) scripts/run_sensitivity.py

paper:
	$(PYTHON) scripts/build_updated_docx.py

all: solve validate test sensitivity paper
