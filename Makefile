PYTHON ?= $(if $(wildcard ../.venv/bin/python),../.venv/bin/python,python3)
SEED ?= 20260911

.PHONY: solve validate

solve:
	$(PYTHON) solve_c.py --seed $(SEED)

validate:
	$(PYTHON) validate_results.py
	$(PYTHON) scripts/validate_input_xml.py
