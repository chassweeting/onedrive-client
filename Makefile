.PHONY: install lint fix format setup

install:
	poetry install

lint:
	poetry run ruff check .

fix:
	poetry run ruff check . --fix

format:
	poetry run ruff format .

setup:
	./scripts/setup.sh
