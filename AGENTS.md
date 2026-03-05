# Repository Guidelines

## Project Structure & Module Organization
- `app.py`: Streamlit entrypoint and UI orchestration.
- `src/domain/`: core models (`SheetConfig`, mappings, enums).
- `src/application/`: transformation/business rules (Excel row -> Azure CSV task).
- `src/infrastructure/`: I/O adapters (Excel reading, CSV writing, profile persistence).
- `tests/`: unit tests for business logic (`pytest`).
- `config/`: static configuration (for example task type catalog).
- `profiles/`: user-saved mapping profiles (runtime data, not source code).

Keep business rules in `src/application` and avoid placing logic directly in `app.py`.

## Build, Test, and Development Commands
- `uv sync`: install dependencies from `pyproject.toml`/`uv.lock`.
- `uv run streamlit run app.py`: run the app locally.
- `uv run pytest -q`: run unit tests.
- `uv run python -c "import app; print('ok')"`: quick import sanity check.

Run tests before creating a commit.

## Coding Style & Naming Conventions
- Python 3.12+, 4-space indentation, UTF-8 files.
- Use `snake_case` for functions/variables, `PascalCase` for classes, `UPPER_SNAKE_CASE` for constants.
- Prefer small, single-purpose functions with explicit inputs/outputs.
- Keep type hints in public/internal function signatures where practical.
- Preserve layering: UI in `app.py`, logic in `src/application`, persistence/IO in `src/infrastructure`.

No formatter/linter is currently enforced in CI; keep style consistent with existing code.

## Testing Guidelines
- Framework: `pytest`.
- Test files: `tests/test_*.py`; test names: `test_<behavior>()`.
- Add tests for every rule change in transformation/filtering logic.
- Cover edge cases (empty titles, summary rows, missing mappings, numeric parsing).

## Commit & Pull Request Guidelines
- Follow Conventional Commits (seen in history): e.g., `feat: ...`, `docs: ...`, `fix: ...`.
- Keep commits focused (one logical change per commit when possible).
- PRs should include:
  - short problem/solution summary,
  - test evidence (`uv run pytest -q` output),
  - UI screenshots/GIFs for Streamlit-visible changes,
  - linked issue/task if available.

## Security & Configuration Tips
- Never commit secrets, tokens, or private data files.
- Treat files in `profiles/` as environment/user data; avoid committing sensitive content.
- Validate external Excel inputs defensively and fail with clear user-facing warnings.
