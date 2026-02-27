# Project Instructions

## Commit Rules

- **No attribution in commit messages.** Do not include `Co-Authored-By`, `Signed-off-by`, or any AI/tool attribution lines.
- Write concise commit messages focused on the "why", not the "what".
- Follow the existing style: imperative mood, no prefix convention required (see git log for examples like `Fix ...`, `Add ...`, `Enhance ...`, `Bump version ...`).

## Linting & Security

```bash
uv run ruff check .
uv run mypy script.py
uv run bandit -c pyproject.toml *.py   # -c flag required; bandit does not auto-detect pyproject.toml
```

## Versioning

Version numbers must always follow the format `YYYY.MMDD` (e.g. `2026.0227` for February 27, 2026).

When asked to "bump the version" or "increase the version", always update the version in **both** files:
1. `pyproject.toml` — `version = "..."` field
2. `script.py` — `__version__ = "..."` at the top of the file
