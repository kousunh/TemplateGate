# Contributing

Thanks for your interest in TemplateGate. Issues and pull requests are welcome.

## Development setup

```bash
git clone https://github.com/kousunh/TemplateGate.git
cd TemplateGate
python -m venv .venv
source .venv/bin/activate    # Windows: .venv\Scripts\activate
python -m pip install -e ".[dev]"
python -m pytest -q
```

Test fixtures are synthetic and generated on the fly by `fixtures/generate.py`;
they contain fictional data only. Run it directly to inspect the documents the
tests use:

```bash
python fixtures/generate.py    # writes to fixtures/generated/ (git-ignored)
```

## Pull requests

- Add a test for any behaviour change. CI runs the suite on Linux and Windows
  across Python 3.10–3.13, and all of them must pass.
- Keep the public surface (`templategate.check`, the CLI flags, the policy
  schema) backward compatible, or say clearly in the PR why a break is needed.
- Note user-visible changes under `## [Unreleased]` in `CHANGELOG.md`.
- Match the surrounding code style; the codebase is typed and ships `py.typed`,
  so keep annotations on public functions.

## Scope

TemplateGate reads documents and never writes them. Changes that make it edit,
convert, or auto-repair a document are out of scope — a failed candidate is for
a human to resolve.

Please do not attach real business documents to an issue. Reduce the problem to
a synthetic file first.
