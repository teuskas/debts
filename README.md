# Debts

Progetto standalone Python (no mobile) con accesso Dropbox.

## Requisiti

- Python 3.11+
- file `.env` con:
  - `DROPBOX_APP_KEY`
  - `DROPBOX_APP_SECRET`
  - `DROPBOX_REFRESH_TOKEN`

## Setup rapido

```bash
cd /home/matteo/workspace/Debts
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Esecuzione

Verifica account:

```bash
python main.py whoami
```

Elenco root Dropbox:

```bash
python main.py list --path "" --limit 20
```

## CI (GitHub Actions)

La pipeline in `.github/workflows/ci.yml` esegue su push e pull request verso `main` e `development`:
- `ruff check .`
- compilazione rapida: `python -m py_compile auth.py main.py`

Verifica locale rapida (opzionale prima del push):

```bash
pip install ruff
ruff check .
python -m py_compile auth.py main.py
```
