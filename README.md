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

## Esecuzione CLI (utility Dropbox)

Verifica account:

```bash
python main.py whoami
```

Elenco root Dropbox:

```bash
python main.py list --path "" --limit 20
```

Avvio UI desktop standalone:

```bash
python main.py ui
```

La UI desktop contiene tre tab:
- `Rate annuali`: fogli anno (`YYYY`) del file `RM+RF+RC.ods`
- `Rate future`: vista del foglio `GenCal` suddivisa per anno tramite righe separatrici grigie
- `Completamento`: menu a tendina con viste `Per Rate`, `Per Totale`, `Per Capitale`

Con path ODS personalizzato:

```bash
python main.py ui --path "/Me/DEBITI/RM+RF+RC.ods"
```

## Parser ODS (standalone)

Smoke test rapido del parser sui fogli annuali di `RM+RF+RC.ods`:

```bash
python smoke_test.py
```

Note:
- percorso predefinito Dropbox: `/Me/DEBITI/RM+RF+RC.ods`
- esclude `GenCal` e legge solo i fogli anno (`YYYY`)

## CI (GitHub Actions)

La pipeline in `.github/workflows/ci.yml` esegue su push e pull request verso `main` e `development`:
- `ruff check .`
- compilazione rapida: `python -m py_compile auth.py main.py`

Verifica locale rapida (opzionale prima del push):

```bash
pip install -r requirements.txt
ruff check .
python -m py_compile auth.py main.py ods_service.py desktop_app.py smoke_test.py
```
