# BROU Conciliador Bancario V2

Un solo repo y un solo Render para correr conciliaciones separadas o conjuntas:

- USD
- UYU
- EUR
- Todo

Incluye interfaz con logo BROU.

## Estructura

```text
app.py
requirements.txt
README.md
templates/
  index.html
static/
  brou-logo.png
scripts/
  SCRIPT_USD_BROU.py
  SCRIPT_UYU_BROU.py
  SCRIPT_EUR_BROU.py
```

## Render

Build Command:

```text
pip install -r requirements.txt
```

Start Command:

```text
gunicorn app:app --timeout 600
```
