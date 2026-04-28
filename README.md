# Conciliador BROU separado

Un solo repo y un solo Render para correr:

- solo USD
- solo UYU
- solo EUR
- las tres monedas juntas

## Estructura

```text
app.py
requirements.txt
README.md
templates/
  index.html
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
