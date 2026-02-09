# Krantenplanner (PARSER → DEEL 1 → DEEL 2)

Web-app (FastAPI) die één Excel upload accepteert (input voor PARSER) en vervolgens automatisch:
1. PARSER draait → maakt Verhalenaanbod/Planningsoverzicht
2. DEEL 1 draait → maakt `Krantenplanning.xlsx`
3. DEEL 2 draait → maakt `Krantenplanning_handout.pdf`

## Lokaal draaien
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```
Ga naar http://127.0.0.1:8000

## Deploy op Render
- Zet deze repo op GitHub
- Maak een nieuwe **Web Service** in Render en koppel je repo
- Render pakt automatisch `render.yaml`

## Outputs
- Download Excel (DEEL 1): `Krantenplanning.xlsx`
- Download PDF (DEEL 2): `Krantenplanning_handout.pdf`
