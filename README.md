# Excel BOQ Parser API

Petit service FastAPI qui :

- reçoit un fichier Excel (`multipart/form-data`, champ `file`)
- ouvre **la dernière feuille active** (celle qui était ouverte à la sauvegarde)
- ignore les **lignes masquées**
- renvoie les lignes visibles dans un JSON exploitable par un agent IA ou n8n.

## Installation locale

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload
