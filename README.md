# FSB Reconciliation Tool

Flask-based E-Way Bill reconciliation tool with Excel auth and Render hosting support.

## Project Structure

```
fsb_app/
├── server.py          # Flask app (auth + file handling)
├── reconcile.py       # Pure reconciliation logic (do not modify)
├── auth_db.xlsx       # User credentials (logid | password)
├── requirements.txt   # Python dependencies
├── render.yaml        # Render auto-deploy config
└── static/
    └── index.html     # Frontend UI
```

## Running Locally

```bash
pip install -r requirements.txt
python server.py
# Open http://localhost:5000
```

## Deploying to Render

### Option A — Auto-deploy with render.yaml (recommended)

1. Push this folder to a GitHub repository.
2. Go to https://render.com → New → Blueprint → connect your repo.
3. Render reads `render.yaml` and sets everything up automatically.
4. Upload `auth_db.xlsx` to the repo (or use Render's Disk feature for persistence).

### Option B — Manual setup on Render

1. New → Web Service → connect GitHub repo.
2. Settings:
   - **Runtime**: Python 3
   - **Build command**: `pip install -r requirements.txt`
   - **Start command**: `gunicorn server:app`
3. Environment variables (under Environment tab):
   - `FSB_SECRET` → any long random string (e.g. generate with `openssl rand -hex 32`)
4. Deploy.

### Important Notes for Render

- **auth_db.xlsx must be included in your repo** (or mounted via Render Disk).
  - If you use Render Disk, set `AUTH_DB_PATH=/mnt/data/auth_db.xlsx` as an env var.
- **Sessions**: The `FSB_SECRET` env var must be set — without it, sessions reset on every deploy/restart.
- **File size**: Default upload limit is 50 MB. Adjust `MAX_CONTENT_LENGTH` in server.py if needed.
- **Free tier**: Render's free tier spins down after inactivity; the first request after sleep may be slow.

## Environment Variables

| Variable       | Required | Description                                      |
|----------------|----------|--------------------------------------------------|
| `FSB_SECRET`   | Yes (prod)| Secret key for Flask sessions                   |
| `AUTH_DB_PATH` | No       | Path to auth_db.xlsx (default: `auth_db.xlsx`)  |
| `PORT`         | No       | Port to listen on (Render sets this automatically)|
| `FLASK_DEBUG`  | No       | Set to `true` for debug mode (never in prod)     |

## Auth DB Format

`auth_db.xlsx` must have exactly two columns:

| logid   | password |
|---------|----------|
| Shaurya | 1234     |
| M.Idris | 4567     |
