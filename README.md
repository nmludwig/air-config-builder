# RingCentral AIR Configuration Builder

A web app for the SE team to generate complete, copy-paste-ready RingCentral AI Receptionist (AIR) configurations from prospect details.

Built with Python + Flask + Anthropic API. Deploys to Render in ~5 minutes.

---

## What it does

Fill in a prospect's practice details across 5 steps:
1. Practice info (name, website, phone, address, services)
2. Top call reasons (tagged FAQ / Partial / Route)
3. Persona & greetings (AIR name, tone, language, company description)
4. Routing & skills (extensions, backup routing, enabled skills)
5. Generate → streams a complete config including:
   - Company description (500 char, ready to paste)
   - Business hours + after-hours greeting scripts
   - FAQ entries with exact Q&A pairs
   - Transfer by Context rules with keyword lists
   - Admin Portal setup checklist

Post-generation quick actions: soften tone, expand routing keywords, generate demo script, add success metrics.

---

## Deploy in 5 minutes

### Step 1 — Fork or clone to GitHub

```bash
git clone https://github.com/YOUR_ORG/air-config-builder
cd air-config-builder
git remote set-url origin https://github.com/YOUR_ORG/air-config-builder
git push -u origin main
```

Or just upload this folder as a new GitHub repo directly at github.com/new.

### Step 2 — Create a Render Web Service

1. Go to [render.com](https://render.com) → **New** → **Web Service**
2. Connect your GitHub account and select the `air-config-builder` repo
3. Render auto-detects `render.yaml` — settings will pre-fill:
   - **Runtime:** Python
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
4. Click **Create Web Service**

### Step 3 — Add your Anthropic API key

In your Render service dashboard:
1. Go to **Environment** tab
2. Click **Add Environment Variable**
3. Key: `ANTHROPIC_API_KEY`
4. Value: your key from [console.anthropic.com](https://console.anthropic.com)
5. Click **Save** → Render redeploys automatically

### Step 4 — Done

Your app is live at `https://air-config-builder.onrender.com` (or similar). Share the URL with your SE team.

---

## Local development

```bash
# Clone
git clone https://github.com/YOUR_ORG/air-config-builder
cd air-config-builder

# Set up Python env
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Add your API key
cp .env.example .env
# Edit .env and paste your ANTHROPIC_API_KEY

# Run
python app.py

# Open http://localhost:5000
```

---

## Project structure

```
air-config-builder/
├── app.py              # Flask server — proxies Anthropic API with SSE streaming
├── requirements.txt    # Python dependencies
├── render.yaml         # Render deployment config
├── .env.example        # Environment variable template
├── .gitignore
└── static/
    └── index.html      # Full frontend (single-file HTML/CSS/JS)
```

---

## Security notes

- The Anthropic API key lives only on the server (Render environment variable) — never in the frontend or git
- The `/api/generate` endpoint validates prompt length (25k char max)
- No data is stored — every generation is stateless
- For internal SE use only — do not expose to end customers without adding auth

---

## Customizing

**Change the AI model:** Edit `model=` in `app.py` (currently `claude-sonnet-4-20250514`)

**Add authentication:** Wrap the Flask app with HTTP basic auth or integrate with your SSO. Example with `flask-httpauth`:
```python
pip install flask-httpauth
```

**Add more quick-action buttons:** Edit the `quick-actions` div in `static/index.html` — each button calls `refine('your instruction here')`

**Adjust the system prompt:** Edit `buildPrompt()` in `static/index.html` to change what sections the AI generates

---

## Built by

Matthew Ludwig — RingCentral Solutions Engineering
