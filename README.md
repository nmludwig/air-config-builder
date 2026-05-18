# RingCentral AIR Configuration Builder

An AI-powered web app for RingCentral Solutions Engineers to generate complete, copy-paste-ready AI Receptionist (AIR) configurations for any prospect — in minutes.

Built with Python + Flask + Anthropic Claude + Firecrawl. Deployed on Render.

---

## What It Does

Enter two things:
1. **The prospect's website URL** — Firecrawl automatically reads the site
2. **Their top call reasons** — paste directly from staff notes or email

Click **Generate AIR Configuration** and get a complete configuration covering all 7 AIR fields:

| Section | AIR Field | Determinism | Admin Portal Path |
|---|---|---|---|
| Company Description | Company Info | **HARD** | AI Receptionist → Company Info |
| Custom Greeting | Business + After-Hours | **HARD** | Skills → Custom Greeting |
| FAQ Entries (max 10) | Frequently Asked Questions | **HARD** | Skills → FAQs → Add |
| Transfer by Context (max 10) | Routing Rules | **SOFT** | Skills → Transfer by Context → Add Rule |
| Knowledge Base Strategy | Knowledge Hub | **MIXED** | Knowledge Hub → Add |
| Appointment Booking + SMS | Booking + SMS Skills | **SOFT** | Skills → Appointment Booking / SMS |
| Lead Capture | Lead Capture Skill | **MIXED** | Skills → Lead Capture |

Plus:
- **Determinism strategy** — which facts must go in FAQ skill (exact wording) vs Knowledge Base (paraphrasing OK)
- **Prohibition rules** — what AIR must never say, embedded into Company Description and Transfer by Context rules
- **Persona recommendation** — AI receptionist name chosen by cultural, demographic, and industry reasoning
- **Deployment recommendations** — what would make this AIR better than average
- **Persona critique** — FAQ answer quality review, non-determinism risk flags, demo wow moments
- **Admin Portal setup checklist** — numbered step-by-step, reference the generated values

---

## Determinism Framework

Understanding how each field works helps SEs write better configurations:

**HARD Deterministic** — AIR uses your exact text. Every word matters.
- Company Description, Custom Greeting, FAQ Skill

**SOFT Deterministic** — You define rules, AIR executes them with slight natural variation.
- Transfer by Context, SMS templates

**MIXED** — You define structure, AIR generates natural language from your content.
- Knowledge Base, Lead Capture

---

## Stack

- **Backend:** Python + Flask + Gunicorn
- **AI:** Anthropic Claude Sonnet (claude-sonnet-4-5), 16,000 max tokens
- **Web fetching:** Firecrawl API — handles Cloudflare, JavaScript rendering, bot protection
- **Hosting:** Render (web service)
- **Frontend:** Single-file HTML/CSS/JS, no framework

---

## Deploy in 5 Minutes

### 1. Fork or clone to GitHub

```bash
git clone https://github.com/nmludwig/air-config-builder
cd air-config-builder
git remote set-url origin https://github.com/YOUR_ORG/air-config-builder.git
git push -u origin main
```

### 2. Create a Render Web Service

1. Go to [render.com](https://render.com) → **New** → **Web Service**
2. Connect your GitHub repo
3. `render.yaml` auto-fills all settings
4. Click **Create Web Service**

### 3. Add Environment Variables

In Render dashboard → **Environment** tab, add:

| Key | Value | Required |
|---|---|---|
| `ANTHROPIC_API_KEY` | Your key from [console.anthropic.com](https://console.anthropic.com) | ✅ Required |
| `FIRECRAWL_API_KEY` | Your key from [firecrawl.dev](https://firecrawl.dev) | ✅ Required for auto-fetch |

Click **Save** — Render redeploys automatically.

### 4. Done

Your app is live at `https://air-config-builder.onrender.com`. Share the URL with your SE team.

---

## Local Development

```bash
git clone https://github.com/nmludwig/air-config-builder
cd air-config-builder
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
# Edit .env and add your ANTHROPIC_API_KEY and FIRECRAWL_API_KEY
python app.py
# Open http://localhost:5000
```

---

## How Website Fetching Works

1. SE enters the prospect's URL — app auto-fetches via Firecrawl
2. Firecrawl handles bot protection, JavaScript rendering, and Cloudflare
3. If the SE pastes content into field 3, that takes priority over auto-fetch
4. If a site is truly inaccessible, field 3 shows clear instructions: open the site, Cmd+A, Cmd+C, paste

---

## Project Structure

```
air-config-builder/
├── app.py              # Flask server — Firecrawl fetch + Anthropic API proxy with SSE streaming
├── requirements.txt    # Python dependencies
├── render.yaml         # Render deployment config
├── .env.example        # Environment variable template
├── .gitignore
└── static/
    └── index.html      # Full frontend — single-file HTML/CSS/JS
```

---

## Security

- API keys live only in Render environment variables — never in frontend or git
- `/api/generate` validates prompt length (40k char max)
- No data is stored — every generation is stateless
- For internal SE use — add authentication before exposing to customers

---

## Built By

Matthew Ludwig — RingCentral Solutions Engineering  
May 2026
