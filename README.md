# RingCentral AIR Configuration Builder

An AI-powered web app for RingCentral Solutions Engineers to generate complete, copy-paste-ready AI Receptionist (AIR) configurations for any prospect — in minutes.

Built with Python + Flask + Anthropic Claude + Firecrawl. Hosted on RingCentral CELAB infrastructure.

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
- **Hosting:** Linux Mint server, Nginx reverse proxy, systemd, RingCentral CELAB
- **Frontend:** Single-file HTML/CSS/JS, no framework

---

## Live URL

**[https://air-config-builder.celab.ringcentral.com/](https://air-config-builder.celab.ringcentral.com/)** — RingCentral internal (SE team access)

## Deploy

```bash
# On the CELAB server (matthew.ludwig@matthewludwig.celab.ringcentral.com)
cd /opt/air-config-builder && git pull origin main && sudo systemctl restart air-config-builder

# Check logs
sudo journalctl -u air-config-builder -f
```

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
├── docx_generator.py   # Word doc generator (python-docx)
├── .env.example        # Environment variable template
├── .gitignore
└── static/
    └── index.html      # Full frontend — single-file HTML/CSS/JS
```

---

## Security

- API keys live in `/opt/air-config-builder/.env` on the server — never in frontend or git
- `/api/generate` validates prompt length (40k char max)
- No data is stored — every generation is stateless
- For internal SE use — add authentication before exposing to customers

---

## Will the Output Work with RingCentral AIR?

The Word doc is a **deliverable for the SE** — it is not uploaded to AIR directly. The SE uses it as a copy-paste reference while configuring the AIR Admin Portal by hand.

"Will it work with AIR" really means: does the content Claude generates match what AIR actually expects in each field?

### Key Validation Checks

After generating, review these before the SE configures AIR:

| Field | What to Check | Why It Matters |
|---|---|---|
| **Company Description** | ≤500 characters, written as a behavioral prompt for AIR (not marketing copy) | This is AIR's core identity — bad description = bad behavior across all calls |
| **FAQ Answers** | 2–4 sentences max, plain conversational language | AIR reads these aloud — paragraph-length answers sound robotic |
| **Transfer by Context keywords** | Specific, caller-realistic phrases (not generic words like "help") | AIR does fuzzy keyword matching — vague keywords cause misfires |
| **Greetings** | Short sentences, speakable out loud, no credentials or marketing | AIR speaks the greeting verbatim — complex grammar sounds unnatural |

### Quick Sanity Test

Run a generation on a real prospect URL and check:
- Company description character count is under 500
- FAQ answers are 2–4 sentences each
- Greetings read naturally when spoken aloud
- Routing keywords are specific to what callers actually say

---

## Built By

Matthew Ludwig — RingCentral Solutions Engineering  
May 2026
