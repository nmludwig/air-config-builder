# AIR Config Builder — Claude Code Context

## What This Project Is
A web app for RingCentral Solutions Engineers to generate complete, copy-paste-ready AI Receptionist (AIR) configurations for any prospect. SE enters a website URL and top call reasons — the app fetches the site via Firecrawl, sends it to Claude, and generates a full configuration document covering every AIR Admin Portal field. Output streams to screen and can be downloaded as a formatted Word doc.

## Live URLs
- **RingCentral internal:** https://air-config-builder.celab.ringcentral.com/
- **Render (backup):** https://air-config-builder.onrender.com
- **GitHub:** https://github.com/nmludwig/air-config-builder (private)

## Server Details
- Linux Mint server, Nginx reverse proxy, systemd, SSL via RingCentral wildcard cert
- **Service name:** air-config-builder
- **App path:** /opt/air-config-builder/
- **Gunicorn:** 127.0.0.1:8001, timeout 300s
- **Venv:** /opt/air-config-builder/venv/
- **Deploy:** git pull origin main && sudo systemctl restart air-config-builder
- **Logs:** sudo journalctl -u air-config-builder -f

## Project Structure
- app.py — Flask server, Firecrawl fetch, Anthropic SSE stream, Word doc export
- docx_generator.py — Python Word doc generator (python-docx)
- static/index.html — Full frontend, single file, contains the Claude prompt
- requirements.txt — flask, anthropic, gunicorn, python-dotenv, python-docx
- render.yaml — Render deployment config

## The Claude Prompt
Lives inside buildPrompt() in static/index.html as a JS template literal. Uses __WEBSITE_CONTENT__ placeholder replaced server-side with Firecrawl content. Outputs 7 sections matching the Dr. Liu AIR Configuration Playbook format.

## The 7 AIR Fields
1. Company Description — HARD (500 char, system prompt for AIR, NOT marketing copy)
2. Custom Greeting — HARD (what AIR can DO, NOT doctor credentials)
3. FAQ Skill — HARD (max 10, MUST be numbered by call reason 1-10)
4. Transfer by Context — SOFT (max 10 routing rules with keywords)
5. Knowledge Base — MIXED (URLs + documents, rendered as 3 tables)
6. SMS Skill — SOFT (message templates)
7. Lead Capture — MIXED (intake questions)

## Word Doc Format
Matches Dr. Liu AIR Configuration Playbook style:
- Navy H1 with blue bottom border, blue H2, blue H3
- Coverage plan table: green/amber/red colored AIR Handles? column
- FAQ tables: navy header QUESTION | ANSWER, italic questions left, answers right
- Routing table: 3-col TRIGGER KEYWORDS | ROUTE TO | COVERS
- Copy box: green-tinted Courier New for company description
- Dark navy box: white italic for greetings
- Callout boxes: blue info (i), amber warning (!), red danger (!), green success (checkmark)
- Page break between every major section

## Key Rules
- FAQ entries MUST be numbered by call reason (Call Reason 4 — Mesh Type), NOT generic names
- Company description = system prompt for AIR behavior, NOT patient-facing marketing
- Greeting = "I can help you [what AIR does]", NOT doctor credentials or marketing
- Prohibition rules go in Company Description (hard) and Transfer by Context (soft)
- Always validate JS after editing index.html: node --check file.js

## Environment Variables
- ANTHROPIC_API_KEY
- FIRECRAWL_API_KEY

## Owner
Matthew Ludwig — RingCentral Solutions Engineering
