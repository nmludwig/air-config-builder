import os
import json
import re
import urllib.request
import urllib.error
from flask import Flask, request, Response, send_from_directory, jsonify
from flask_cors import CORS
import anthropic
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__, static_folder="static", static_url_path="")
CORS(app)

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))


def fetch_via_firecrawl(url):
    """Fetch website content via Firecrawl — handles Cloudflare, JS rendering, bot protection."""
    api_key = os.environ.get("FIRECRAWL_API_KEY", "")
    if not api_key:
        return None, "FIRECRAWL_API_KEY not configured on server"

    try:
        payload = json.dumps({
            "url": url,
            "formats": ["markdown"],
            "onlyMainContent": True,
            "timeout": 20000
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.firecrawl.dev/v1/scrape",
            data=payload,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST"
        )

        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode("utf-8"))

            if result.get("success") and result.get("data", {}).get("markdown"):
                content = result["data"]["markdown"]
                # Only flag genuine block pages — not cookie banners or normal pages
                # Require multiple signals AND very short content to avoid false positives
                hard_block_signals = ["verify you are human", "just a moment...", "enable javascript and cookies to continue"]
                is_hard_block = any(s in content.lower() for s in hard_block_signals) and len(content.strip()) < 500
                if is_hard_block:
                    return None, "CAPTCHA_PROTECTED"
                return content[:15000], None
            else:
                return None, result.get("error", "Firecrawl returned no content")

    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        return None, f"Firecrawl error {e.code}: {body[:200]}"
    except Exception as e:
        return None, f"Fetch failed: {str(e)}"


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


SHEETS_WEBHOOK = os.getenv("SHEETS_WEBHOOK_URL", "")

def get_client_ip():
    # Respect X-Forwarded-For set by Nginx reverse proxy
    forwarded = request.headers.get("X-Forwarded-For", "")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return request.remote_addr or ""

def log_to_sheet(email, url=""):
    if not SHEETS_WEBHOOK:
        return
    try:
        payload = json.dumps({"email": email, "url": url, "ip": get_client_ip()}).encode("utf-8")
        req = urllib.request.Request(SHEETS_WEBHOOK, data=payload,
                                     headers={"Content-Type": "application/json"}, method="POST")
        urllib.request.urlopen(req, timeout=5)
    except Exception:
        pass

APP_PASSWORD = os.getenv("APP_PASSWORD", "")

def is_rc_email(email):
    return isinstance(email, str) and email.strip().lower().endswith("@ringcentral.com")

def require_rc_auth():
    email = request.headers.get("X-User-Email", "")
    password = request.headers.get("X-App-Password", "")
    if not is_rc_email(email):
        return jsonify({"error": "Unauthorized — RingCentral email required"}), 401
    if APP_PASSWORD and password != APP_PASSWORD:
        return jsonify({"error": "Unauthorized — incorrect password"}), 401
    return None

@app.route("/api/verify", methods=["POST"])
def verify():
    err = require_rc_auth()
    if err: return err
    return jsonify({"ok": True})


@app.route("/api/generate", methods=["POST"])
def generate():
    auth_err = require_rc_auth()
    if auth_err: return auth_err

    data = request.get_json(silent=True)
    log_to_sheet(request.headers.get("X-User-Email", ""), (data or {}).get("url", ""))

    if not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify({"error": "ANTHROPIC_API_KEY not set on server"}), 500

    if not data or "prompt" not in data:
        return jsonify({"error": "Missing prompt"}), 400

    prompt = data["prompt"]
    if not isinstance(prompt, str) or len(prompt) > 40000:
        return jsonify({"error": "Invalid prompt"}), 400

    # Resolve website content
    if "__WEBSITE_CONTENT__" in prompt:
        url = data.get("url", "").strip()
        pasted = data.get("pasted", "").strip()

        if pasted and len(pasted) > 50:
            # Pasted content always takes priority
            site_content = pasted[:15000]
        elif url:
            # Auto-fetch via Firecrawl
            content, error = fetch_via_firecrawl(url)
            if content is None:
                if error == "CAPTCHA_PROTECTED":
                    return jsonify({
                        "error": "CAPTCHA_PROTECTED",
                        "message": "This website is bot-protected. Please open it in your browser, press Cmd+A then Cmd+C to copy all text, and paste it into field 3."
                    }), 422
                elif "FIRECRAWL_API_KEY not configured" in str(error):
                    return jsonify({
                        "error": "FIRECRAWL_NOT_CONFIGURED",
                        "message": "Firecrawl API key not set. Please paste website content manually in field 3."
                    }), 422
                else:
                    return jsonify({
                        "error": "FETCH_FAILED",
                        "message": f"Could not fetch website ({error}). Please paste website content manually in field 3."
                    }), 422
            site_content = content
        else:
            site_content = "[No website content provided]"

        prompt = prompt.replace("__WEBSITE_CONTENT__", site_content)

    def stream():
        with client.messages.stream(
            model="claude-sonnet-4-5",
            max_tokens=16000,
            messages=[{"role": "user", "content": prompt}],
        ) as s:
            for text in s.text_stream:
                payload = json.dumps({"text": text})
                yield f"data: {payload}\n\n"
        yield "data: [DONE]\n\n"

    return Response(stream(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/api/suggest", methods=["POST"])
def suggest():
    auth_err = require_rc_auth()
    if auth_err: return auth_err

    data = request.get_json(silent=True)
    url = (data or {}).get("url", "").strip()
    if not url:
        return jsonify({"error": "Missing URL"}), 400

    content, error = fetch_via_firecrawl(url)
    if content is None:
        return jsonify({"error": "FETCH_FAILED", "message": f"Could not read website ({error}). Try pasting content into field 3."}), 422

    prompt = f"""You are an expert at understanding small and medium businesses from their website.

Read the following website content and identify the top 10 reasons callers most likely contact this business by phone.

Write them as a numbered list, one per line, phrased exactly as a staff member or office manager would describe them — short, plain language, specific to this business. No generic entries like "general inquiries."

Website: {url}

Website content:
{content[:8000]}

Output ONLY the numbered list, nothing else. Example format:
1. Schedule a new patient consultation
2. Check on insurance coverage
3. Questions about post-op care"""

    result = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=512,
        messages=[{"role": "user", "content": prompt}]
    )
    reasons = result.content[0].text.strip()
    return jsonify({"reasons": reasons})


@app.route("/api/export", methods=["POST"])
def export_docx():
    auth_err = require_rc_auth()
    if auth_err: return auth_err

    from docx_generator import generate_docx
    data = request.get_json(silent=True)
    if not data or "content" not in data:
        return jsonify({"error": "Missing content"}), 400

    biz_name = data.get("bizName", "Practice").replace('**', '').replace('*', '').strip()
    prepared_by = data.get("preparedBy", "RingCentral SE")
    content_text = data.get("content", "")

    try:
        docx_bytes = generate_docx(biz_name, content_text, prepared_by)
        safe_name = re.sub(r'[^a-zA-Z0-9_-]', '_', biz_name)[:40]
        return Response(
            docx_bytes,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers={'Content-Disposition': f'attachment; filename="AIR_Config_{safe_name}.docx"'}
        )
    except Exception as e:
        import traceback
        print("Export error:", traceback.format_exc())
        return jsonify({"error": "Doc generation failed", "detail": str(e)}), 500



if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
