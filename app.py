import os
import json
import re
import secrets
import urllib.request
import urllib.error
from flask import Flask, request, Response, send_from_directory, jsonify, session, redirect
from flask_cors import CORS
import anthropic
import requests as http_requests
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__, static_folder="static", static_url_path="")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", secrets.token_hex(32))
CORS(app, supports_credentials=True)

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

RC_CLIENT_ID     = os.environ.get("RC_CLIENT_ID", "")
RC_CLIENT_SECRET = os.environ.get("RC_CLIENT_SECRET", "")
RC_REDIRECT_URI  = os.environ.get("RC_REDIRECT_URI", "https://air-config-builder.celab.ringcentral.com/auth/callback")
RC_AUTH_URL      = "https://platform.ringcentral.com/restapi/oauth/authorize"
RC_TOKEN_URL     = "https://platform.ringcentral.com/restapi/oauth/token"
RC_USERINFO_URL  = "https://platform.ringcentral.com/restapi/v1.0/account/~/extension/~"  # fallback

def decode_jwt_payload(token):
    """Decode JWT payload without verifying signature — just to read claims."""
    import base64
    try:
        payload_b64 = token.split(".")[1]
        # Fix padding
        payload_b64 += "=" * (4 - len(payload_b64) % 4)
        return json.loads(base64.urlsafe_b64decode(payload_b64))
    except Exception:
        return {}


# ── Firecrawl ─────────────────────────────────────────────────────────────────

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


# ── Google Sheets logging ──────────────────────────────────────────────────────

SHEETS_WEBHOOK = os.getenv("SHEETS_WEBHOOK_URL", "")

def get_client_ip():
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


# ── Auth helpers ───────────────────────────────────────────────────────────────

def current_user():
    """Return the logged-in user's email from session, or None."""
    return session.get("user_email")

def require_auth():
    """Return a 401 response if not logged in, else None."""
    if not current_user():
        return jsonify({"error": "Unauthorized — please sign in"}), 401
    return None


# ── OAuth routes ───────────────────────────────────────────────────────────────

@app.route("/auth/login")
def auth_login():
    state = secrets.token_urlsafe(16)
    session["oauth_state"] = state
    params = (
        f"?response_type=code"
        f"&client_id={RC_CLIENT_ID}"
        f"&redirect_uri={RC_REDIRECT_URI}"
        f"&state={state}"
    )
    return redirect(RC_AUTH_URL + params)


@app.route("/auth/callback")
def auth_callback():
    error = request.args.get("error")
    if error:
        return f"OAuth error: {error}", 400

    state = request.args.get("state", "")
    if state != session.pop("oauth_state", None):
        return "Invalid state parameter", 400

    code = request.args.get("code", "")
    if not code:
        return "Missing authorization code", 400

    # Exchange code for token
    try:
        token_resp = http_requests.post(
            RC_TOKEN_URL,
            data={
                "grant_type": "authorization_code",
                "code": code,
                "redirect_uri": RC_REDIRECT_URI,
            },
            auth=(RC_CLIENT_ID, RC_CLIENT_SECRET),
            timeout=10,
        )
        token_resp.raise_for_status()
        token_data = token_resp.json()
    except Exception as e:
        return f"Token exchange failed: {e}", 500

    access_token = token_data.get("access_token", "")

    # Try to get email from JWT claims first (no extra API call needed)
    claims = decode_jwt_payload(access_token)
    email  = (claims.get("email") or claims.get("preferred_username") or "").strip().lower()
    name   = claims.get("name") or claims.get("given_name") or ""

    # Fall back to userinfo API if JWT didn't have email
    if not email:
        try:
            userinfo_resp = http_requests.get(
                RC_USERINFO_URL,
                headers={"Authorization": f"Bearer {access_token}"},
                timeout=10,
            )
            userinfo_resp.raise_for_status()
            userinfo = userinfo_resp.json()
            email = (userinfo.get("contact", {}).get("email") or "").strip().lower()
            name  = name or f"{userinfo.get('contact', {}).get('firstName', '')} {userinfo.get('contact', {}).get('lastName', '')}".strip()
        except Exception as e:
            return f"Failed to fetch user info: {e}", 500

    # Also check token_data for email (RC sometimes puts it there)
    if not email:
        email = (token_data.get("owner_id") or "").strip().lower()

    if not email:
        # Use owner_id (numeric RC user ID) as fallback identifier
        email = token_data.get("owner_id", "unknown@ringcentral")

    session["user_email"] = email
    session["user_name"]  = name
    session.permanent = True

    log_to_sheet(email)
    return redirect("/")


@app.route("/auth/logout")
def auth_logout():
    session.clear()
    return redirect("/")


@app.route("/auth/me")
def auth_me():
    email = current_user()
    if not email:
        return jsonify({"authenticated": False}), 401
    return jsonify({"authenticated": True, "email": email, "name": session.get("user_name", email)})


# ── App routes ─────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/generate", methods=["POST"])
def generate():
    auth_err = require_auth()
    if auth_err: return auth_err

    data = request.get_json(silent=True)
    log_to_sheet(current_user(), (data or {}).get("url", ""))

    if not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify({"error": "ANTHROPIC_API_KEY not set on server"}), 500

    if not data or "prompt" not in data:
        return jsonify({"error": "Missing prompt"}), 400

    prompt = data["prompt"]
    if not isinstance(prompt, str) or len(prompt) > 40000:
        return jsonify({"error": "Invalid prompt"}), 400

    # Resolve website content
    if "__WEBSITE_CONTENT__" in prompt:
        url    = data.get("url", "").strip()
        pasted = data.get("pasted", "").strip()

        if pasted and len(pasted) > 50:
            site_content = pasted[:15000]
        elif url:
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
    auth_err = require_auth()
    if auth_err: return auth_err

    data = request.get_json(silent=True)
    url  = (data or {}).get("url", "").strip()
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
    auth_err = require_auth()
    if auth_err: return auth_err

    from docx_generator import generate_docx
    data = request.get_json(silent=True)
    if not data or "content" not in data:
        return jsonify({"error": "Missing content"}), 400

    biz_name    = data.get("bizName", "Practice").replace('**', '').replace('*', '').strip()
    prepared_by = data.get("preparedBy", "RingCentral SE")
    content_text = data.get("content", "")

    try:
        docx_bytes = generate_docx(biz_name, content_text, prepared_by)
        safe_name  = re.sub(r'[^a-zA-Z0-9_-]', '_', biz_name)[:40]
        return Response(
            docx_bytes,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers={'Content-Disposition': f'attachment; filename="AIR_Config_{safe_name}.docx"'}
        )
    except Exception as e:
        import traceback
        print("Export error:", traceback.format_exc())
        return jsonify({"error": "Doc generation failed"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
