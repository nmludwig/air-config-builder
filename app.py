import os
import json
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


def fetch_via_jina(url):
    """Fetch website content via Jina AI Reader."""
    try:
        jina_url = f"https://r.jina.ai/{url}"
        headers = {
            "User-Agent": "Mozilla/5.0 (compatible; AIRConfigBuilder/1.0)",
            "Accept": "text/plain",
            "X-Return-Format": "markdown",
            "X-Remove-Selector": "nav,footer,header,.cookie-banner,#cookie-notice,.ads",
        }
        # Use API key if configured for higher rate limits
        jina_key = os.environ.get("JINA_API_KEY", "")
        if jina_key:
            headers["Authorization"] = f"Bearer {jina_key}"

        req = urllib.request.Request(jina_url, headers=headers)
        with urllib.request.urlopen(req, timeout=25) as resp:
            content = resp.read(100000).decode("utf-8", errors="replace")
            return content[:15000]
    except Exception as e:
        return f"[Website fetch failed: {e}. SE should paste website content manually into field 3.]"


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/generate", methods=["POST"])
def generate():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify({"error": "ANTHROPIC_API_KEY not set on server"}), 500

    data = request.get_json(silent=True)
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
            # Pasted content takes priority
            site_content = pasted[:15000]
        elif url:
            # Auto-fetch via Jina
            site_content = fetch_via_jina(url)
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
