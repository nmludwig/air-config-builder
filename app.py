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


def fetch_website(url):
    """Fetch website content, return text or error string."""
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; AIRConfigBuilder/1.0)"}
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw = resp.read(80000)  # cap at ~80KB
            text = raw.decode("utf-8", errors="replace")
            # strip obvious HTML tags roughly
            import re
            text = re.sub(r"<style[^>]*>.*?</style>", " ", text, flags=re.DOTALL)
            text = re.sub(r"<script[^>]*>.*?</script>", " ", text, flags=re.DOTALL)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text).strip()
            return text[:12000]  # send first 12k chars to Claude
    except Exception as e:
        return f"[Could not fetch website: {e}]"


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
    if not isinstance(prompt, str) or len(prompt) > 30000:
        return jsonify({"error": "Invalid prompt"}), 400

    # If the prompt contains a URL to fetch, do it server-side
    url = data.get("url", "").strip()
    if url:
        site_content = fetch_website(url)
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
