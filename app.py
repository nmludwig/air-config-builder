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

            print("Firecrawl result keys:", list(result.keys()))
            print("Firecrawl success:", result.get("success"))
            print("Firecrawl data keys:", list(result.get("data", {}).keys()) if result.get("data") else "no data")
            md = result.get("data", {}).get("markdown", "")
            print("Markdown length:", len(md))
            print("Markdown preview:", repr(md[:200]))
            if result.get("success") and result.get("data", {}).get("markdown"):
                content = result["data"]["markdown"]
                # Sanity check — detect if we got a CAPTCHA/block page anyway
                block_signals = ["captcha", "just a moment", "enable javascript", "verify you are human", "access denied"]
                if any(s in content.lower() for s in block_signals) or len(content.strip()) < 100:
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
