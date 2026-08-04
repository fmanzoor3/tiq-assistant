"""Diagnose why the local LLM returns an empty response.

Sends a few probe requests to the configured Qwen endpoint with different
parameter combinations and prints the FULL raw JSON response for each, so we
can see exactly where the content is (content vs reasoning_content), what
finish_reason says, and how tokens were spent.

Run on the WORK LAPTOP (where the endpoint is reachable):

    cd "C:\\path\\to\\tiq-assistant"
    python tools\\diagnose_llm.py

Then paste the output back. It reads the endpoint from your saved settings, so
enable + configure the AI Assistant in Settings first (or edit BASE_URL below).
"""

import json
import ssl
import sys
import urllib.request
import urllib.error
from pathlib import Path

# Make the package importable when run from the repo root.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from tiq_assistant.services.entry_generation_service import load_llm_config


def post(base_url, api_key, verify_ssl, payload, timeout=120):
    url = base_url.rstrip("/") + "/chat/completions"
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url, data=data, method="POST",
        headers={"Content-Type": "application/json",
                 "Authorization": f"Bearer {api_key}"},
    )
    ctx = None
    if base_url.lower().startswith("https") and not verify_ssl:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        return json.loads(resp.read().decode("utf-8"))


def get_models(base_url, api_key, verify_ssl, timeout=60):
    url = base_url.rstrip("/") + "/models"
    req = urllib.request.Request(url, method="GET",
                                 headers={"Authorization": f"Bearer {api_key}"})
    ctx = None
    if base_url.lower().startswith("https") and not verify_ssl:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        return json.loads(resp.read().decode("utf-8"))


def show(label, base, key, verify, payload):
    print("\n" + "=" * 70)
    print(f"PROBE: {label}")
    print("payload params:", {k: v for k, v in payload.items() if k != "messages"})
    try:
        result = post(base, key, verify, payload)
    except urllib.error.HTTPError as e:
        body = ""
        try:
            body = e.read().decode("utf-8", "ignore")[:500]
        except Exception:
            pass
        print(f"HTTP {e.code}: {body or e.reason}")
        return
    except Exception as e:
        print(f"ERROR: {e}")
        return

    print("--- RAW RESPONSE ---")
    print(json.dumps(result, indent=2)[:3000])

    # Highlight the key fields.
    try:
        choice = result["choices"][0]
        msg = choice.get("message", {})
        print("\n--- SUMMARY ---")
        print("finish_reason   :", choice.get("finish_reason"))
        print("content (len)   :", len(msg.get("content") or ""))
        print("content (start) :", repr((msg.get("content") or "")[:200]))
        if "reasoning_content" in msg:
            rc = msg.get("reasoning_content") or ""
            print("reasoning_content (len)  :", len(rc))
            print("reasoning_content (start):", repr(rc[:200]))
        print("usage           :", result.get("usage"))
    except Exception as e:
        print("Could not summarise:", e)


def main():
    cfg = load_llm_config()
    base = cfg.base_url
    key = cfg.api_key
    verify = cfg.verify_ssl
    print("Endpoint:", base, "| verify_ssl:", verify)

    try:
        models = get_models(base, key, verify)
        model_id = (models.get("data") or [{}])[0].get("id")
        print("Model:", model_id)
    except Exception as e:
        print("Could not list models:", e)
        model_id = cfg.model or ""

    msgs = [{"role": "user", "content":
             'Return ONLY this JSON, nothing else: {"entries":[{"hours":1,"description":"test"}]}'}]

    # Probe 1: bare minimal request (no thinking switches, no json mode, big budget)
    show("minimal, max_tokens=2048", base, key, verify, {
        "model": model_id, "messages": msgs, "temperature": 0, "max_tokens": 2048,
    })

    # Probe 2: with enable_thinking=False switches
    show("enable_thinking=False", base, key, verify, {
        "model": model_id, "messages": msgs, "temperature": 0, "max_tokens": 2048,
        "chat_template_kwargs": {"enable_thinking": False},
    })

    # Probe 3: Qwen /no_think directive in the prompt
    show("/no_think directive", base, key, verify, {
        "model": model_id,
        "messages": [{"role": "user", "content": "/no_think " + msgs[0]["content"]}],
        "temperature": 0, "max_tokens": 2048,
    })

    # Probe 4: json mode on
    show("response_format=json_object", base, key, verify, {
        "model": model_id, "messages": msgs, "temperature": 0, "max_tokens": 2048,
        "response_format": {"type": "json_object"},
    })

    print("\nDone. Paste the whole output above back for diagnosis.")


if __name__ == "__main__":
    main()
