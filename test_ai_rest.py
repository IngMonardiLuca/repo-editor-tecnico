"""
Test headless della conversione AI da SDK a REST (GeminiWorker).

NON usa rete ne' chiavi vere: sostituisce urllib.request.urlopen con un
finto che cattura la richiesta (URL/header/body) e restituisce una risposta
JSON prefabbricata per ciascun provider. Verifica che:
  - venga costruito l'URL/endpoint corretto per ogni provider;
  - gli header di autenticazione siano quelli giusti;
  - la modalita' "avanzato" aggiunga i parametri di reasoning attesi;
  - il parsing della risposta restituisca il testo corretto;
  - un errore HTTP venga trasformato in un messaggio leggibile.

Esecuzione (dalla cartella PROGETTO EDITOR):
    QT_QPA_PLATFORM=offscreen ./.venv/Scripts/python.exe test_ai_rest.py
"""

import io
import json
import os
import urllib.request
import urllib.error

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import main


failures = []


def check(cond, msg):
    if cond:
        print(f"  OK   {msg}")
    else:
        print(f"  FAIL {msg}")
        failures.append(msg)


class FakeResp:
    """Context manager che imita la risposta di urlopen."""
    def __init__(self, payload_dict):
        self._data = json.dumps(payload_dict).encode("utf-8")

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def read(self):
        return self._data


def install_capture(response_dict, captured):
    """Sostituisce urlopen catturando la Request e tornando response_dict."""
    def fake_urlopen(req, timeout=None):
        captured["url"] = req.full_url
        captured["headers"] = dict(req.header_items())
        captured["body"] = json.loads(req.data.decode("utf-8"))
        captured["timeout"] = timeout
        return FakeResp(response_dict)
    urllib.request.urlopen = fake_urlopen


def make_worker(provider, reasoning="base"):
    return main.GeminiWorker(
        api_key="FAKE-KEY-123",
        model_name="model-x",
        prompt_text="ciao",
        mode="Genera testo",
        style="Tecnico formale",
        provider=provider,
        reasoning=reasoning,
    )


def test_gemini():
    print("== Gemini (base) ==")
    cap = {}
    install_capture(
        {"candidates": [{"content": {"parts": [{"text": "  risposta gemini "}]}}]},
        cap,
    )
    w = make_worker("gemini")
    out = w._run_gemini("PROMPT-G")

    check("generativelanguage.googleapis.com" in cap["url"],
          "Gemini: host REST corretto")
    check("model-x:generateContent" in cap["url"],
          "Gemini: modello + :generateContent nell'URL")
    check("key=FAKE-KEY-123" in cap["url"],
          "Gemini: api_key passata in querystring")
    check(cap["body"]["contents"][0]["parts"][0]["text"] == "PROMPT-G",
          "Gemini: prompt nel body")
    check("generationConfig" not in cap["body"],
          "Gemini base: nessun thinkingConfig")
    check(out == "risposta gemini",
          "Gemini: testo estratto e strip-ato")

    print("== Gemini (avanzato) ==")
    cap = {}
    install_capture(
        {"candidates": [{"content": {"parts": [{"text": "x"}]}}]}, cap)
    make_worker("gemini", "avanzato")._run_gemini("P")
    check(cap["body"]["generationConfig"]["thinkingConfig"]["thinkingBudget"] == 8192,
          "Gemini avanzato: thinkingBudget=8192")


def test_openai():
    print("== OpenAI (base) ==")
    cap = {}
    install_capture(
        {"output": [{"content": [{"type": "output_text", "text": "risposta openai"}]}]},
        cap,
    )
    w = make_worker("openai")
    out = w._run_openai("PROMPT-O")

    check(cap["url"] == "https://api.openai.com/v1/responses",
          "OpenAI: endpoint Responses corretto")
    # gli header sono normalizzati Capitalized da urllib
    auth = cap["headers"].get("Authorization", "")
    check(auth == "Bearer FAKE-KEY-123",
          "OpenAI: header Authorization Bearer")
    check(cap["body"]["input"] == "PROMPT-O",
          "OpenAI: prompt nel campo input")
    check("reasoning" not in cap["body"],
          "OpenAI base: nessun reasoning")
    check(out == "risposta openai",
          "OpenAI: testo ricostruito da output[*].content[*]")

    print("== OpenAI (avanzato) ==")
    cap = {}
    install_capture({"output": []}, cap)
    make_worker("openai", "avanzato")._run_openai("P")
    check(cap["body"]["reasoning"] == {"effort": "high"},
          "OpenAI avanzato: reasoning.effort=high")


def test_anthropic():
    print("== Anthropic (base) ==")
    cap = {}
    install_capture(
        {"content": [
            {"type": "text", "text": "ciao "},
            {"type": "text", "text": "claude"},
            {"type": "thinking", "text": "IGNORA"},
        ]},
        cap,
    )
    w = make_worker("anthropic")
    out = w._run_anthropic("PROMPT-A")

    check(cap["url"] == "https://api.anthropic.com/v1/messages",
          "Anthropic: endpoint Messages corretto")
    check(cap["headers"].get("X-api-key") == "FAKE-KEY-123",
          "Anthropic: header x-api-key")
    check(cap["headers"].get("Anthropic-version") == "2023-06-01",
          "Anthropic: header anthropic-version")
    check(cap["body"]["messages"][0]["content"] == "PROMPT-A",
          "Anthropic: prompt nel messaggio")
    check("thinking" not in cap["body"],
          "Anthropic base: nessun thinking")
    check(out == "ciao claude",
          "Anthropic: concatena solo i blocchi 'text'")

    print("== Anthropic (avanzato) ==")
    cap = {}
    install_capture({"content": []}, cap)
    make_worker("anthropic", "avanzato")._run_anthropic("P")
    check(cap["body"]["max_tokens"] == 8192,
          "Anthropic avanzato: max_tokens=8192")
    check(cap["body"]["thinking"]["type"] == "enabled",
          "Anthropic avanzato: thinking abilitato")


def test_http_error():
    print("== Errore HTTP leggibile ==")

    def fake_urlopen(req, timeout=None):
        raise urllib.error.HTTPError(
            req.full_url, 401, "Unauthorized", {},
            io.BytesIO(b'{"error":"chiave non valida"}'),
        )
    urllib.request.urlopen = fake_urlopen

    w = make_worker("gemini")
    try:
        w._run_gemini("P")
        check(False, "doveva sollevare un errore")
    except RuntimeError as e:
        msg = str(e)
        check("401" in msg and "chiave non valida" in msg,
              "Errore HTTP: codice e dettaglio del provider nel messaggio")


def test_no_sdk_imports():
    print("== Nessun import SDK residuo ==")
    import inspect
    src = inspect.getsource(main.GeminiWorker)
    check("from google import genai" not in src,
          "Nessun import google.genai nella classe")
    check("from openai import" not in src,
          "Nessun import openai nella classe")
    check("from anthropic import" not in src,
          "Nessun import anthropic nella classe")


def main_test():
    _orig = urllib.request.urlopen
    try:
        test_gemini()
        test_openai()
        test_anthropic()
        test_http_error()
        test_no_sdk_imports()
    finally:
        urllib.request.urlopen = _orig

    print()
    if failures:
        print(f"RISULTATO: {len(failures)} test FALLITI")
        return 1
    print("RISULTATO: tutti i test PASSATI")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main_test())
