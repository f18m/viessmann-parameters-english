#!/usr/bin/env python3
"""
Viessmann Excel parameter translator.

Workflow:
  1. Run --extract to create translation_review.csv.
  2. Fill EnglishTranslation for tokens you want to override.
  3. Add/adjust dictionary.json.
  4. Run --translate to create the translated Excel.

For automatic translation, set GOOGLE_APPLICATION_CREDENTIALS and use:
  python viessmann_translate.py --translate --provider google

The program preserves technical prefixes, hexadecimal addresses and all
non-Parameter cells.
"""

import argparse, csv, json, os, re, time
from pathlib import Path
from openpyxl import load_workbook

PARAM_HEADERS = {"parameter", "parameter name", "parametername", "param", "name"}

def load_dictionary(path):
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))

def split_parameter(s):
    suffix = ""
    core = s
    m = re.search(r"(~0x[0-9A-Fa-f]+)$", core)
    if m:
        suffix = m.group(1)
        core = core[:m.start()]
    return core.split("_"), suffix

def load_review(path):
    d = {}
    if not path.exists():
        return d
    with path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            token = row.get("GermanToken", "").strip()
            tr = row.get("EnglishTranslation", "").strip()
            if token and tr:
                d[token] = tr
    return d

def discover_parameter_columns(ws):
    try:
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    except StopIteration:
        return []
    result = []
    for i, h in enumerate(header):
        h = str(h).strip().lower() if h is not None else ""
        if h in PARAM_HEADERS or "parameter" in h:
            result.append(i)
    return result

def build_token_translator(dictionary, review, provider):
    merged = {**dictionary, **review}
    cache = {}

    # Google Cloud Translation v3/v2 fallback through google-cloud-translate.
    google_client = None
    if provider == "google":
        try:
            from google.cloud import translate_v2 as translate
            google_client = translate.Client()
        except ImportError:
            raise SystemExit(
                "Google provider selected but google-cloud-translate is not installed. "
                "Install with: pip install google-cloud-translate"
            )
    elif provider == "deepl":
        try:
            import deepl
            auth_key = os.environ["DEEPL_API_KEY"]
            deepl_client = deepl.DeepLClient(auth_key)
        except ImportError:
            raise SystemExit(
                "DeepL provider selected but not working"
            )
        

    def translate_token(token):
        if token in merged:
            return merged[token]
        if token in cache:
            return cache[token]
        if len(token) == 0:
            return ""

        if provider == "none":
            result = token
        elif provider == "google":
            result = google_client.translate(token, source_language="de",
                                             target_language="en")["translatedText"]
            # Parameter names are identifiers: remove spaces/punctuation.
            result = re.sub(r"[^A-Za-z0-9]+", "_", result).strip("_")
        elif provider == "deepl":
            result = deepl_client.translate_text(token, source_lang="de",
                                             target_lang="en-us")
            # Parameter names are identifiers: remove spaces/punctuation.
            result = str(result)
            result = re.sub(r"[^A-Za-z0-9]+", "_", result).strip("_")
        else:
            raise ValueError(provider)

        cache[token] = result
        return result

    return translate_token

def translate_parameter(s, translate_token):
    parts, suffix = split_parameter(s)
    translated = []

    for part in parts:
        # Preserve obvious technical identifiers and numeric fragments.
        if re.fullmatch(r"(?:[A-Z0-9]+|[0-9]+)", part):
            translated.append(part)
        else:
            translated.append(translate_token(part))

    return "_".join(translated) + suffix

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="Viessmann_Parameter_Adresses_EN.xlsx")
    ap.add_argument("--output", default="Viessmann_Parameter_Adresses_EN_translated.xlsx")
    ap.add_argument("--dictionary", default="dictionary.json")
    ap.add_argument("--review", default="translation_review.csv")
    ap.add_argument("--provider", choices=["none", "google", "deepl"], default="none")
    ap.add_argument("--delay", type=float, default=0.05)
    args = ap.parse_args()

    src = Path(args.input)
    dictionary = load_dictionary(Path(args.dictionary))
    review = load_review(Path(args.review))
    
    print(f"Using translator: {args.provider}")
    translate_token = build_token_translator(dictionary, review, args.provider)

    wb = load_workbook(src)
    changed = 0
    total = 0

    for ws in wb.worksheets:
        for col_idx in discover_parameter_columns(ws):
            for row in ws.iter_rows(min_row=2):
                cell = row[col_idx]
                if isinstance(cell.value, str) and cell.value.strip():
                    total += 1
                    old = cell.value
                    new = translate_parameter(old, translate_token)
                    if new != old:
                        cell.value = new
                        changed += 1
                    if args.provider != "none":
                        time.sleep(args.delay)
                        if (total % 100) == 0:
                            print(f"translated {total} params")

    wb.save(args.output)
    print(f"Processed: {total}")
    print(f"Changed:   {changed}")
    print(f"Output:    {args.output}")

if __name__ == "__main__":
    main()
