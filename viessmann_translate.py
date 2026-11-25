#!/usr/bin/env python3
"""
Viessmann Excel parameter translator.

Workflow:
  1. Fill EnglishTranslation in translation_review.csv for tokens to override.
  2. Add/adjust dictionary.json.
  3. Run the translator to create the translated Excel.

For automatic translation with Google Cloud Translation v3 (default when
--project or GOOGLE_CLOUD_PROJECT is set, falling back to v2 otherwise):
  python viessmann_translate.py --provider google --project <PROJECT_ID>

You can also force v3 or v2 explicitly with --provider google-v3 / google-v2.

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

def sanitize_identifier(text):
    """Normalize translated text into an underscore-separated identifier."""
    return re.sub(r"[^A-Za-z0-9]+", "_", str(text)).strip("_")

def resolve_google_project(explicit_project=None):
    """Resolve Google Cloud project ID from CLI arg, env vars, or ADC."""
    if explicit_project and explicit_project.strip():
        return explicit_project.strip()
    for var in ("GOOGLE_CLOUD_PROJECT", "GCLOUD_PROJECT", "CLOUDSDK_CORE_PROJECT"):
        val = os.environ.get(var, "").strip()
        if val:
            return val
    try:
        import google.auth
        _, project_id = google.auth.default()
        if project_id:
            return project_id
    except Exception:
        pass
    return None

def _build_google_v3_translator(project, location):
    """Create a token translator backed by Google Cloud Translation v3."""
    try:
        from google.cloud import translate_v3
    except ImportError:
        raise SystemExit(
            "Google provider selected but google-cloud-translate is not installed. "
            "Install with: pip install google-cloud-translate"
        )
    client = translate_v3.TranslationServiceClient()
    parent = f"projects/{project}/locations/{location}"
    print(f"Google Cloud Translation API: v3 ({parent})")

    def translate(token):
        resp = client.translate_text(
            request={
                "parent": parent,
                "contents": [token],
                "mime_type": "text/plain",
                "source_language_code": "de",
                "target_language_code": "en",
            }
        )
        return resp.translations[0].translated_text

    return translate

def _build_google_v2_translator():
    """Create a token translator backed by Google Cloud Translation v2."""
    try:
        from google.cloud import translate_v2 as translate
    except ImportError:
        raise SystemExit(
            "Google provider selected but google-cloud-translate is not installed. "
            "Install with: pip install google-cloud-translate"
        )
    google_client = translate.Client()
    print("Google Cloud Translation API: v2")

    def translate(token):
        return google_client.translate(
            token, source_language="de", target_language="en"
        )["translatedText"]

    return translate

def _build_deepl_translator():
    """Create a token translator backed by DeepL."""
    try:
        import deepl
    except ImportError:
        raise SystemExit("DeepL provider selected but deepl is not installed.")
    auth_key = os.environ.get("DEEPL_API_KEY", "").strip()
    if not auth_key:
        raise SystemExit("DeepL provider selected but DEEPL_API_KEY is not set.")
    deepl_client = deepl.DeepLClient(auth_key)

    def translate(token):
        return deepl_client.translate_text(
            token, source_lang="de", target_lang="en-us"
        )

    return translate

def build_token_translator(
    dictionary, review, provider, project=None, location="global", delay=0.0
):
    merged = {**dictionary, **review}
    cache = {}
    api_translate = None

    if provider == "google-v3":
        resolved_project = resolve_google_project(project)
        if not resolved_project:
            raise SystemExit(
                "Google Cloud Translation v3 requires a project ID. "
                "Pass --project <PROJECT_ID> or set GOOGLE_CLOUD_PROJECT."
            )
        api_translate = _build_google_v3_translator(resolved_project, location)
    elif provider == "google-v2":
        api_translate = _build_google_v2_translator()
    elif provider == "google":
        resolved_project = resolve_google_project(project)
        if resolved_project:
            api_translate = _build_google_v3_translator(resolved_project, location)
        else:
            api_translate = _build_google_v2_translator()
    elif provider == "deepl":
        api_translate = _build_deepl_translator()
    elif provider != "none":
        raise ValueError(provider)

    def translate_token(token):
        if token in merged:
            return merged[token]
        if token in cache:
            return cache[token]
        if len(token) == 0:
            return ""

        if provider == "none":
            result = token
        else:
            result = sanitize_identifier(api_translate(token))
            if delay > 0:
                time.sleep(delay)

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
    ap.add_argument(
        "--provider",
        choices=["none", "google", "google-v2", "google-v3", "deepl"],
        default="none",
    )
    ap.add_argument(
        "--project",
        default=None,
        help="Google Cloud project ID for Cloud Translation v3.",
    )
    ap.add_argument(
        "--location",
        default="global",
        help="Google Cloud location for Translation v3 (default: global).",
    )
    ap.add_argument("--delay", type=float, default=0.05)
    args = ap.parse_args()

    src = Path(args.input)
    dictionary = load_dictionary(Path(args.dictionary))
    review = load_review(Path(args.review))
    
    print(f"Using translator: {args.provider}")
    translate_token = build_token_translator(
        dictionary,
        review,
        args.provider,
        project=args.project,
        location=args.location,
        delay=args.delay,
    )

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
                    if args.provider != "none" and (total % 100) == 0:
                        print(f"translated {total} params")

    wb.save(args.output)
    print(f"Processed: {total}")
    print(f"Changed:   {changed}")
    print(f"Output:    {args.output}")

if __name__ == "__main__":
    main()
