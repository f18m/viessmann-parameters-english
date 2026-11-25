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
KEEP_ORIGINAL_TRUTHY = {"1", "true", "yes", "y", "x"}
CAMEL_SKIP_PREFIXES = ("nvi", "nvo", "nci", "Vito", "resProd", "valmRPC", "vlogRPC")

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
    parts = [p for p in re.split(r"[._]+", core) if p]
    return parts, suffix

def split_camel_case(token):
    """Split a CamelCase identifier into component words."""
    s = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", token)
    s = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", s)
    return [p for p in s.split() if p]

def load_review(path):
    """Load EnglishTranslation overrides and KeepOriginal flags from CSV."""
    overrides = {}
    keep_original = set()
    if not path.exists():
        return overrides, keep_original
    with path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            token = row.get("GermanToken", "").strip()
            if not token:
                continue
            keep = row.get("KeepOriginal", "").strip().lower()
            tr = row.get("EnglishTranslation", "").strip()
            if keep in KEEP_ORIGINAL_TRUTHY:
                keep_original.add(token)
            elif tr:
                overrides[token] = tr
    return overrides, keep_original

def discover_parameter_columns(ws):
    cols, _ = discover_parameter_target(ws)
    return cols

def discover_parameter_target(ws):
    """Return (column_indices, start_row) for a worksheet's parameter column(s)."""
    try:
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    except StopIteration:
        return [], 2

    header_cols = []
    for i, h in enumerate(header):
        h_norm = str(h).strip().lower() if h is not None else ""
        if h_norm in PARAM_HEADERS or "parameter" in h_norm:
            header_cols.append(i)
    if header_cols:
        return header_cols, 2

    # Fallback for headerless parameter sheets (e.g., V200KW1_6, V200KW2_6, VPlusHO1_40).
    col_hits = {}
    for row in ws.iter_rows(min_row=1, max_row=50, values_only=True):
        for i, val in enumerate(row):
            if isinstance(val, str) and re.search(r"~0x[0-9A-Fa-f]+$", val.strip()):
                col_hits[i] = col_hits.get(i, 0) + 1
    inferred_cols = sorted(i for i, count in col_hits.items() if count >= 3)
    if inferred_cols:
        return inferred_cols, 1
    return [], 2

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

def build_lookup_tables(dictionary, review, keep_original=None):
    """Build exact and case-insensitive override/keep lookup tables."""
    merged = {**dictionary, **review}
    merged_lower = {}
    for k, v in merged.items():
        merged_lower.setdefault(k.lower(), v)
    keep_exact = set(keep_original or ())
    keep_lower = {k.lower() for k in keep_exact}
    return merged, merged_lower, keep_exact, keep_lower

def lookup_override(token, merged, merged_lower, keep_exact, keep_lower):
    """Return an explicit override/keep value for token, or None if not overridden."""
    if token in keep_exact or token.lower() in keep_lower:
        return token
    if token in merged:
        return merged[token]
    return merged_lower.get(token.lower())

def build_token_translator(
    dictionary,
    review,
    provider,
    project=None,
    location="global",
    delay=0.0,
    keep_original=None,
):
    merged, merged_lower, keep_exact, keep_lower = build_lookup_tables(
        dictionary, review, keep_original
    )
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

    def _translate_atomic(token):
        if len(token) == 0:
            return ""
        if is_technical_token(token):
            return token
        override = lookup_override(token, merged, merged_lower, keep_exact, keep_lower)
        if override is not None:
            return override
        if token in cache:
            return cache[token]
        if provider == "none":
            return token
        result = sanitize_identifier(api_translate(token))
        if delay > 0:
            time.sleep(delay)
        cache[token] = result
        return result

    def translate_token(token):
        if len(token) == 0:
            return ""
        override = lookup_override(token, merged, merged_lower, keep_exact, keep_lower)
        if override is not None:
            return override
        if token in cache:
            return cache[token]
        if provider == "none":
            cache[token] = token
            return token

        san_orig = sanitize_identifier(token)
        result = _translate_atomic(token)

        # Fallback: if API echoed a CamelCase compound unchanged, split and translate parts.
        if (
            result == san_orig
            and len(token) > 4
            and not token.startswith(CAMEL_SKIP_PREFIXES)
        ):
            subparts = split_camel_case(token)
            if len(subparts) > 1:
                sub_results = [_translate_atomic(sub) for sub in subparts]
                combined = "_".join(r for r in sub_results if r)
                if combined and combined != san_orig:
                    result = combined

        cache[token] = result
        return result

    return translate_token

def is_technical_token(part):
    """Return True for technical identifiers or numeric fragments preserved as-is."""
    return bool(re.fullmatch(r"(?:[A-Z0-9]+|[0-9]+)", part))

def translate_parameter(s, translate_token):
    parts, suffix = split_parameter(s)
    translated = []

    for part in parts:
        # Preserve obvious technical identifiers and numeric fragments.
        if is_technical_token(part):
            translated.append(part)
        else:
            translated.append(translate_token(part))

    return "_".join(translated) + suffix

def collect_untranslated(
    sheet_name,
    old_val,
    new_val,
    merged,
    token_stats,
    merged_lower=None,
    keep_exact=None,
    keep_lower=None,
):
    """Record non-technical, non-dictionary tokens that remained unchanged."""
    sparts, _ = split_parameter(old_val)
    dparts, _ = split_parameter(new_val) if isinstance(new_val, str) else ([], "")
    merged_lower = merged_lower or {}
    keep_exact = keep_exact or set()
    keep_lower = keep_lower or set()
    found = False

    for sp in sparts:
        if not sp or is_technical_token(sp):
            continue
        if lookup_override(sp, merged, merged_lower, keep_exact, keep_lower) is not None:
            continue
        san = sanitize_identifier(sp)
        if sp in dparts or (san and san in dparts):
            found = True
            entry = token_stats.setdefault(
                sp,
                {
                    "count": 0,
                    "sheets": set(),
                    "example_src": old_val,
                    "example_dst": new_val or "",
                },
            )
            entry["count"] += 1
            entry["sheets"].add(sheet_name)

    return found

def write_untranslated_report(path, token_stats):
    """Write untranslated tokens sorted by frequency descending."""
    rows = sorted(
        token_stats.items(),
        key=lambda item: (-item[1]["count"], item[0].lower(), item[0]),
    )
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Token", "Occurrences", "Sheets", "ExampleSource", "ExampleTranslated"])
        for token, info in rows:
            writer.writerow([
                token,
                info["count"],
                ", ".join(sorted(info["sheets"])),
                info["example_src"],
                info["example_dst"],
            ])

def run_report_only(src_path, dst_path, merged, merged_lower, keep_exact, keep_lower):
    """Compare existing input and translated workbooks without calling APIs."""
    src_wb = load_workbook(src_path, read_only=True)
    dst_wb = load_workbook(dst_path, read_only=True)
    total = 0
    changed = 0
    unchanged = 0
    params_with_untranslated = 0
    token_stats = {}
    skipped_sheets = []

    for name in src_wb.sheetnames:
        sws = src_wb[name]
        cols, start_row = discover_parameter_target(sws)
        if not cols:
            skipped_sheets.append(name)
            continue
        if name not in dst_wb.sheetnames:
            continue
        dws = dst_wb[name]
        for sr, dr in zip(
            sws.iter_rows(min_row=start_row, values_only=True),
            dws.iter_rows(min_row=start_row, values_only=True),
        ):
            for col_idx in cols:
                old = sr[col_idx] if col_idx < len(sr) else None
                new = dr[col_idx] if col_idx < len(dr) else None
                if not (isinstance(old, str) and old.strip()):
                    continue
                total += 1
                if new != old:
                    changed += 1
                else:
                    unchanged += 1
                if collect_untranslated(
                    name, old, new, merged, token_stats, merged_lower, keep_exact, keep_lower
                ):
                    params_with_untranslated += 1

    return total, changed, unchanged, params_with_untranslated, token_stats, skipped_sheets

def print_summary(
    total,
    changed,
    unchanged,
    params_with_untranslated=None,
    token_stats=None,
    skipped_sheets=None,
    untranslated_path=None,
    output_path=None,
):
    print(f"Processed: {total}")
    print(f"Changed:   {changed}")
    print(f"Unchanged: {unchanged}")
    if token_stats is not None:
        print(f"Params with untranslated tokens: {params_with_untranslated}")
        print(f"Untranslated unique tokens:      {len(token_stats)}")
    if skipped_sheets:
        print(f"Skipped sheets (no parameter column detected): {', '.join(skipped_sheets)}")
    if untranslated_path is not None:
        print(f"Untranslated report: {untranslated_path}")
    if output_path is not None:
        print(f"Output:    {output_path}")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="Viessmann.Parameter.Adresses.xlsx")
    ap.add_argument("--output", default="Viessmann_Parameter_Adresses_EN_translated.xlsx")
    ap.add_argument("--dictionary", default="dictionary.json")
    ap.add_argument("--review", default="translation_review.csv")
    ap.add_argument("--untranslated", default="untranslated_tokens.csv")
    ap.add_argument(
        "--report-only",
        action="store_true",
        help="Compare --input and --output to report untranslated tokens without translating.",
    )
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
    review, keep_original = load_review(Path(args.review))
    merged, merged_lower, keep_exact, keep_lower = build_lookup_tables(
        dictionary, review, keep_original
    )

    if args.report_only:
        total, changed, unchanged, params_with_untranslated, token_stats, skipped_sheets = (
            run_report_only(
                src, Path(args.output), merged, merged_lower, keep_exact, keep_lower
            )
        )
        write_untranslated_report(Path(args.untranslated), token_stats)
        print_summary(
            total,
            changed,
            unchanged,
            params_with_untranslated=params_with_untranslated,
            token_stats=token_stats,
            skipped_sheets=skipped_sheets,
            untranslated_path=args.untranslated,
        )
        return

    print(f"Using translator: {args.provider}")
    translate_token = build_token_translator(
        dictionary,
        review,
        args.provider,
        project=args.project,
        location=args.location,
        delay=args.delay,
        keep_original=keep_original,
    )

    wb = load_workbook(src)
    changed = 0
    unchanged = 0
    total = 0
    params_with_untranslated = 0
    token_stats = {}
    skipped_sheets = []

    for ws in wb.worksheets:
        cols, start_row = discover_parameter_target(ws)
        if not cols:
            print(f"Processing sheet: {ws.title} (skipped: no parameter column detected)", flush=True)
            skipped_sheets.append(ws.title)
            continue
        print(f"Processing sheet: {ws.title}", flush=True)
        for col_idx in cols:
            for row in ws.iter_rows(min_row=start_row):
                if col_idx >= len(row):
                    continue
                cell = row[col_idx]
                if isinstance(cell.value, str) and cell.value.strip():
                    total += 1
                    old = cell.value
                    new = translate_parameter(old, translate_token)
                    if new != old:
                        cell.value = new
                        changed += 1
                    else:
                        unchanged += 1
                    if args.provider != "none":
                        if collect_untranslated(
                            ws.title,
                            old,
                            new,
                            merged,
                            token_stats,
                            merged_lower,
                            keep_exact,
                            keep_lower,
                        ):
                            params_with_untranslated += 1
                        if (total % 100) == 0:
                            print(f"translated {total} params", flush=True)

    wb.save(args.output)
    untranslated_path = None
    if args.provider != "none":
        write_untranslated_report(Path(args.untranslated), token_stats)
        untranslated_path = args.untranslated
    print_summary(
        total,
        changed,
        unchanged,
        params_with_untranslated=params_with_untranslated if args.provider != "none" else None,
        token_stats=token_stats if args.provider != "none" else None,
        skipped_sheets=skipped_sheets,
        untranslated_path=untranslated_path,
        output_path=args.output,
    )

if __name__ == "__main__":
    main()
