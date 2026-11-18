# Scheduled constituent refresh: fetch all eligible indices and preserve last-known-good lists.
import ast
import csv
import io
import json
import os
import re
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

import requests

SUPABASE_URL = os.environ["SUPABASE_URL"].rstrip("/")
SUPABASE_SERVICE_KEY = os.environ["SUPABASE_SERVICE_KEY"]
BUCKET = "excel-files"
MANIFEST_PATH = "index-constituents/manifest.json"
ARCHIVE_PATH = "index-constituents/constituents.json"
REFRESH_DAYS = 15
TIMEOUT = 25

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131 Safari/537.36",
    "Accept": "text/csv,application/json,text/plain,*/*",
    "Referer": "https://www.niftyindices.com/",
})

def now_iso():
    return datetime.now(timezone.utc).isoformat()

def clean_name(value):
    return " ".join(re.sub(r"[^A-Z0-9]+", " ", str(value).upper()).split())

def storage_url(path):
    return f"{SUPABASE_URL}/storage/v1/object/{BUCKET}/{quote(path, safe='/')}"

def storage_headers(content_type="application/json"):
    return {
        "apikey": SUPABASE_SERVICE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "Content-Type": content_type,
        "x-upsert": "true",
    }

def storage_download(path):
    response = SESSION.get(storage_url(path), headers=storage_headers(), timeout=TIMEOUT)
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return response.content

def storage_upload(path, payload):
    response = SESSION.put(
        storage_url(path),
        headers=storage_headers(),
        data=json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"),
        timeout=TIMEOUT,
    )
    response.raise_for_status()

def load_official_filename_map():
    # Read the same official filename map maintained in app.py to avoid drift.
    try:
        with open("app.py", "r", encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                if any(isinstance(t, ast.Name) and t.id == "official_files" for t in node.targets):
                    value = ast.literal_eval(node.value)
                    if isinstance(value, dict) and value:
                        return value
    except Exception as exc:
        print(f"Could not read app.py filename map: {exc}")
    return {}

def csv_members(content):
    text = content.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        return []
    symbol_column = None
    company_column = None
    industry_column = None
    for column in reader.fieldnames:
        key = re.sub(r"[^a-z]", "", str(column).lower())
        if key in {"symbol", "nse symbol", "tradingsymbol", "securitysymbol"} and symbol_column is None:
            symbol_column = column
        if key in {"companyname", "company", "securityname"}:
            company_column = column
        if key in {"industry", "industryname"}:
            industry_column = column
    if symbol_column is None:
        for column in reader.fieldnames:
            if "symbol" in str(column).lower():
                symbol_column = column
                break
    if symbol_column is None:
        return []
    members = []
    seen = set()
    for row in reader:
        symbol = str(row.get(symbol_column, "") or "").strip()
        if not symbol or symbol.lower() in {"nan", "symbol"}:
            continue
        normalized = symbol.upper()
        if normalized in seen:
            continue
        seen.add(normalized)
        members.append({
            "No.": len(members) + 1,
            "NSE Symbol": symbol,
            "Company": str(row.get(company_column, "") or "").strip() if company_column else "",
            "Industry": str(row.get(industry_column, "") or "").strip() if industry_column else "",
        })
    return members

def api_members(index_name):
    base = "https://api.dharun.dev/v1/indices"
    response = SESSION.get(base, timeout=TIMEOUT)
    response.raise_for_status()
    items = response.json().get("data", [])
    target = clean_name(index_name)
    match = next((item for item in items if clean_name(item.get("name", "")) == target), None)
    if not match:
        return []
    detail = SESSION.get(f"{base}/{match['id']}", timeout=TIMEOUT)
    detail.raise_for_status()
    rows = detail.json().get("data", {}).get("members", [])
    members = []
    for row in rows:
        meta = row.get("meta") or {}
        symbol = row.get("symbol") or row.get("Symbol") or row.get("SYMBOL")
        if not symbol:
            continue
        members.append({
            "No.": len(members) + 1,
            "NSE Symbol": str(symbol).strip(),
            "Company": str(row.get("company") or row.get("Company Name") or meta.get("companyName") or ""),
            "Industry": str(row.get("industry") or meta.get("industry") or ""),
        })
    return members

def nse_api_members(index_name):
    session = requests.Session()
    session.headers.update(SESSION.headers)
    session.get("https://www.nseindia.com/", timeout=TIMEOUT)
    response = session.get(
        "https://www.nseindia.com/api/equity-stockIndices",
        params={"index": " ".join(index_name.upper().split())},
        headers={"Accept": "application/json, text/plain, */*"},
        timeout=TIMEOUT,
    )
    response.raise_for_status()
    members = []
    for row in response.json().get("data", []):
        symbol = row.get("symbol") or row.get("Symbol")
        if not symbol or clean_name(symbol) == clean_name(index_name):
            continue
        members.append({
            "No.": len(members) + 1,
            "NSE Symbol": str(symbol).strip(),
            "Company": str(row.get("companyName") or row.get("meta", {}).get("companyName") or ""),
            "Industry": str(row.get("industry") or ""),
        })
    return members

def official_csv_members(index_name, filename_map):
    filename = filename_map.get(index_name)
    slug = re.sub(r"[^a-z0-9]+", "", index_name.lower())
    candidates = []
    if filename:
        candidates.append(filename)
    candidates.extend([
        f"ind_{slug}list.csv",
        f"ind_{slug}_list.csv",
        f"ind_{slug.lower()}List.csv",
    ])
    hosts = [
        "https://www.niftyindices.com/IndexConstituent/",
        "https://niftyindices.com/IndexConstituent/",
        "https://nsearchives.nseindia.com/content/indices/",
    ]
    errors = []
    for candidate in dict.fromkeys(candidates):
        for host in hosts:
            url = host + candidate
            try:
                response = SESSION.get(url, timeout=TIMEOUT)
                response.raise_for_status()
                members = csv_members(response.content)
                if members:
                    return members, url
                errors.append(f"{url}: no symbol rows")
            except Exception as exc:
                errors.append(f"{url}: {exc}")
    raise RuntimeError("No official constituent CSV found. " + " | ".join(errors[-3:]))

def fetch_members(index_name, filename_map):
    errors = []
    for label, fetcher in (
        ("India API", lambda: (api_members(index_name), "India API archive")),
        ("NSE API", lambda: (nse_api_members(index_name), "NSE equity-stockIndices API")),
    ):
        try:
            members, source = fetcher()
            if members:
                return members, source
            errors.append(f"{label} returned no components")
        except Exception as exc:
            errors.append(f"{label}: {exc}")
    members, source = official_csv_members(index_name, filename_map)
    return members, source

def main():
    raw_manifest = storage_download(MANIFEST_PATH)
    if not raw_manifest:
        print("No index manifest found yet. Upload an index file in the Streamlit dashboard first.")
        return
    manifest = json.loads(raw_manifest.decode("utf-8"))
    index_names = manifest.get("index_names", [])
    if not index_names:
        print("Manifest has no indices above 10,000.")
        return

    raw_archive = storage_download(ARCHIVE_PATH)
    archive = json.loads(raw_archive.decode("utf-8")) if raw_archive else {"indices": {}}
    archive.setdefault("indices", {})
    filename_map = load_official_filename_map()
    now = datetime.now(timezone.utc)
    refreshed = 0
    unchanged = 0
    unavailable = 0

    for index_name in index_names:
        key = clean_name(index_name)
        previous = archive["indices"].get(key, {})
        last_checked = previous.get("last_verified_at")
        due = True
        if last_checked:
            try:
                checked_at = datetime.fromisoformat(last_checked.replace("Z", "+00:00"))
                due = now - checked_at >= timedelta(days=REFRESH_DAYS)
            except Exception:
                due = True
        if not due:
            continue

        print(f"Checking constituents: {index_name}")
        try:
            members, source = fetch_members(index_name, filename_map)
            old_symbols = [x.get("NSE Symbol") for x in previous.get("members", [])]
            new_symbols = [x.get("NSE Symbol") for x in members]
            changed = old_symbols != new_symbols
            record = dict(previous)
            record.update({
                "index_name": index_name,
                "members": members,
                "status": "available",
                "source": source,
                "last_verified_at": now_iso(),
            })
            if changed or not previous.get("last_updated_at"):
                record["last_updated_at"] = now_iso()
                if old_symbols:
                    record["previous_symbols"] = old_symbols
            archive["indices"][key] = record
            if changed:
                refreshed += 1
            else:
                unchanged += 1
            print(f"  OK: {len(members)} stocks; {'changed' if changed else 'unchanged'}")
        except Exception as exc:
            record = dict(previous)
            record.update({
                "index_name": index_name,
                "status": "stale-preserved" if previous.get("members") else "unavailable",
                # Do not advance last_verified_at after a failed download.
                # This keeps failed/new indices eligible for retry on the next
                # daily workflow run instead of waiting another 15 days.
                "last_error": str(exc)[:1200],
                "last_failed_at": now_iso(),
            })
            archive["indices"][key] = record
            unavailable += 1
            print(f"  SKIP: {exc}")

    archive["last_run_at"] = now_iso()
    archive["manifest_updated_at"] = manifest.get("updated_at")
    storage_upload(ARCHIVE_PATH, archive)
    print(f"Finished. Updated/verified: {refreshed}; unavailable: {unavailable}; unchanged or not due: {unchanged}.")
    
if __name__ == "__main__":
    main()
