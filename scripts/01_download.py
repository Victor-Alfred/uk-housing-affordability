"""
Download ONS housing datasets. Nothing is transformed here.

Three things make this worth more than a curl command:

1. The file URL is resolved from the dataset page at run time. ONS renames
   the workbook between editions, so a hardcoded URL is a pipeline with a
   six-month fuse.

2. ONS uses two publishing patterns and this handles both. Some datasets
   list dated editions on the page (housebuilding, newbuild_sales). Others
   publish a single /current/ file and keep history on a separate
   "Previous versions" page (both affordability datasets). --all-editions
   covers the first, --previous the second.

3. Every download is recorded in data/raw/manifest.jsonl with a SHA-256,
   the resolved URL, the edition label and a UTC timestamp - plus, for
   superseded files, ONS's own reason for the update.

Point 3 is what makes the October 2025 correction usable. ONS swapped
existing-dwellings estimates with all-dwellings estimates in the lower
quartile prices (tables 2a/4a/6a) and ratios (2c/4c/6c). Version v10 is
the faulty file, v11 the corrected replacement. That pair is a real
regression fixture, not a planted one.

Usage:
    python scripts/01_download.py --all --dry-run
    python scripts/01_download.py --all
    python scripts/01_download.py --dataset affordability --previous
    python scripts/01_download.py --dataset housebuilding --all-editions
"""

import argparse
import hashlib
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
MANIFEST = RAW / "manifest.jsonl"

ONS = "https://www.ons.gov.uk"
TIMEOUT = 120
CHUNK = 1 << 20  # 1 MiB

# Source registry. Page URLs, never file URLs - see the docstring.
DATASETS = {
    "affordability": {
        "title": "House price to residence-based earnings ratio (LQ and median)",
        "url": f"{ONS}/peoplepopulationandcommunity/housing/datasets/"
               "ratioofhousepricetoresidencebasedearningslowerquartileandmedian",
    },
    "housebuilding": {
        "title": "House building: dwellings started and completed by local authority",
        "url": f"{ONS}/peoplepopulationandcommunity/housing/datasets/"
               "housebuildingukpermanentdwellingsstartedandcompletedbylocalauthority",
    },
    "affordability_newbuild": {
        "title": "House price (newly built) to residence-based earnings ratio",
        "url": f"{ONS}/peoplepopulationandcommunity/housing/datasets/"
               "housepricenewlybuiltdwellingstoresidencebasedearningsratio",
    },
    "newbuild_sales": {
        "title": "Residential property sales, admin geographies (newly built)",
        "url": f"{ONS}/peoplepopulationandcommunity/housing/datasets/"
               "residentialpropertysalesforadministrativegeographiesnewlybuiltdwellings",
    },
}


def slug(text):
    """Filesystem-safe directory name from an edition label."""
    text = re.sub(r"[^\w\s-]", "", text.lower()).strip()
    return re.sub(r"[\s_-]+", "-", text)[:80] or "unlabelled"


def is_spreadsheet(href):
    return "/file?uri=" in href and re.search(r"\.xlsx?(\?|$)", href, re.IGNORECASE)


def find_editions(html, page_url):
    """Every .xlsx download advertised on the dataset page itself.

    Anchored on the links rather than the headings, because heading levels
    vary between dataset types. The label comes from walking backwards to
    the nearest heading mentioning 'edition'. Page order is publication
    order, newest first.
    """
    soup = BeautifulSoup(html, "html.parser")
    found, seen = [], set()

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not is_spreadsheet(href):
            continue

        url = urljoin(ONS, href)
        if url in seen:
            continue
        seen.add(url)

        label = "unlabelled"
        for prev in a.find_all_previous(["h2", "h3", "h4", "h5", "strong"]):
            text = prev.get_text(strip=True)
            if text and "edition" in text.lower():
                label = text
                break

        found.append({
            "label": label,
            "url": url,
            "filename": href.rsplit("/", 1)[-1].split("?")[0],
        })

    if not found:
        raise RuntimeError(
            f"No .xlsx links found on {page_url}. The page structure has "
            "probably changed - look at it before loosening this parser."
        )
    return found


def find_previous_versions(page_html, session):
    """Follow the 'Previous versions' link and parse the superseded table.

    Columns are: superseded file, reason for update, date superseded. The
    reason column is ONS's own change log - worth capturing, because it
    says which revisions were scheduled and which were corrections.

    Returns [] for datasets that don't use this pattern.
    """
    soup = BeautifulSoup(page_html, "html.parser")

    link = None
    for a in soup.find_all("a", href=True):
        if "previous version" in a.get_text(strip=True).lower():
            link = urljoin(ONS, a["href"])
            break
    if link is None:
        return []

    resp = session.get(link, timeout=TIMEOUT)
    resp.raise_for_status()

    out = []
    for row in BeautifulSoup(resp.text, "html.parser").find_all("tr"):
        a = row.find("a", href=True)
        if not a or not is_spreadsheet(a["href"]):
            continue

        cells = [td.get_text(" ", strip=True) for td in row.find_all("td")]
        url = urljoin(ONS, a["href"])
        m = re.search(r"/previous/(v\d+)/", url)

        out.append({
            "label": f"{m.group(1) if m else 'unknown'}-superseded-"
                     f"{cells[2] if len(cells) > 2 else ''}",
            "url": url,
            "filename": a["href"].rsplit("/", 1)[-1].split("?")[0],
            "reason": cells[1] if len(cells) > 1 else "",
            "superseded": cells[2] if len(cells) > 2 else "",
        })
    return out


def read_manifest():
    if not MANIFEST.exists():
        return []
    with MANIFEST.open() as f:
        return [json.loads(line) for line in f if line.strip()]


def download(edition, key, session, manifest):
    """Fetch one file. Streams to a .part while hashing, so a truncated
    download never masquerades as a complete one. Skips anything whose
    hash is already held."""
    target_dir = RAW / key / slug(edition["label"])
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / edition["filename"]
    tmp = target.with_suffix(target.suffix + ".part")

    digest, size = hashlib.sha256(), 0
    with session.get(edition["url"], stream=True, timeout=TIMEOUT) as r:
        r.raise_for_status()
        with tmp.open("wb") as f:
            for chunk in r.iter_content(chunk_size=CHUNK):
                if chunk:
                    f.write(chunk)
                    digest.update(chunk)
                    size += len(chunk)

    sha = digest.hexdigest()
    if sha in {row["sha256"] for row in manifest}:
        tmp.unlink(missing_ok=True)
        print(f"    skip   {edition['filename']} (already held)")
        return

    tmp.replace(target)

    record = {
        "dataset": key,
        "edition": edition["label"],
        "source_url": edition["url"],
        "filename": edition["filename"],
        "path": str(target.relative_to(ROOT)),
        "bytes": size,
        "sha256": sha,
        "retrieved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    # Present only for superseded files, where ONS states why it changed.
    for field in ("reason", "superseded"):
        if edition.get(field):
            record[field] = edition[field]

    MANIFEST.parent.mkdir(parents=True, exist_ok=True)
    with MANIFEST.open("a") as f:
        f.write(json.dumps(record, sort_keys=True) + "\n")
    manifest.append(record)
    print(f"    saved  {edition['filename']}  {size / 1e6:.2f} MB  sha {sha[:12]}")


def handle(edition, key, session, manifest, dry_run, note=""):
    print(f"  - {edition['label']}{note}")
    if dry_run:
        print(f"    would fetch {edition['url']}")
    else:
        download(edition, key, session, manifest)


def main():
    p = argparse.ArgumentParser()
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--dataset", choices=list(DATASETS))
    g.add_argument("--all", action="store_true")
    p.add_argument("--all-editions", action="store_true",
                   help="every edition listed on the dataset page, not just the latest")
    p.add_argument("--previous", action="store_true",
                   help="also fetch superseded files from the 'Previous versions' page")
    p.add_argument("--dry-run", action="store_true",
                   help="resolve URLs and print them, download nothing")
    args = p.parse_args()

    keys = list(DATASETS) if args.all else [args.dataset]
    manifest = read_manifest()
    session = requests.Session()
    session.headers["User-Agent"] = "uk-housing-affordability/0.1"

    failures = 0
    for key in keys:
        spec = DATASETS[key]
        print(f"\n{key}\n  {spec['title']}")
        try:
            resp = session.get(spec["url"], timeout=TIMEOUT)
            resp.raise_for_status()

            editions = find_editions(resp.text, spec["url"])
            print(f"  {len(editions)} edition(s) on the page")
            for ed in (editions if args.all_editions else editions[:1]):
                handle(ed, key, session, manifest, args.dry_run)

            if args.previous:
                prev = find_previous_versions(resp.text, session)
                print(f"  {len(prev)} superseded version(s)")
                for ed in prev:
                    note = f"  [{ed['reason'][:45]}]" if ed.get("reason") else ""
                    handle(ed, key, session, manifest, args.dry_run, note)

        except Exception as exc:
            failures += 1
            print(f"  FAILED: {exc}")

    print(f"\nmanifest: {MANIFEST}  ({len(manifest)} record(s))")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())