"""
build.py — Fetches theNaughtyKettle's masterlist spreadsheet as CSV,
normalizes the data, and outputs audios.json for the site.

One spreadsheet, one output. Everything the site shows comes from the sheet,
so adding a row there is all that's needed to add an audio to the site.

Run manually:            python build.py
Against a local file:    python build.py --xlsx theNaughtyKettle_Audios_MERGED_fixed.xlsx
Run via GitHub Actions:  see .github/workflows/build.yml

To get the CSV URL: Google Sheets -> File -> Share -> Publish to web
-> pick the sheet -> Comma-separated values (.csv) -> Publish. Use the
resulting URL, which looks like:
  https://docs.google.com/spreadsheets/d/e/XXXXX/pub?gid=0&single=true&output=csv

Expected columns (extra columns are ignored, missing ones are tolerated):
  Title, Tags, Date, Writer, Script Link, Description, Post Link,
  Collab Partners, Duration, Editor, Type,
  Soundgasm, Audiochan, Hot Audio, Patreon, C4S
"""

import argparse
import csv
import io
import json
import os
import re
import sys
import unicodedata
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime, timezone

# ─── Source ──────────────────────────────────────────────────────────
SHEET_CSV_URL = os.environ.get("SHEET_CSV_URL", "YOUR_PUBLISHED_CSV_URL_HERE")
OUTPUT_FILE = os.environ.get("OUTPUT_FILE", "audios.json")

# ─── Cells that mean "nothing here" ──────────────────────────────────
# The sheet uses "x" for "not on this platform" / "no credit".
BLANK = {"", "x", "-", "--", "n/a", "na", "none", "#n/a", "#ref!", "?", "tbd"}

# ─── Type normalization ──────────────────────────────────────────────
# Matching is case-insensitive, so "improv", "Improv" and "IMPROV" all
# collapse to the canonical spelling. Add aliases as the sheet grows.
TYPE_CANONICAL = [
    "Ramblefap", "Improv", "Improv RP", "Script Fill", "Narrative",
    "Collab", "Big Collab", "Series", "Poetry", "Kettlecast",
    "For Fun", "Update", "Remix", "ASMR", "Interview",
]
TYPE_ALIASES = {
    "improv rp": "Improv RP",
    "improvrp": "Improv RP",
    "rp": "Improv RP",
    "scriptfill": "Script Fill",
    "script-fill": "Script Fill",
    "ramble": "Ramblefap",
    "ramblefaps": "Ramblefap",
    "big collabs": "Big Collab",
    "bigcollab": "Big Collab",
    "podcast": "Kettlecast",
    "oc": "Improv",
}

# ─── Link columns -> key in the JSON "links" object ──────────────────
# Order here is the order the buttons appear on a card.
LINK_COLUMNS = [
    ("post link", "reddit"),
    ("soundgasm", "soundgasm"),
    ("audiochan", "audiochan"),
    ("hot audio", "hotaudio"),
    ("patreon", "patreon"),
    ("subscribestar", "subscribestar"),
    ("c4s", "c4s"),
    ("script link", "script"),
]

EXCLUSIVE_RE = re.compile(r"^\s*[\[\(]\s*exclusive\s*[\]\)]\s*[-–:]?\s*", re.I)
SERIES_RE = re.compile(r"^(.{3,44}?)\s*#?\s*(\d+(?:\.\d+)?)\s*[:\-–]\s*\S")

EN_MONTHS = {m: f"{i:02d}" for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}
PT_MONTHS = {m: f"{i:02d}" for i, m in enumerate(
    ["jan", "fev", "mar", "abr", "mai", "jun",
     "jul", "ago", "set", "out", "nov", "dez"], start=1)}


# ─── Fetching ────────────────────────────────────────────────────────

def fetch_csv(url):
    """Download the published sheet and return a list of dicts."""
    print(f"  fetching {url[:90]}{'...' if len(url) > 90 else ''}")
    req = urllib.request.Request(url, headers={"User-Agent": "KettleBuildBot/1.0"})
    with urllib.request.urlopen(req, timeout=45) as resp:
        text = resp.read().decode("utf-8-sig")
    if text.lstrip()[:15].lower().startswith("<!doctype html") or "<html" in text[:400].lower():
        raise SystemExit(
            "The URL returned a web page, not CSV.\n"
            "Use the 'Publish to web' CSV link (…/pub?gid=0&single=true&output=csv),\n"
            "not the normal /edit link from the address bar."
        )
    return list(csv.DictReader(io.StringIO(text)))


def read_xlsx(path, sheet=None):
    """Read a local .xlsx instead of the published sheet (for offline runs)."""
    import openpyxl
    book = openpyxl.load_workbook(path, data_only=True)
    page = book[sheet] if sheet else book[book.sheetnames[0]]
    header = [c.value for c in page[1]]
    rows = []
    for values in page.iter_rows(min_row=2, values_only=True):
        if all(v in (None, "") for v in values):
            continue
        rows.append({(h or f"col{i}"): v for i, (h, v) in enumerate(zip(header, values))})
    return rows


# ─── Cell helpers ────────────────────────────────────────────────────

def row_getter(row):
    """Look up a column by name, ignoring case and surrounding whitespace."""
    lookup = {}
    for key, value in row.items():
        if key is None:
            continue
        lookup.setdefault(str(key).strip().lower(), value)

    def get(name):
        return clean(lookup.get(name.strip().lower()))
    return get


def clean(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    return "" if text.lower() in BLANK else text


def slugify(text, fallback="entry"):
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.category(c).startswith("So"))
    text = re.sub(r"[^a-zA-Z0-9]+", "-", text).strip("-").lower()
    return text[:60].strip("-") or fallback


def normalize_type(raw):
    if not raw:
        return "Uncategorized"
    key = re.sub(r"\s+", " ", raw.strip()).lower()
    if key in TYPE_ALIASES:
        return TYPE_ALIASES[key]
    for canonical in TYPE_CANONICAL:
        if canonical.lower() == key:
            return canonical
    return raw.strip()


# Audience tags look like F4M, F4A, FF4M, F4TF — speaker(s) on the left of
# the 4, listener on the right. Only the listener side drives the site's
# 4F / 4M / 4A filter, and a trans listener tag counts as its gender.
AUDIENCE_RE = re.compile(r"^([fma]{1,8})4([fmatnb]{1,8})$", re.I)


def audience_from_tags(tags):
    targets = set()
    for tag in tags:
        m = AUDIENCE_RE.fullmatch(tag.strip())
        if not m:
            continue
        listener = m.group(2).upper()
        if "A" in listener:
            targets.add("A")
        if "M" in listener:
            targets.add("M")
        if "F" in listener:
            targets.add("F")
    return sorted(targets)


def parse_tags(raw):
    """Comma-separated, or [bracketed]. Keeps order, drops duplicates."""
    if not raw:
        return []
    bracketed = re.findall(r"\[([^\]]+)\]", raw)
    parts = bracketed if bracketed else re.split(r"[,•|]", raw)
    out, seen = [], set()
    for part in parts:
        tag = part.strip().strip("[]()").strip()
        if not tag or tag.lower() in BLANK or len(tag) > 60:
            continue
        if tag.lower() not in seen:
            seen.add(tag.lower())
            out.append(tag)
    return out


def split_people(raw):
    """'A, B' / 'A & B' / 'A x B' -> ['A', 'B'], placeholders dropped."""
    if not raw:
        return []
    out = []
    for part in re.split(r"[,/;&]| \bx\b | \+ ", raw):
        name = re.sub(r"^(?:/?u/)", "", part.strip().lstrip("@").strip(), flags=re.I).strip()
        if name and name.lower() not in BLANK and name not in out:
            out.append(name)
    return out


def person(name):
    return {"name": name, "url": f"https://www.reddit.com/user/{name}"}


def credit(names):
    """The card renders one credit line, so several names are joined."""
    if not names:
        return None
    if len(names) == 1:
        return person(names[0])
    return {"name": ", ".join(names), "url": person(names[0])["url"]}


def normalize_url(url):
    if not url:
        return None
    url = url.strip()
    if url.startswith("//"):
        url = "https:" + url
    elif not re.match(r"^https?://", url, re.I):
        if re.match(r"^(www\.|reddit\.com|soundgasm|audiochan|hotaudio|patreon|clips4sale|subscribestar)",
                    url, re.I):
            url = "https://" + url
        else:
            return None
    return re.sub(r"[?&]utm_[^&]*", "", url).rstrip("?&")


# ─── Duration ────────────────────────────────────────────────────────

def parse_duration(raw):
    """
    Normalize to 'M:SS', or 'H:MM:SS' for anything an hour or longer.

    Google Sheets treats an 'MM:SS' cell as a time, so it exports 26:00
    (26 minutes) as '26:00:00' — 26 hours. The giveaway is the seconds
    component: a mangled MM:SS always lands on :00, while a genuine
    H:MM:SS keeps real seconds. The one case this can't tell apart is a
    real duration of exactly H:MM:00, which is reported at the end.
    """
    raw = (raw or "").strip()
    if not raw or raw.lower() in BLANK:
        return "", False
    if "-" in raw and re.search(r"\d\s*[mh]", raw):
        return raw, False                              # a range, e.g. "15m-40m"

    def fmt(total_seconds):
        hours, rest = divmod(int(total_seconds), 3600)
        minutes, seconds = divmod(rest, 60)
        return f"{hours}:{minutes:02d}:{seconds:02d}" if hours else f"{minutes}:{seconds:02d}"

    # "22m 52s" / "28m" / "45s" / "1h12m"
    m = re.fullmatch(r"(?:(\d+)\s*h)?\s*(?:(\d+)\s*m)?\s*(?:(\d+)\s*s)?", raw, re.I)
    if m and any(m.groups()):
        h, mi, se = (int(g) if g else 0 for g in m.groups())
        return fmt(h * 3600 + mi * 60 + se), False

    # "26:00:00" (mangled) or "1:07:12" (genuine)
    m = re.fullmatch(r"(\d+):(\d{2}):(\d{2})", raw)
    if m:
        h, mi, se = (int(g) for g in m.groups())
        if se == 0:
            # Read as mangled MM:SS. Only worth flagging when the H:MM:00
            # reading is also believable — "13:52:00" is not 13 hours of
            # audio, but "2:05:00" really could be two hours five minutes.
            return fmt(h * 60 + mi), (1 <= h <= 3 and mi > 0)
        return fmt(h * 3600 + mi * 60 + se), False

    # "13:52" — already plain text, exactly what we want
    m = re.fullmatch(r"(\d+):(\d{2})", raw)
    if m:
        mi, se = int(m.group(1)), int(m.group(2))
        return fmt(mi * 60 + se), False

    return raw, False


# ─── Dates ───────────────────────────────────────────────────────────

def detect_date_order(values):
    """
    Decide whether slash dates are D/M/Y or M/D/Y by looking at the whole
    column: any day above 12 settles it. Sheets exports in the document's
    locale, so this avoids hard-coding an assumption that breaks silently.
    """
    first_over_12 = second_over_12 = 0
    for value in values:
        m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", (value or "").strip())
        if not m:
            continue
        a, b = int(m.group(1)), int(m.group(2))
        if a > 12:
            first_over_12 += 1
        if b > 12:
            second_over_12 += 1
    if first_over_12 and not second_over_12:
        return "DMY"
    if second_over_12 and not first_over_12:
        return "MDY"
    return "MDY" if second_over_12 <= first_over_12 else "DMY"


def parse_date(raw, order="MDY"):
    raw = (raw or "").strip()
    if not raw:
        return None

    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
        return raw
    m = re.match(r"(\d{4}-\d{2}-\d{2})[T ]", raw)
    if m:
        return m.group(1)

    # "21 jul., 2025" / "21 Jul 2025"
    m = re.match(r"(\d{1,2})\s+(\w+)\.?,?\s*(\d{4})$", raw)
    if m:
        abbr = m.group(2).lower()[:3]
        month = PT_MONTHS.get(abbr) or EN_MONTHS.get(abbr)
        if month:
            return f"{m.group(3)}-{month}-{int(m.group(1)):02d}"

    # "Jul 21, 2025" / "July 21 2025"
    m = re.match(r"([A-Za-zçÇ]+)\.?\s+(\d{1,2}),?\s*(\d{4})$", raw)
    if m:
        abbr = m.group(1).lower()[:3]
        month = EN_MONTHS.get(abbr) or PT_MONTHS.get(abbr)
        if month:
            return f"{m.group(3)}-{month}-{int(m.group(2)):02d}"

    # 21/07/2025, 21-07-2025, 21.07.2025
    m = re.fullmatch(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})", raw)
    if m:
        a, b, year = int(m.group(1)), int(m.group(2)), m.group(3)
        if a > 12:
            day, month = a, b
        elif b > 12:
            day, month = b, a
        else:
            day, month = (a, b) if order == "DMY" else (b, a)
        return f"{year}-{month:02d}-{day:02d}"

    return None


# ─── Series ──────────────────────────────────────────────────────────

# ─── Tag index (powers the site's search) ────────────────────────────

# Tags that mean the same thing but are never written the same way, so
# co-occurrence can't find them: a synonym is a *substitute*, meaning the
# two rarely appear on the same audio. Each row is one concept; searching
# any word in it finds all the others. Add rows as the sheet's vocabulary
# grows — everything else in the index is derived automatically.
#
# Keep these strictly to different wordings of ONE concept. Things that
# merely go together — garden/birds/nature, praise/good girl — belong to
# the co-occurrence map below, which ranks them as related rather than
# exact. Putting them here promotes every bird audio into the exact
# results for "garden", which is how this list was wrong the first time.
TAG_SYNONYMS = [
    ["fdom", "femdom", "female dom", "female dominant", "domme"],
    ["mdom", "male dom", "male dominant"],
    ["fsub", "female sub", "female submissive"],
    ["msub", "male sub", "male submissive"],
    ["joi", "jerk off instruction", "guided masturbation", "guided wank"],
    ["gfe", "girlfriend experience", "girlfriend"],
    ["bfe", "boyfriend experience", "boyfriend"],
    ["cnc", "consensual non consent"],
    ["asmr", "tingles", "trigger sounds"],
    ["d2l", "down to listener"],
    ["sfw", "safe for work", "clean", "no sex"],
    ["nsfw", "not safe for work", "explicit"],
    ["l-bombs", "l bombs", "i love you", "love confession"],
    ["blowjob", "bj", "oral", "sucking"],
    ["cunnilingus", "going down", "eating out"],
    ["ramblefap", "ramble", "rambling", "stream of consciousness"],
    ["improv", "improvised", "unscripted", "off the cuff"],
    ["script fill", "scripted", "fill"],
    ["aftercare", "cuddles", "cuddling"],
    ["pegging", "strap on", "strapon"],
    ["hypno", "hypnosis", "hypnotic", "trance"],
    ["degradation", "degrading"],
    ["praise", "praise kink"],
    ["wet sounds", "wet noises", "wet pussy sounds", "wet pussy noises"],
    ["listener orgasm", "you cum"],
    ["mutual orgasm", "cum together", "cum with me"],
    ["overstim", "overstimulation", "oversensitive"],
    ["edging", "orgasm denial"],
    ["whisper", "whispers", "whispering", "whispery", "soft spoken"],
    ["moaning", "moans", "moan"],
    ["gentle", "soft", "tender"],
    ["sleep aid", "sleep", "bedtime", "falling asleep"],
]

# How a tag is written on a card vs. how it is matched: casing and spacing
# vary across the sheet ("Script Fill", "script fill", "SCRIPT FILL").
def norm_tag(tag):
    return re.sub(r"\s+", " ", (tag or "").strip().lower())


def build_tag_index(entries, min_freq=4, min_shared=3, max_related=8):
    """
    A vocabulary plus a 'related concepts' map for the search box.

    Relatedness is measured by co-occurrence: two tags are related when the
    audios carrying one tend to carry the other. Jaccard (shared / either)
    rather than raw counts, so a common tag like 'moaning' isn't related to
    everything. Rare tags are skipped — with ~220 audios, a tag used twice
    can't say anything reliable, and the long tail here is mostly one-off
    descriptive phrases rather than real concepts.
    """
    freq = Counter()
    per_entry = []
    for entry in entries:
        tags = {norm_tag(t) for t in entry["tags"] if norm_tag(t)}
        per_entry.append(tags)
        freq.update(tags)

    concepts = {t for t, c in freq.items() if c >= min_freq}

    shared = defaultdict(Counter)
    for tags in per_entry:
        present = sorted(tags & concepts)
        for i, a in enumerate(present):
            for b in present[i + 1:]:
                shared[a][b] += 1
                shared[b][a] += 1

    related = {}
    for tag in concepts:
        scored = []
        for other, together in shared[tag].items():
            if together < min_shared:
                continue
            union = freq[tag] + freq[other] - together
            if union:
                scored.append((together / union, other))
        scored.sort(reverse=True)
        if scored:
            related[tag] = [name for _, name in scored[:max_related]]

    # Keep only synonym rows that touch this catalogue, so the shipped index
    # reflects the actual vocabulary rather than a generic word list.
    known = set(freq)
    synonyms = []
    for row in TAG_SYNONYMS:
        cleaned = [norm_tag(w) for w in row]
        if any(w in known for w in cleaned):
            synonyms.append(cleaned)

    return {
        "vocab": {t: freq[t] for t in sorted(freq) if freq[t] >= 2},
        "related": {t: related[t] for t in sorted(related)},
        "synonyms": synonyms,
    }


def detect_series(titles):
    """Label a series only where a numbered prefix is shared by 2+ entries."""
    groups = defaultdict(list)
    for title in titles:
        m = SERIES_RE.match(title)
        if m:
            name = m.group(1).strip(" -–:#")
            groups[re.sub(r"s$", "", name.lower())].append((title, name))
    series = {}
    for members in groups.values():
        if len(members) < 2:
            continue
        label = Counter(name for _, name in members).most_common(1)[0][0]
        for title, _ in members:
            series[title] = label
    return series


# ─── Build ───────────────────────────────────────────────────────────

def build(rows):
    getters = [row_getter(r) for r in rows]
    date_order = detect_date_order([g("Date") for g in getters])
    titles = [g("Title") for g in getters]
    series_of = detect_series([t for t in titles if t])

    entries, used_ids, notes = [], set(), Counter()

    for get in getters:
        raw_title = get("Title")
        if not raw_title:
            notes["row skipped: no title"] += 1
            continue

        exclusive = bool(EXCLUSIVE_RE.match(raw_title))
        title = EXCLUSIVE_RE.sub("", raw_title).strip() if exclusive else raw_title

        entry_id = slugify(title)
        if entry_id in used_ids:
            n = 2
            while f"{entry_id}-{n}" in used_ids:
                n += 1
            entry_id = f"{entry_id}-{n}"
            notes["duplicate title, id given a suffix"] += 1
        used_ids.add(entry_id)

        links = {}
        for column, key in LINK_COLUMNS:
            url = normalize_url(get(column))
            if url:
                links[key] = url
        if not links:
            notes["no links at all"] += 1

        writers, editors = split_people(get("Writer")), split_people(get("Editor"))
        collab = split_people(get("Collab Partners"))

        date = parse_date(get("Date"), date_order)
        if get("Date") and not date:
            notes["date not understood"] += 1
        elif not date:
            notes["no date"] += 1

        duration, ambiguous = parse_duration(get("Duration"))
        if ambiguous:
            notes["duration could be MM:SS or H:MM:00"] += 1

        description = get("Description")
        if not description:
            notes["no description"] += 1

        entry_type = normalize_type(get("Type"))
        if entry_type == "Uncategorized":
            notes["no type"] += 1

        all_tags = parse_tags(get("Tags"))
        audience = audience_from_tags(all_tags)
        if not audience:
            notes["no audience tag (F4M / F4A / …)"] += 1

        # The audience now has its own filter in the sidebar, so F4M / F4A /
        # FF4M and friends are dropped from the chips on the card. They also
        # sat on nearly every audio, which made them turn up as "related
        # concepts" for almost any search. The site still matches them: a
        # query for f4m is answered from the audience field above.
        tags = [t for t in all_tags if not AUDIENCE_RE.fullmatch(t.strip())]

        entries.append({
            "id": entry_id,
            "title": title,
            "date": date,
            "duration": duration or None,
            "series": series_of.get(raw_title),
            "type": entry_type,
            "audience": audience,
            "tags": tags,
            "description": description or None,
            "image": None,
            "exclusive": exclusive,
            "writer": credit(writers),
            "editor": credit(editors),
            "collab": [person(n) for n in collab] or None,
            "links": links,
        })

    entries.sort(key=lambda e: (e["date"] or "0000-00-00"), reverse=True)
    return entries, notes, date_order


def main():
    ap = argparse.ArgumentParser(description="Build audios.json from the masterlist spreadsheet.")
    ap.add_argument("--url", default=SHEET_CSV_URL, help="published Google Sheets CSV URL")
    ap.add_argument("--xlsx", help="read a local .xlsx instead of fetching")
    ap.add_argument("--sheet", help="sheet name when using --xlsx")
    ap.add_argument("-o", "--out", default=OUTPUT_FILE)
    ap.add_argument("--indent", type=int, default=1, help="0 for a minified file")
    args = ap.parse_args()

    print("Building audios.json...")
    print("\n1. Reading the spreadsheet...")
    if args.xlsx:
        rows = read_xlsx(args.xlsx, args.sheet)
        source = os.path.basename(args.xlsx)
    else:
        if args.url.startswith("YOUR_"):
            raise SystemExit(
                "No sheet URL set. Either export SHEET_CSV_URL, pass --url, "
                "or build from a local file with --xlsx <file>."
            )
        rows = fetch_csv(args.url)
        source = "published sheet"
    print(f"   {len(rows)} rows from {source}")

    print("\n2. Processing entries...")
    entries, notes, date_order = build(rows)
    print(f"   {len(entries)} entries | slash dates read as {date_order}")

    types = sorted({e["type"] for e in entries})
    tag_index = build_tag_index(entries)
    payload = {
        "lastUpdated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "totalEntries": len(entries),
        "types": types,
        "tagIndex": tag_index,
        "audios": entries,
    }

    print("\n3. Writing...")
    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False,
                  indent=args.indent or None,
                  separators=(",", ":") if not args.indent else None)

    size = os.path.getsize(args.out)
    print(f"\nDone. Wrote {len(entries)} entries to {args.out} ({size/1024:.0f} KB)")
    if entries:
        print(f"   dates      {entries[-1]['date']} .. {entries[0]['date']}")
    print(f"   types      {', '.join(f'{t} {c}' for t, c in Counter(e['type'] for e in entries).most_common())}")
    print(f"   links      {', '.join(f'{k} {c}' for k, c in Counter(k for e in entries for k in e['links']).most_common())}")
    print(f"   series     {sum(1 for e in entries if e['series'])} entries in "
          f"{len({e['series'] for e in entries if e['series']})} series")
    print(f"   tag index  {len(tag_index['vocab'])} tags in the vocabulary, "
          f"{len(tag_index['related'])} with related concepts, "
          f"{len(tag_index['synonyms'])} synonym groups in use")
    aud = Counter()
    for e in entries:
        aud["+".join(e["audience"]) or "(none)"] += 1
    print(f"   audience   {', '.join(f'{k} {v}' for k, v in aud.most_common())}")
    print(f"   exclusive  {sum(1 for e in entries if e['exclusive'])}")
    print(f"   credits    writer {sum(1 for e in entries if e['writer'])}, "
          f"editor {sum(1 for e in entries if e['editor'])}, "
          f"collab {sum(1 for e in entries if e['collab'])}")
    if notes:
        print("   worth a look:")
        for label, count in notes.most_common():
            print(f"      {count:4d}  {label}")


if __name__ == "__main__":
    sys.exit(main())
