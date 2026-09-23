#!/usr/bin/env python3
"""
Fracton streamer finder (v3)
----------------------------
Finds small English-language Twitch streamers in the UK, US and Canada who are
good candidates for a Fracton POC, scores them, and writes both an Excel pack
for the team and an interactive HTML board for you.

Commands:
  python streamer_finder_v3.py login       one time: log in with Twitch
  python streamer_finder_v3.py collect     one sweep (live streams + channel details)
  python streamer_finder_v3.py loop 60     sweep every 60 minutes (Ctrl+C to stop)
  python streamer_finder_v3.py report      CSV exports
  python streamer_finder_v3.py excel       fracton_streamers.xlsx (for the team)
  python streamer_finder_v3.py board       fracton_board.html (for you, open in browser)
  python streamer_finder_v3.py all         report + excel + board

Needs: pip install requests openpyxl
Uses the same streamers.db, credentials.json and token.json as before.
"""
import csv
import html
import json
import os
import re
import secrets
import sqlite3
import statistics
import sys
import time
import webbrowser
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlencode, urlparse

import requests

# ======================= SETTINGS (edit these) =======================
FOLLOWERS_MIN = 100
FOLLOWERS_MAX = 10_000
VIEWERS_MIN = 5            # average live viewers, lower bound
VIEWERS_MAX = 150          # upper bound the server can carry -> confirm with the tech team
MIN_DAYS_SEEN = 2          # must be seen live on at least this many different days
TARGET_REGIONS = {"UK", "US", "CA"}
ENRICH_MIN_SIGHTINGS = 2   # only look up channels seen live in this many sweeps
MAX_STREAM_HOURS = 14      # longer than this on average = 24/7 rerun channel, excluded
MAX_PAGES = 400            # safety cap per sweep (100 streams per page)
REFRESH_CHANNEL_DAYS = 3   # re-fetch follower counts older than this
# =====================================================================

BASE = os.path.dirname(os.path.abspath(__file__))
CRED_FILE = os.path.join(BASE, "credentials.json")
TOKEN_FILE = os.path.join(BASE, "token.json")
DB_FILE = os.path.join(BASE, "streamers.db")
XLSX_FILE = os.path.join(BASE, "fracton_streamers.xlsx")
HTML_FILE = os.path.join(BASE, "fracton_board.html")
REDIRECT = "http://localhost:3000"
API = "https://api.twitch.tv/helix/"


# ----------------------------- auth ---------------------------------
def load_creds():
    env_id = os.environ.get("TWITCH_CLIENT_ID")
    env_secret = os.environ.get("TWITCH_CLIENT_SECRET")
    if env_id and env_secret:
        return env_id, env_secret
    if not os.path.exists(CRED_FILE):
        sys.exit("No credentials. Set TWITCH_CLIENT_ID and TWITCH_CLIENT_SECRET, or create "
                 'credentials.json with: {"client_id": "...", "client_secret": "..."}')
    with open(CRED_FILE) as f:
        c = json.load(f)
    return c["client_id"], c["client_secret"]


def save_token(t):
    t["expires_at"] = time.time() + t.get("expires_in", 3600) - 60
    with open(TOKEN_FILE, "w") as f:
        json.dump(t, f)


def login():
    cid, secret = load_creds()
    state = secrets.token_urlsafe(16)
    url = "https://id.twitch.tv/oauth2/authorize?" + urlencode({
        "client_id": cid, "redirect_uri": REDIRECT,
        "response_type": "code", "scope": "", "state": state})
    result = {}

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            q = parse_qs(urlparse(self.path).query)
            result.update({k: v[0] for k, v in q.items()})
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(b"<h2>Done - close this tab and go back to the terminal.</h2>")

        def log_message(self, *args):
            pass

    server = HTTPServer(("localhost", 3000), Handler)
    print("Opening your browser to log in with Twitch...")
    print("If nothing opens, paste this link into your browser:\n" + url)
    webbrowser.open(url)
    while "code" not in result and "error" not in result:
        server.handle_request()
    server.server_close()
    if "code" not in result or result.get("state") != state:
        sys.exit(f"Login failed: {result.get('error_description', result)}")
    r = requests.post("https://id.twitch.tv/oauth2/token", data={
        "client_id": cid, "client_secret": secret, "code": result["code"],
        "grant_type": "authorization_code", "redirect_uri": REDIRECT}, timeout=60)
    r.raise_for_status()
    save_token(r.json())
    print("Logged in.")


_APP_TOKEN = {"value": None, "expires_at": 0}


def app_token():
    """Client-credentials token: no browser login needed. Enough for stream lists,
    profiles and follower totals, which is everything this script reads."""
    if _APP_TOKEN["value"] and time.time() < _APP_TOKEN["expires_at"]:
        return _APP_TOKEN["value"]
    cid, secret = load_creds()
    r = requests.post("https://id.twitch.tv/oauth2/token", data={
        "client_id": cid, "client_secret": secret,
        "grant_type": "client_credentials"}, timeout=60)
    r.raise_for_status()
    t = r.json()
    _APP_TOKEN["value"] = t["access_token"]
    _APP_TOKEN["expires_at"] = time.time() + t.get("expires_in", 3600) - 120
    return _APP_TOKEN["value"]


def get_token(force_refresh=False):
    if not os.path.exists(TOKEN_FILE):
        return app_token()
    with open(TOKEN_FILE) as f:
        t = json.load(f)
    if force_refresh or time.time() > t["expires_at"]:
        cid, secret = load_creds()
        r = requests.post("https://id.twitch.tv/oauth2/token", data={
            "client_id": cid, "client_secret": secret,
            "grant_type": "refresh_token", "refresh_token": t["refresh_token"]}, timeout=60)
        if r.status_code != 200:
            sys.exit("Session expired. Run: python streamer_finder_v3.py login")
        t = r.json()
        save_token(t)
    return t["access_token"]


class Twitch:
    def __init__(self):
        self.cid, _ = load_creds()
        self.s = requests.Session()
        self._set_auth()

    def _set_auth(self, force=False):
        self.s.headers.update({"Client-Id": self.cid,
                               "Authorization": f"Bearer {get_token(force)}"})

    def get(self, path, params):
        for attempt in range(8):
            try:
                r = self.s.get(API + path, params=params, timeout=60)
            except requests.exceptions.RequestException as e:
                wait = min(60, 5 * (attempt + 1))
                print(f"    network hiccup ({type(e).__name__}); retrying in {wait}s")
                time.sleep(wait)
                continue
            if r.status_code == 401:
                self._set_auth(force=True)
                continue
            if r.status_code == 429 or r.status_code >= 500:
                reset = int(r.headers.get("Ratelimit-Reset", time.time() + 5))
                time.sleep(max(1, reset - time.time()) + 1)
                continue
            r.raise_for_status()
            if int(r.headers.get("Ratelimit-Remaining", 100)) < 10:
                reset = int(r.headers.get("Ratelimit-Reset", time.time() + 5))
                time.sleep(max(1, reset - time.time()) + 1)
            return r.json()
        raise RuntimeError(f"Twitch API kept failing on {path}")


# ----------------------------- storage ------------------------------
def open_db():
    db = sqlite3.connect(DB_FILE)
    db.execute("""CREATE TABLE IF NOT EXISTS sightings (
        user_id TEXT, login TEXT, name TEXT, seen_at TEXT, stream_id TEXT,
        started_at TEXT, viewers INTEGER, game TEXT, title TEXT, tags TEXT,
        PRIMARY KEY (user_id, seen_at))""")
    db.execute("""CREATE TABLE IF NOT EXISTS channels (
        user_id TEXT PRIMARY KEY, login TEXT, name TEXT, broadcaster_type TEXT,
        description TEXT, created_at TEXT, followers INTEGER, updated_at REAL)""")
    return db


# ----------------------------- collect ------------------------------
def collect(api, db):
    seen_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M")
    cursor, pages, recorded, low_pages = None, 0, 0, 0
    print(f"[{seen_at} UTC] Scanning live English streams...")
    while pages < MAX_PAGES:
        params = {"language": "en", "first": 100, "type": "live"}
        if cursor:
            params["after"] = cursor
        data = api.get("streams", params)
        streams = data.get("data", [])
        if not streams:
            break
        pages += 1
        for s in streams:
            vc = s["viewer_count"]
            if VIEWERS_MIN <= vc <= VIEWERS_MAX * 2:
                db.execute("INSERT OR IGNORE INTO sightings VALUES (?,?,?,?,?,?,?,?,?,?)", (
                    s["user_id"], s["user_login"], s["user_name"], seen_at, s["id"],
                    s["started_at"], vc, s.get("game_name", ""), s.get("title", ""),
                    ",".join(s.get("tags") or [])))
                recorded += 1
        if max(s["viewer_count"] for s in streams) < VIEWERS_MIN:
            low_pages += 1
            if low_pages >= 2:
                break
        cursor = data.get("pagination", {}).get("cursor")
        if not cursor:
            break
    db.commit()
    print(f"  {pages} pages scanned, {recorded} streams in your viewer range recorded.")


def enrich(api, db):
    cutoff = time.time() - REFRESH_CHANNEL_DAYS * 86400
    ids = [r[0] for r in db.execute("""
        SELECT s.user_id FROM sightings s
        LEFT JOIN channels c ON c.user_id = s.user_id
        WHERE c.user_id IS NULL OR c.updated_at < ?
        GROUP BY s.user_id HAVING COUNT(*) >= ?""", (cutoff, ENRICH_MIN_SIGHTINGS))]
    if not ids:
        print("  No new channels to look up.")
        return
    print(f"  Looking up {len(ids)} channels (profile + follower count)...")
    for i in range(0, len(ids), 100):
        batch = ids[i:i + 100]
        users = api.get("users", [("id", x) for x in batch]).get("data", [])
        for u in users:
            f = api.get("channels/followers", {"broadcaster_id": u["id"], "first": 1})
            db.execute("INSERT OR REPLACE INTO channels VALUES (?,?,?,?,?,?,?,?)", (
                u["id"], u["login"], u["display_name"], u.get("broadcaster_type") or "none",
                u.get("description", ""), u.get("created_at", "")[:10],
                f.get("total", 0), time.time()))
        db.commit()
        print(f"    {min(i + 100, len(ids))}/{len(ids)} done")


# ------------------------- signal detection -------------------------
TARGET_WORDS = {
    "UK": ["uk", "u.k.", "united kingdom", "england", "scotland", "scottish", "wales",
           "welsh", "northern ireland", "british", "brit", "london", "manchester",
           "birmingham", "liverpool", "leeds", "glasgow", "edinburgh", "cardiff",
           "bristol", "newcastle", "sheffield", "nottingham", "yorkshire", "essex",
           "kent", "devon", "cornwall", "brummie", "geordie", "scouse"],
    "US": ["usa", "u.s.a.", "united states", "american", "texas", "california",
           "florida", "new york", "nyc", "chicago", "los angeles", "atlanta", "seattle",
           "ohio", "michigan", "arizona", "colorado", "tennessee", "north carolina",
           "south carolina", "boston", "philadelphia", "houston", "dallas", "las vegas",
           "portland", "denver", "minnesota", "missouri", "alabama", "oregon", "utah",
           "virginia", "wisconsin", "indiana", "kentucky", "louisiana", "oklahoma"],
    "CA": ["canada", "canadian", "toronto", "vancouver", "montreal", "calgary", "ottawa",
           "edmonton", "winnipeg", "ontario", "alberta", "quebec", "nova scotia",
           "british columbia", "saskatchewan", "manitoba", "halifax"],
}
TIMEZONE_WORDS = {
    "UK": ["gmt", "bst"],
    "US": ["est", "edt", "cst", "cdt", "pst", "pdt", "mst", "mdt", "eastern time",
           "pacific time", "central time"],
    "CA": ["atlantic time", "nst"],
}
# Anyone signalling these is not in the target market, whatever their hours look like.
EXCLUDE_WORDS = [
    "australia", "australian", "aussie", "sydney", "melbourne", "brisbane", "perth",
    "adelaide", "aest", "aedt", "new zealand", "kiwi", "auckland", "nzst",
    "philippines", "filipino", "filipina", "pinoy", "manila", "cebu", "pht",
    "india", "indian", "mumbai", "delhi", "bangalore", "ist", "pakistan", "karachi",
    "malaysia", "singapore", "indonesia", "jakarta", "thailand", "vietnam",
    "south africa", "nigeria", "nigerian", "lagos", "kenya", "ghana", "egypt",
    "germany", "german", "deutschland", "berlin", "cet", "cest", "netherlands",
    "dutch", "amsterdam", "belgium", "sweden", "swedish", "norway", "norwegian",
    "finland", "denmark", "danish", "poland", "polish", "czech", "portugal",
    "portuguese", "spain", "spanish", "espa\u00f1ol", "italy", "italian", "france",
    "french", "fran\u00e7ais", "greece", "romania", "hungary", "austria", "switzerland",
    "ireland", "irish", "dublin", "brazil", "brasil", "brazilian", "mexico", "mexican",
    "argentina", "chile", "colombia", "peru", "japan", "japanese", "tokyo", "korea",
    "korean", "seoul", "china", "chinese", "taiwan", "hong kong", "turkey", "israel",
    "uae", "dubai", "saudi", "russia", "russian", "ukraine", "ukrainian",
]
# 24/7 reruns, drop farms and watch parties: high viewers, no real community.
JUNK_PATTERN = re.compile(
    r"24\s*[/\-]?\s*7|24hr|rerun|re-run|replay stream|watch\s?part(y|ies)|afk\s*(stream|farm)?|"
    r"drop\s?farm|badge\s?(farm|collect)|always\s?on|idle stream|bot stream|"
    r"vod channel|second channel|drops enabled|=\s*badge|watch\s*\d+\s*h(r|our)?\s*=", re.I)
JUNK_NAME_PATTERN = re.compile(r"badge|vods?\b|247|24_7|dropfarm|afkzone", re.I)
GIVEAWAY_PATTERN = re.compile(
    r"giveaway|give\s?away|raffle|!giveaway|prize|wheel spin|sub goal reward", re.I)
CODE_PATTERN = re.compile(
    r"(?:use|promo|discount|creator|my)\s+code[:\s]+([A-Za-z0-9_]{2,18})|"
    r"code[:\s]+([A-Z][A-Z0-9_]{2,17})\b", re.I)
SPONSOR_PATTERN = re.compile(
    r"sponsored|#ad\b|partnered with|brand deal|ambassador|affiliate link|"
    r"!code\b|!sponsor\b|!merch\b|!partner\b", re.I)
BRANDS = [
    "G FUEL", "GFuel", "Gamer Supps", "Sneak Energy", "Ghost Energy", "Rogue Energy",
    "X-Gamer", "Red Bull", "Monster Energy", "Bang Energy", "Celsius", "Prime Hydration",
    "Elgato", "HyperX", "Razer", "Logitech", "SteelSeries", "Corsair", "Turtle Beach",
    "Astro Gaming", "SCUF Gaming", "KontrolFreek", "Glorious Gaming", "NZXT", "MSI",
    "Alienware", "Lenovo Legion", "Secretlab", "Herman Miller", "Shure", "Blue Yeti",
    "Audio-Technica", "Displate", "NordVPN", "ExpressVPN", "Opera GX", "Streamlabs",
    "StreamElements", "Manscaped", "HelloFresh", "DoorDash", "Chipotle", "Crunchyroll",
    "Prime Gaming", "Discord Nitro", "Raid Shadow Legends", "Gymshark", "Lootcrate",
    "Nerd Focus", "Advanced.gg", "JuJu Supps", "Ridge Wallet", "Displate",
]
BRAND_PATTERN = re.compile(r"\b(" + "|".join(re.escape(b) for b in BRANDS) + r")\b", re.I)
BRAND_CANON = {b.lower(): b for b in BRANDS}


def word_pattern(words):
    return re.compile(r"\b(" + "|".join(re.escape(w) for w in words) + r")\b", re.I)


TARGET_PATTERNS = {k: word_pattern(v) for k, v in TARGET_WORDS.items()}
TZ_PATTERNS = {k: word_pattern(v) for k, v in TIMEZONE_WORDS.items()}
EXCLUDE_PATTERN = word_pattern(EXCLUDE_WORDS)
FLAGS = {"UK": ["\U0001F1EC\U0001F1E7", "\U0001F3F4"], "US": ["\U0001F1FA\U0001F1F8"],
         "CA": ["\U0001F1E8\U0001F1E6"]}


def guess_region(text, start_hours):
    """Returns (region, confidence, evidence). Profile words beat timezone beat hours."""
    low = text.lower()
    hits = {}
    for region, pat in TARGET_PATTERNS.items():
        found = {m.lower() for m in pat.findall(text)} | {f for f in FLAGS[region] if f in text}
        if region == "UK" and "british columbia" in low:
            found.discard("british")
        if found:
            hits[region] = found
    if hits:
        best = max(hits, key=lambda r: len(hits[r]))
        return best, "high", ", ".join(sorted(hits[best]))

    outside = {m.lower() for m in EXCLUDE_PATTERN.findall(text)}
    if outside:
        return "other", "high", "outside target market: " + ", ".join(sorted(outside))

    for region, pat in TZ_PATTERNS.items():
        found = {m.lower() for m in pat.findall(text)}
        if found:
            return region, "medium", "timezone in profile: " + ", ".join(sorted(found))

    if start_hours:
        n = len(start_hours)
        uk = sum(1 for h in start_hours if 16 <= h <= 22) / n
        na = sum(1 for h in start_hours if h >= 23 or h <= 6) / n
        apac = sum(1 for h in start_hours if 7 <= h <= 13) / n
        if apac >= 0.7:
            return "other", "low", "goes live in Asia-Pacific hours"
        if uk >= 0.7:
            return "UK?", "low", "goes live in UK evening (could be Europe)"
        if na >= 0.7:
            return "US/CA?", "low", "goes live in North American evening"
    return "unknown", "none", ""


def find_brands(text):
    found = {BRAND_CANON.get(m.lower(), m) for m in BRAND_PATTERN.findall(text)}
    found = {"G FUEL" if b.lower() in ("gfuel", "g fuel") else b for b in found}
    return sorted(found)


def find_codes(text):
    codes = set()
    for a, b in CODE_PATTERN.findall(text):
        c = (a or b).strip()
        if c and c.lower() not in ("of", "the", "for", "your", "here"):
            codes.add(c.upper())
    return sorted(codes)


# ----------------------------- rows ---------------------------------
def parse_seen(s):
    return datetime.strptime(s, "%Y-%m-%dT%H:%M").replace(tzinfo=timezone.utc)


def parse_started(s):
    return datetime.strptime(s[:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)


def build_rows(db):
    channels = {r[0]: r for r in db.execute("SELECT * FROM channels")}
    sweeps = [r[0] for r in db.execute("SELECT DISTINCT seen_at FROM sightings")]
    sweep_days = len({s[:10] for s in sweeps}) or 1
    by_user = defaultdict(list)
    for r in db.execute("SELECT user_id, started_at, viewers, game, title, tags, seen_at, "
                        "stream_id FROM sightings ORDER BY seen_at"):
        by_user[r[0]].append(r)

    rows = []
    for uid, sights in by_user.items():
        ch = channels.get(uid)
        if not ch:
            continue
        _, login, name, btype, bio, created, followers, _ = ch
        viewers = [s[2] for s in sights]
        starts = {s[1] for s in sights}
        start_hours = [int(s[11:13]) for s in starts]
        days = {s[:10] for s in starts}
        games = Counter(s[3] for s in sights if s[3])
        titles = " | ".join(dict.fromkeys(s[4] for s in sights))
        tags = ", ".join(dict.fromkeys(t for s in sights for t in s[5].split(",") if t))

        # how long each stream had been running when we saw it
        per_stream = defaultdict(list)
        for s in sights:
            per_stream[s[7]].append((parse_seen(s[6]) - parse_started(s[1])).total_seconds() / 3600)
        durations = [max(v) for v in per_stream.values()]
        median_hours = round(statistics.median(durations), 1) if durations else 0

        text = f"{bio} {titles} {tags}"
        region, conf, evidence = guess_region(text, start_hours)
        brands = find_brands(text)
        codes = find_codes(text)
        sponsor = bool(SPONSOR_PATTERN.search(text)) or bool(codes)
        giveaway = bool(GIVEAWAY_PATTERN.search(text))
        avg_v = round(statistics.mean(viewers), 1)
        eng = round(100 * avg_v / followers, 2) if followers else 0

        junk = JUNK_PATTERN.search(f"{name} {titles} {bio}") or JUNK_NAME_PATTERN.search(f"{name} {login}")
        excluded = ""
        if junk:
            excluded = f"looks like a rerun/farm channel ({junk.group(0).strip()})"
        elif median_hours > MAX_STREAM_HOURS:
            excluded = f"streams run {median_hours}h on average (24/7 channel)"
        elif not (FOLLOWERS_MIN <= followers <= FOLLOWERS_MAX):
            excluded = f"{followers} followers is outside {FOLLOWERS_MIN}-{FOLLOWERS_MAX}"
        elif not (VIEWERS_MIN <= avg_v <= VIEWERS_MAX):
            excluded = f"{avg_v} average viewers is outside {VIEWERS_MIN}-{VIEWERS_MAX}"
        elif len(days) < MIN_DAYS_SEEN:
            excluded = f"only seen live on {len(days)} day(s)"
        elif region.rstrip("?") not in TARGET_REGIONS and region != "US/CA?":
            excluded = f"region: {region}"

        rows.append({
            "streamer": name,
            "login": login,
            "twitch_url": f"https://twitch.tv/{login}",
            "sullygnome_url": f"https://sullygnome.com/channel/{login}",
            "region": region,
            "location_confidence": conf,
            "location_evidence": evidence,
            "status": btype,
            "followers": followers,
            "avg_viewers": avg_v,
            "peak_viewers": max(viewers),
            "engagement_pct": eng,
            "days_seen_live": len(days),
            "streams_seen": len(starts),
            "avg_stream_hours": median_hours,
            "typical_start_utc": f"{int(statistics.median(start_hours)):02d}:00" if start_hours else "",
            "main_games": ", ".join(g for g, _ in games.most_common(3)),
            "does_giveaways": "yes" if giveaway else "",
            "has_sponsor_signal": "yes" if sponsor else "",
            "discount_codes": ", ".join(codes),
            "brands_seen": ", ".join(brands),
            "brand_count": len(brands),
            "account_created": created,
            "tags": tags,
            "bio": (bio or "").replace("\n", " ")[:400],
            "recent_title": sights[-1][4],
            "excluded_reason": excluded,
        })

    for r in rows:
        r["fit_score"], r["score_breakdown"] = fit_score(r, sweep_days)
        r["pitch_angle"] = pitch_angle(r)
    rows.sort(key=lambda r: -r["fit_score"])
    return rows, sweep_days


def fit_score(r, sweep_days):
    """0-100, with every component visible so the team can argue with it."""
    eng = min(r["engagement_pct"] / 10, 1) * 30          # 10%+ engagement = full marks
    consistency = min(r["days_seen_live"] / max(sweep_days, 1), 1) * 20
    if r["does_giveaways"] and (r["has_sponsor_signal"] or r["brand_count"]):
        signal = 20
    elif r["does_giveaways"]:
        signal = 15
    elif r["has_sponsor_signal"] or r["brand_count"]:
        signal = 10
    else:
        signal = 0
    loc = {"high": 15, "medium": 11, "low": 6, "none": 0}[r["location_confidence"]]
    v = r["avg_viewers"]
    size = 15 if 20 <= v <= 80 else (10 if VIEWERS_MIN <= v <= VIEWERS_MAX else 3)
    total = round(eng + consistency + signal + loc + size)
    breakdown = (f"engagement {eng:.0f}/30, consistency {consistency:.0f}/20, "
                 f"giveaway+brand {signal}/20, location {loc}/15, audience size {size}/15")
    return total, breakdown


def pitch_angle(r):
    bits = []
    if r["does_giveaways"] and r["discount_codes"]:
        bits.append("Already runs giveaways and carries a sponsor code, so they understand "
                    "brand deals. Lead with Fracton as proof of engagement they can show "
                    "sponsors: rewards land on the viewers who actually show up.")
    elif r["does_giveaways"]:
        bits.append("Already runs giveaways manually. Lead with what Fracton removes: no "
                    "random draws to defend in chat, no manual tracking, and loyal viewers "
                    "get rewarded automatically.")
    elif r["discount_codes"] or r["has_sponsor_signal"] or r["brand_count"]:
        bits.append("Already works with sponsors. Lead with Fracton as engagement data they "
                    "can take to brands, plus a reward loop that keeps regulars coming back.")
    elif r["engagement_pct"] >= 3:
        bits.append(f"Small following but {r['engagement_pct']}% of it watches live, so the "
                    "community is unusually tight. Lead with rewarding the regulars by name.")
    else:
        bits.append("No giveaway history spotted, so lead with the free POC: Fracton seeds "
                    "the prize pool, they keep full control, and it costs them nothing to try.")
    if r["status"] == "none":
        bits.append("Not yet an affiliate, so there is no sub revenue to protect - a reward "
                    "tool is pure upside for them.")
    if r["typical_start_utc"]:
        bits.append(f"Usually goes live around {r['typical_start_utc']} UTC; message them a "
                    "couple of hours before that.")
    return " ".join(bits)


def shortlist_of(rows):
    return [r for r in rows if not r["excluded_reason"]]


def brand_table(rows):
    per_brand = defaultdict(list)
    reach = Counter()
    for r in rows:
        for b in (r["brands_seen"].split(", ") if r["brands_seen"] else []):
            per_brand[b].append(r["streamer"])
            reach[b] += r["followers"]
    out = [{"brand": b, "streamers_count": len(v), "streamers": ", ".join(sorted(v)),
            "combined_followers": reach[b]} for b, v in per_brand.items()]
    out.sort(key=lambda x: -x["streamers_count"])
    return out


# ----------------------------- CSV ----------------------------------
def write_csv(path, rows):
    if not rows:
        print(f"  {os.path.basename(path)}: nothing to write yet.")
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"  {os.path.basename(path)}: {len(rows)} rows")


def report(db):
    rows, _ = build_rows(db)
    short = shortlist_of(rows)
    write_csv(os.path.join(BASE, "shortlist.csv"), short)
    write_csv(os.path.join(BASE, "all_candidates.csv"), rows)
    write_csv(os.path.join(BASE, "brands.csv"), brand_table(short))
    print(f"  {len(short)} shortlisted out of {len(rows)} channels seen.")


# ----------------------------- Excel --------------------------------
TOP_COLS = ["fit_score", "streamer", "region", "location_confidence", "followers",
            "avg_viewers", "engagement_pct", "status", "does_giveaways",
            "discount_codes", "brands_seen", "typical_start_utc", "main_games",
            "twitch_url", "pitch_angle"]


def excel(db):
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    rows, sweep_days = build_rows(db)
    short = shortlist_of(rows)
    if not short:
        sys.exit("No shortlisted streamers yet - run more collect sweeps first.")
    brands = brand_table(short)

    wb = Workbook()
    arial = Font(name="Arial")
    bold = Font(name="Arial", bold=True)
    white_bold = Font(name="Arial", bold=True, color="FFFFFF")
    purple = PatternFill("solid", fgColor="6441A5")
    light = PatternFill("solid", fgColor="F2EEFA")
    thin = Side(style="thin", color="D9D2E9")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    def sheet_from(name, data, cols, table_name):
        ws = wb.create_sheet(name)
        ws.append(cols)
        for r in data:
            ws.append([r[c] for c in cols])
        for c in ws[1]:
            c.font, c.fill = white_bold, purple
            c.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = arial
                c.alignment = Alignment(vertical="top", wrap_text=False)
        t = Table(displayName=table_name,
                  ref=f"A1:{get_column_letter(len(cols))}{len(data) + 1}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(t)
        ws.freeze_panes = "B2"
        for i, h in enumerate(cols, start=1):
            wide = {"pitch_angle": 70, "bio": 50, "recent_title": 50, "main_games": 28,
                    "location_evidence": 34, "brands_seen": 24, "streamers": 40,
                    "score_breakdown": 55, "excluded_reason": 34, "tags": 30}
            ws.column_dimensions[get_column_letter(i)].width = wide.get(h, 15)
        return ws

    # --- Top targets -------------------------------------------------
    ws_top = sheet_from("Top targets", short[:40], TOP_COLS, "TopTargets")
    ws_top.conditional_formatting.add(
        f"A2:A{len(short[:40]) + 1}",
        ColorScaleRule(start_type="num", start_value=0, start_color="F8CBAD",
                       end_type="num", end_value=100, end_color="A9D08E"))
    ws_top.column_dimensions["A"].width = 11
    ws_top.column_dimensions["B"].width = 22

    full_cols = [c for c in short[0].keys() if c not in ("login",)]
    sheet_from("Shortlist", short, full_cols, "ShortlistTable")
    sheet_from("All channels", rows, full_cols, "AllTable")
    if brands:
        sheet_from("Brands", brands, ["brand", "streamers_count", "combined_followers",
                                      "streamers"], "BrandsTable")

    # --- Dashboard ---------------------------------------------------
    ws = wb["Sheet"]
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False
    S, last = "Shortlist", len(short) + 1
    col = {h: get_column_letter(i + 1) for i, h in enumerate(full_cols)}

    def rng(field):
        return f"{S}!${col[field]}$2:${col[field]}${last}"

    ws["A1"] = "Fracton - Twitch POC candidates"
    ws["A1"].font = Font(name="Arial", bold=True, size=16)
    ws["A2"] = (f"UK / US / Canada, {sweep_days} days of live scanning, "
                f"generated {datetime.now(timezone.utc):%d %b %Y}. "
                "Every figure is a live formula over the Shortlist sheet.")
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="666666")

    cards = [
        ("Shortlisted", f"=COUNTA({rng('streamer')})", "streamers fit every filter"),
        ("Location confirmed", f'=COUNTIF({rng("location_confidence")},"high")',
         "country stated in their own profile"),
        ("Already do giveaways", f'=COUNTIF({rng("does_giveaways")},"yes")',
         "warmest possible opening line"),
        ("Sponsor signals", f'=COUNTIF({rng("has_sponsor_signal")},"yes")',
         "codes or brand deals visible"),
        ("Median audience", f"=ROUND(MEDIAN({rng('avg_viewers')}),0)",
         "average live viewers per stream"),
        ("Load if all tested at once", f"=ROUND(SUM({rng('avg_viewers')}),0)",
         "compare with the server limit"),
    ]
    for i, (label, formula, note) in enumerate(cards):
        c = 1 + (i % 3) * 4
        r = 4 + (i // 3) * 4
        ws.cell(r, c, label).font = Font(name="Arial", bold=True, size=10, color="6441A5")
        v = ws.cell(r + 1, c, formula)
        v.font = Font(name="Arial", bold=True, size=20)
        ws.cell(r + 2, c, note).font = Font(name="Arial", size=8, color="666666")
        for rr in range(r, r + 3):
            for cc in range(c, c + 3):
                ws.cell(rr, cc).fill = light
                ws.cell(rr, cc).border = box

    def block(row, title, labels, formulas, anchor, chart_title, x_title, y_title,
              horizontal=False):
        ws.cell(row, 1, title).font = bold
        for j, (lab, f) in enumerate(zip(labels, formulas), start=row + 1):
            ws.cell(j, 1, lab).font = arial
            ws.cell(j, 2, f).font = arial
        ch = BarChart()
        ch.type = "bar" if horizontal else "col"
        ch.title = chart_title
        ch.legend = None
        ch.height, ch.width = 6.5, 10
        ch.add_data(Reference(ws, min_col=2, min_row=row + 1, max_row=row + len(labels)))
        ch.set_categories(Reference(ws, min_col=1, min_row=row + 1, max_row=row + len(labels)))
        ch.x_axis.title, ch.y_axis.title = x_title, y_title
        ch.x_axis.delete, ch.y_axis.delete = False, False   # openpyxl hides them otherwise
        ch.y_axis.majorGridlines = None
        ws.add_chart(ch, anchor)

    reg = rng("region")
    block(13, "Where they are",
          ["UK (confirmed)", "US (confirmed)", "Canada (confirmed)", "North America (likely)",
           "UK (likely)"],
          [f'=COUNTIF({reg},"UK")', f'=COUNTIF({reg},"US")', f'=COUNTIF({reg},"CA")',
           f'=COUNTIF({reg},"US/CA?")', f'=COUNTIF({reg},"UK?")'],
          "D13", "Where the shortlist is based", "Country", "Streamers")

    fol = rng("followers")
    block(21, "How big they are",
          ["100-499", "500-999", "1k-2.5k", "2.5k-5k", "5k-10k"],
          [f'=COUNTIFS({fol},">=100",{fol},"<500")',
           f'=COUNTIFS({fol},">=500",{fol},"<1000")',
           f'=COUNTIFS({fol},">=1000",{fol},"<2500")',
           f'=COUNTIFS({fol},">=2500",{fol},"<5000")',
           f'=COUNTIFS({fol},">=5000",{fol},"<=10000")'],
          "K13", "Followers", "Follower band", "Streamers")

    gv, sp = rng("does_giveaways"), rng("has_sponsor_signal")
    block(29, "How ready they are for Fracton",
          ["Giveaways + sponsors", "Giveaways only", "Sponsors only", "Neither yet"],
          [f'=COUNTIFS({gv},"yes",{sp},"yes")', f'=COUNTIFS({gv},"yes",{sp},"")',
           f'=COUNTIFS({gv},"",{sp},"yes")', f'=COUNTIFS({gv},"",{sp},"")'],
          "D21", "Giveaway and sponsor experience", "Streamers", "Experience",
          horizontal=True)

    if brands:
        ws["A36"] = "Brands already appearing in the shortlist"
        ws["A36"].font = bold
        ws["A37"], ws["B37"], ws["C37"] = "Brand", "Streamers", "Who"
        for c in "ABC":
            ws[f"{c}37"].font = bold
        for i, b in enumerate(brands[:12], start=38):
            ws[f"A{i}"], ws[f"C{i}"] = b["brand"], b["streamers"]
            ws[f"B{i}"] = f'=COUNTIF({rng("brands_seen")},"*"&$A{i}&"*")'
            for c in "ABC":
                ws[f"{c}{i}"].font = arial
        ch = BarChart()
        ch.type, ch.legend, ch.title = "bar", None, "Brands these streamers work with"
        ch.height, ch.width = 8, 10
        n = min(12, len(brands))
        ch.add_data(Reference(ws, min_col=2, min_row=38, max_row=37 + n))
        ch.set_categories(Reference(ws, min_col=1, min_row=38, max_row=37 + n))
        ch.x_axis.title, ch.y_axis.title = "Streamers", "Brand"
        ch.x_axis.delete, ch.y_axis.delete = False, False
        ws.add_chart(ch, "K21")

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 40

    # --- Read me -----------------------------------------------------
    rm = wb.create_sheet("Read me", 0)
    rm.sheet_view.showGridLines = False
    rm.column_dimensions["A"].width = 30
    rm.column_dimensions["B"].width = 95
    r = [1]

    def line(a="", b="", style=""):
        rm.cell(r[0], 1, a)
        rm.cell(r[0], 2, b)
        if style == "title":
            rm.cell(r[0], 1).font = Font(name="Arial", bold=True, size=15)
        elif style == "head":
            rm.cell(r[0], 1).font = Font(name="Arial", bold=True, size=11, color="6441A5")
            rm.cell(r[0], 1).fill = light
            rm.cell(r[0], 2).fill = light
        else:
            rm.cell(r[0], 1).font = bold
            rm.cell(r[0], 2).font = arial
        rm.cell(r[0], 2).alignment = Alignment(wrap_text=True, vertical="top")
        r[0] += 1

    line("Fracton - Twitch candidate research", "", "title")
    line("Prepared by", "Makuo - Fracton internship")
    line("Generated", f"{datetime.now(timezone.utc):%d %B %Y, %H:%M} UTC")
    line("Scope", "English-language Twitch streamers based in the UK, US or Canada, "
                  "sized for a POC test the platform can host comfortably.")
    line()
    line("WHAT THIS IS", "", "head")
    line("Source", "Twitch's official API (Helix). No scraping, no third-party data. "
                   "Live streams are sampled repeatedly and stored, so the figures are "
                   "observed over time rather than a single snapshot.")
    line("Sampling", f"{sweep_days} day(s) of scanning so far. Every scan records all live "
                     "English streams inside the viewer range, then looks up each channel's "
                     "profile and follower count.")
    line("Channels scanned", f"{len(rows)} channels met the viewer range at least twice.")
    line("Shortlisted", f"=COUNTA(Shortlist!$A$2:$A${last})")
    line()
    line("WHO MAKES THE SHORTLIST", "", "head")
    line("Followers", f"Between {FOLLOWERS_MIN:,} and {FOLLOWERS_MAX:,}. The brief is small "
                      "creators, not big names.")
    line("Average viewers", f"Between {VIEWERS_MIN} and {VIEWERS_MAX}. The upper limit is the "
                            "load the platform can carry during a test - confirm it with the "
                            "tech team and re-run if it changes.")
    line("Consistency", f"Seen live on at least {MIN_DAYS_SEEN} different days.")
    line("Location", "UK, US or Canada. 'UK' or 'US' means their own profile says so; "
                     "'UK?' or 'US/CA?' means it is inferred from the hours they go live "
                     "and still needs a manual check.")
    line("Excluded", "24/7 reruns, watch parties, AFK and badge-farm channels (they show "
                     "high viewer counts with no real community), plus anyone whose profile "
                     "points outside the target market. Reasons are listed per channel on "
                     "the All channels sheet.")
    line()
    line("HOW THE FIT SCORE WORKS", "", "head")
    line("Out of 100", "Engagement 30 - average viewers as a share of followers, capped at 10%. "
                       "Consistency 20 - days seen live against days scanned. "
                       "Giveaway and sponsor experience 20. "
                       "Location certainty 15. "
                       "Audience size 15 - best marks for 20 to 80 average viewers.")
    line("Every row shows its own breakdown", "See the score_breakdown column, so the ranking "
                                              "can be challenged rather than trusted blindly.")
    line()
    line("WHAT EACH SHEET CONTAINS", "", "head")
    line("Dashboard", "Headline numbers and charts, all live formulas over the Shortlist sheet.")
    line("Top targets", "The 40 highest-scoring streamers with their pitch angle. Start here.")
    line("Shortlist", "Every streamer who passed all filters, full detail.")
    line("All channels", "Everyone scanned, including those excluded, with the reason why.")
    line("Brands", "Brands visible in these streamers' profiles and titles, with the names of "
                   "the streamers attached to each one.")
    line()
    line("COLUMN GLOSSARY", "", "head")
    for k, v in [
        ("fit_score", "0-100 ranking score, see above."),
        ("region / location_confidence / location_evidence",
         "Where they appear to be, how sure we are, and the exact words or hours behind it."),
        ("status", "Twitch status: none, affiliate or partner."),
        ("followers / avg_viewers / peak_viewers",
         "Follower total, plus viewers averaged and peaked across every sighting."),
        ("engagement_pct", "Average viewers as a percentage of followers. The key quality "
                           "signal: a small channel with a high figure has a loyal community."),
        ("days_seen_live / streams_seen / avg_stream_hours",
         "Consistency measures. Long average hours flag 24/7 style channels."),
        ("typical_start_utc", "When they usually go live, in UTC. Use it to time outreach."),
        ("does_giveaways / has_sponsor_signal / discount_codes / brands_seen",
         "Evidence they already run giveaways or work with brands, taken from their bio, "
         "stream titles and tags."),
        ("pitch_angle", "Suggested opening line for that specific streamer, based on their "
                        "own signals."),
        ("excluded_reason", "Why a channel did not make the shortlist (All channels sheet)."),
    ]:
        line(k, v)
    line()
    line("LIMITS TO BE AWARE OF", "", "head")
    line("Location is inferred", "Twitch does not publish a creator's country. Confirmed rows "
                                 "come from their own words; the rest are informed guesses and "
                                 "should be verified before outreach.")
    line("Sponsorships are under-counted", "Most sponsor deals live in Twitch panels and on X, "
                                           "which the API does not expose. What is flagged here "
                                           "is real, but absence of a flag proves nothing.")
    line("Only live channels are visible", "The API lists streams that are live at the moment "
                                           "of each scan, so more days of scanning means a "
                                           "fuller and fairer picture.")
    line("Giveaway rules differ by country", "US sweepstakes law is state by state and the UK "
                                             "has its own prize-draw rules. Worth legal input "
                                             "before the campaign wording is fixed.")

    wb.save(XLSX_FILE)
    print(f"  {os.path.basename(XLSX_FILE)}: {len(short)} shortlisted, {len(rows)} scanned, "
          f"{len(brands)} brands.")


# ----------------------------- HTML board ---------------------------
def board(db):
    rows, sweep_days = build_rows(db)
    short = shortlist_of(rows)
    if not short:
        sys.exit("No shortlisted streamers yet - run more collect sweeps first.")
    payload = json.dumps({"generated": datetime.now(timezone.utc).strftime("%d %b %Y %H:%M UTC"),
                          "sweep_days": sweep_days,
                          "scanned": len(rows),
                          "streamers": short,
                          "brands": brand_table(short)}, ensure_ascii=False)
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(HTML_TEMPLATE.replace("/*DATA*/", payload))
    print(f"  {os.path.basename(HTML_FILE)}: {len(short)} streamers. "
          "Open it in your browser (double-click the file).")


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fracton - streamer board</title>
<style>
:root{--bg:#0e0b16;--card:#191426;--line:#2c2440;--txt:#efeaf7;--dim:#9b90b5;
--acc:#a970ff;--ok:#3fbf7f;--warn:#e0a458}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--txt);font:15px/1.5 "Inter",system-ui,-apple-system,sans-serif}
header{padding:22px 26px;border-bottom:1px solid var(--line);position:sticky;top:0;background:var(--bg);z-index:5}
h1{margin:0;font-size:20px}
.sub{color:var(--dim);font-size:12px;margin-top:4px}
.stats{display:flex;gap:26px;margin-top:14px;flex-wrap:wrap}
.stat b{display:block;font-size:22px}.stat span{color:var(--dim);font-size:11px;text-transform:uppercase;letter-spacing:.06em}
.controls{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}
input,select{background:var(--card);color:var(--txt);border:1px solid var(--line);border-radius:8px;padding:8px 11px;font-size:13px}
input{min-width:230px}
main{padding:22px 26px;display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:15px;cursor:pointer;transition:.15s}
.card:hover{border-color:var(--acc)}
.card.done{opacity:.5}
.top{display:flex;justify-content:space-between;align-items:flex-start;gap:10px}
.name{font-weight:600;font-size:16px}
.score{background:var(--acc);color:#140f20;font-weight:700;border-radius:8px;padding:3px 9px;font-size:13px}
.meta{color:var(--dim);font-size:12px;margin-top:3px}
.tags{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}
.tag{font-size:11px;border:1px solid var(--line);border-radius:999px;padding:2px 9px;color:var(--dim)}
.tag.on{border-color:var(--ok);color:var(--ok)}
.tag.brand{border-color:var(--warn);color:var(--warn)}
.nums{display:flex;gap:16px;margin-top:11px;font-size:13px}
.nums div span{color:var(--dim);font-size:11px;display:block}
.detail{display:none;margin-top:13px;border-top:1px solid var(--line);padding-top:11px;font-size:13px;cursor:auto}
.card.open .detail{display:block}
.detail h4{margin:12px 0 4px;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--dim)}
.pitch{background:#221a33;border-left:3px solid var(--acc);padding:10px;border-radius:0 8px 8px 0}
.row{display:flex;gap:8px;flex-wrap:wrap;margin-top:11px}
a.btn,button{background:var(--acc);color:#140f20;border:0;border-radius:8px;padding:7px 12px;font-size:12px;font-weight:600;text-decoration:none;cursor:pointer}
button.ghost{background:transparent;border:1px solid var(--line);color:var(--dim)}
.empty{color:var(--dim);padding:40px;grid-column:1/-1;text-align:center}
</style></head><body>
<header>
  <h1>Fracton - Twitch POC candidates</h1>
  <div class="sub" id="sub"></div>
  <div class="stats" id="stats"></div>
  <div class="controls">
    <input id="q" placeholder="Search name, game, bio, brand...">
    <select id="region"><option value="">All regions</option><option>UK</option><option>US</option><option>CA</option><option value="likely">Likely only (unconfirmed)</option></select>
    <select id="signal"><option value="">Any experience</option><option value="giveaway">Does giveaways</option><option value="sponsor">Has sponsor signals</option><option value="both">Giveaways + sponsors</option></select>
    <select id="status"><option value="">Any Twitch status</option><option>none</option><option>affiliate</option><option>partner</option></select>
    <select id="sort"><option value="fit_score">Sort: fit score</option><option value="engagement_pct">Sort: engagement %</option><option value="avg_viewers">Sort: viewers</option><option value="followers">Sort: followers</option></select>
    <select id="mine"><option value="">All</option><option value="todo">Not contacted</option><option value="contacted">Contacted</option><option value="interested">Interested</option><option value="no">Not a fit</option></select>
    <button class="ghost" id="export">Export my picks (CSV)</button>
  </div>
</header>
<main id="grid"></main>
<script>
const DATA = /*DATA*/;
let marks = {};
try { marks = JSON.parse(localStorage.getItem("fracton_marks") || "{}"); } catch(e) {}
function saveMarks(){ try { localStorage.setItem("fracton_marks", JSON.stringify(marks)); } catch(e) {} }
const esc = s => String(s==null?"":s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));

document.getElementById("sub").textContent =
  DATA.generated + " - " + DATA.scanned + " channels scanned over " + DATA.sweep_days + " day(s)";
const S = DATA.streamers;
const stats = [
  ["Shortlisted", S.length],
  ["Location confirmed", S.filter(s=>s.location_confidence==="high").length],
  ["Do giveaways", S.filter(s=>s.does_giveaways).length],
  ["Sponsor signals", S.filter(s=>s.has_sponsor_signal).length],
  ["Median viewers", S.map(s=>s.avg_viewers).sort((a,b)=>a-b)[Math.floor(S.length/2)]],
  ["Combined viewers", Math.round(S.reduce((a,s)=>a+s.avg_viewers,0))]
];
document.getElementById("stats").innerHTML = stats.map(([k,v])=>
  `<div class="stat"><b>${v}</b><span>${k}</span></div>`).join("");

function matches(s){
  const q = document.getElementById("q").value.toLowerCase();
  if(q && !((s.streamer+" "+s.main_games+" "+s.bio+" "+s.brands_seen+" "+s.recent_title).toLowerCase().includes(q))) return false;
  const r = document.getElementById("region").value;
  if(r === "likely" && s.location_confidence === "high") return false;
  if(r && r !== "likely" && !s.region.startsWith(r)) return false;
  const sig = document.getElementById("signal").value;
  if(sig === "giveaway" && !s.does_giveaways) return false;
  if(sig === "sponsor" && !s.has_sponsor_signal) return false;
  if(sig === "both" && !(s.does_giveaways && s.has_sponsor_signal)) return false;
  const st = document.getElementById("status").value;
  if(st && s.status !== st) return false;
  const mine = document.getElementById("mine").value;
  if(mine === "todo" && marks[s.login]) return false;
  if(mine && mine !== "todo" && marks[s.login] !== mine) return false;
  return true;
}

function render(){
  const key = document.getElementById("sort").value;
  const list = S.filter(matches).sort((a,b)=>b[key]-a[key]);
  const grid = document.getElementById("grid");
  if(!list.length){ grid.innerHTML = '<div class="empty">Nothing matches those filters.</div>'; return; }
  grid.innerHTML = list.map(s => `
  <div class="card ${marks[s.login]==="no"?"done":""}" data-login="${esc(s.login)}">
    <div class="top">
      <div><div class="name">${esc(s.streamer)}</div>
        <div class="meta">${esc(s.region)} · ${esc(s.status)} · live ~${esc(s.typical_start_utc)} UTC</div></div>
      <div class="score">${s.fit_score}</div>
    </div>
    <div class="nums">
      <div>${s.followers.toLocaleString()}<span>followers</span></div>
      <div>${s.avg_viewers}<span>avg viewers</span></div>
      <div>${s.engagement_pct}%<span>engagement</span></div>
      <div>${s.days_seen_live}<span>days live</span></div>
    </div>
    <div class="tags">
      ${s.does_giveaways?'<span class="tag on">giveaways</span>':""}
      ${s.has_sponsor_signal?'<span class="tag on">sponsor signals</span>':""}
      ${s.brands_seen?`<span class="tag brand">${esc(s.brands_seen)}</span>`:""}
      ${marks[s.login]?`<span class="tag">${esc(marks[s.login])}</span>`:""}
    </div>
    <div class="detail">
      <h4>Pitch angle</h4><div class="pitch">${esc(s.pitch_angle)}</div>
      <h4>Why this score</h4><div>${esc(s.score_breakdown)}</div>
      <h4>Location evidence</h4><div>${esc(s.location_evidence) || "none - hours only"}</div>
      <h4>Plays</h4><div>${esc(s.main_games) || "-"}</div>
      ${s.discount_codes?`<h4>Codes in titles</h4><div>${esc(s.discount_codes)}</div>`:""}
      <h4>Bio</h4><div>${esc(s.bio) || "-"}</div>
      <h4>Latest title</h4><div>${esc(s.recent_title)}</div>
      <div class="row">
        <a class="btn" href="${esc(s.twitch_url)}" target="_blank" rel="noopener">Twitch</a>
        <a class="btn" href="${esc(s.sullygnome_url)}" target="_blank" rel="noopener">SullyGnome</a>
        <button class="ghost" data-mark="contacted">Contacted</button>
        <button class="ghost" data-mark="interested">Interested</button>
        <button class="ghost" data-mark="no">Not a fit</button>
        <button class="ghost" data-mark="">Clear</button>
      </div>
    </div>
  </div>`).join("");
}

document.getElementById("grid").addEventListener("click", e => {
  const card = e.target.closest(".card"); if(!card) return;
  const mark = e.target.getAttribute("data-mark");
  if(mark !== null){ mark ? marks[card.dataset.login] = mark : delete marks[card.dataset.login];
    saveMarks(); render(); return; }
  if(e.target.closest("a")) return;
  card.classList.toggle("open");
});
["q","region","signal","status","sort","mine"].forEach(id =>
  document.getElementById(id).addEventListener("input", render));
document.getElementById("export").addEventListener("click", () => {
  const cols = ["streamer","my_status","fit_score","region","followers","avg_viewers",
    "engagement_pct","status","does_giveaways","discount_codes","brands_seen",
    "typical_start_utc","twitch_url","pitch_angle"];
  const picked = S.filter(s => marks[s.login]);
  if(!picked.length){ alert("Mark some streamers first (Contacted / Interested / Not a fit)."); return; }
  const csv = [cols.join(",")].concat(picked.map(s => cols.map(c => {
    const v = c === "my_status" ? marks[s.login] : s[c];
    return '"' + String(v==null?"":v).replace(/"/g,'""') + '"';
  }).join(","))).join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], {type:"text/csv"}));
  a.download = "fracton_my_picks.csv"; a.click();
});
render();
</script></body></html>
"""


# ----------------------------- change notes -------------------------
def notes(db):
    """Compare this run's shortlist with the previous one and write a change note."""
    db.execute("""CREATE TABLE IF NOT EXISTS snapshots (
        run_at TEXT, login TEXT, streamer TEXT, fit_score INTEGER, region TEXT,
        followers INTEGER, avg_viewers REAL, engagement REAL, giveaways TEXT,
        codes TEXT, PRIMARY KEY (run_at, login))""")
    rows, sweep_days = build_rows(db)
    short = shortlist_of(rows)
    now = datetime.now(timezone.utc)
    stamp = now.strftime("%Y-%m-%dT%H:%M")

    prev_run = db.execute("SELECT run_at FROM snapshots ORDER BY run_at DESC LIMIT 1").fetchone()
    prev = {}
    if prev_run:
        prev = {r[0]: r for r in db.execute(
            "SELECT login, streamer, fit_score, region, followers, avg_viewers, engagement, "
            "giveaways, codes FROM snapshots WHERE run_at = ?", (prev_run[0],))}

    for r in short:
        db.execute("INSERT OR REPLACE INTO snapshots VALUES (?,?,?,?,?,?,?,?,?,?)", (
            stamp, r["login"], r["streamer"], r["fit_score"], r["region"], r["followers"],
            r["avg_viewers"], r["engagement_pct"], r["does_giveaways"], r["discount_codes"]))
    db.commit()

    now_by_login = {r["login"]: r for r in short}
    added = [r for r in short if r["login"] not in prev]
    left = [prev[k] for k in prev if k not in now_by_login]
    moved, new_signals = [], []
    for login, r in now_by_login.items():
        p = prev.get(login)
        if not p:
            continue
        if abs(r["fit_score"] - p[2]) >= 5:
            moved.append((r, p[2]))
        if (r["does_giveaways"] and not p[7]) or (r["discount_codes"] and not p[8]):
            new_signals.append(r)

    L = [f"# Scan {now:%d %b %Y %H:%M} UTC", ""]
    L.append(f"**{len(short)} on the shortlist** out of {len(rows)} channels tracked, "
             f"over {sweep_days} day(s) of scanning. "
             f"{sum(1 for r in short if r['location_confidence'] == 'high')} with confirmed "
             f"location, {sum(1 for r in short if r['does_giveaways'])} running giveaways, "
             f"{round(sum(r['avg_viewers'] for r in short))} combined live viewers.")
    L.append("")
    if not prev:
        L.append("First recorded scan - nothing to compare against yet.")
    if added:
        L.append(f"## New on the shortlist ({len(added)})")
        L.append("")
        L.append("| Score | Streamer | Where | Followers | Avg viewers | Engagement | Signals |")
        L.append("|---|---|---|---|---|---|---|")
        for r in sorted(added, key=lambda r: -r["fit_score"])[:25]:
            sig = ", ".join(x for x in [
                "giveaways" if r["does_giveaways"] else "",
                f"code {r['discount_codes']}" if r["discount_codes"] else "",
                r["brands_seen"]] if x) or "-"
            L.append(f"| {r['fit_score']} | [{r['streamer']}]({r['twitch_url']}) | "
                     f"{r['region']} | {r['followers']:,} | {r['avg_viewers']} | "
                     f"{r['engagement_pct']}% | {sig} |")
        L.append("")
    if new_signals:
        L.append(f"## Newly spotted giveaway or sponsor activity ({len(new_signals)})")
        for r in new_signals:
            what = "started advertising a giveaway" if r["does_giveaways"] else ""
            if r["discount_codes"]:
                what = (what + " and " if what else "") + f"is carrying code {r['discount_codes']}"
            L.append(f"- **{r['streamer']}** {what} - {r['recent_title'][:120]}")
        L.append("")
    if moved:
        L.append("## Score moves")
        for r, old in sorted(moved, key=lambda m: -(m[0]['fit_score'] - m[1]))[:15]:
            arrow = "up" if r["fit_score"] > old else "down"
            L.append(f"- {r['streamer']}: {old} -> {r['fit_score']} ({arrow}), "
                     f"{r['avg_viewers']} avg viewers, {r['days_seen_live']} days live")
        L.append("")
    if left:
        L.append(f"## Dropped off ({len(left)})")
        L.append(", ".join(p[1] for p in left[:40]))
        L.append("")
    if not (added or left or moved or new_signals) and prev:
        L.append("No changes to the shortlist this scan.")
        L.append("")

    os.makedirs(os.path.join(BASE, "notes"), exist_ok=True)
    path = os.path.join(BASE, "notes", f"{now:%Y-%m-%d_%H%M}.md")
    text = "\n".join(L)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    with open(os.path.join(BASE, "LATEST.md"), "w", encoding="utf-8") as f:
        f.write(text)
    print(f"  notes/{os.path.basename(path)}: +{len(added)} new, -{len(left)} dropped, "
          f"{len(moved)} score moves, {len(new_signals)} new signals")
    return {"added": len(added), "left": len(left), "moved": len(moved),
            "signals": len(new_signals), "shortlist": len(short),
            "top_new": sorted(added, key=lambda r: -r["fit_score"])[:5]}


def notify(result):
    """Flag anything worth acting on. Writes data/ALERT.md, which the workflow turns into
    a GitHub issue (so it reaches you by email with no extra service). If ALERT_WEBHOOK_URL
    is set - Slack, Discord, Telegram, ntfy or anything else that accepts a POST - it also
    posts there."""
    alert_path = os.path.join(BASE, "ALERT.md")
    if os.path.exists(alert_path):
        os.remove(alert_path)
    if not (result["signals"] or result["top_new"]):
        return

    lines = [f"Shortlist is now {result['shortlist']} streamers: "
             f"{result['added']} new since the last scan, "
             f"{result['signals']} showing new giveaway or sponsor activity.", ""]
    if result["top_new"]:
        lines.append("**Worth a look:**")
        lines.append("")
        for r in result["top_new"]:
            sig = []
            if r["does_giveaways"]:
                sig.append("runs giveaways")
            if r["discount_codes"]:
                sig.append(f"code {r['discount_codes']}")
            tail = f" - {', '.join(sig)}" if sig else ""
            lines.append(f"- **[{r['streamer']}]({r['twitch_url']})** - {r['fit_score']}/100, "
                         f"{r['region']}, {r['followers']:,} followers, "
                         f"{r['avg_viewers']} avg viewers, {r['engagement_pct']}% "
                         f"engagement{tail}")
        lines.append("")
    lines.append("Full notes in `data/LATEST.md`; full data in the Excel pack and the board.")
    text = "\n".join(lines)
    with open(alert_path, "w", encoding="utf-8") as f:
        f.write(text)
    print("  ALERT.md written")

    url = os.environ.get("ALERT_WEBHOOK_URL")
    if url:
        payload = {"content": text} if "discord" in url else {"text": text}
        try:
            requests.post(url, json=payload, timeout=30)
            print("  webhook alert sent")
        except requests.exceptions.RequestException as e:
            print(f"  webhook alert failed ({type(e).__name__}) - ignoring")


def weekly(db):
    """Roll the last 7 days of scans into one summary, for a team update."""
    db.execute("""CREATE TABLE IF NOT EXISTS snapshots (
        run_at TEXT, login TEXT, streamer TEXT, fit_score INTEGER, region TEXT,
        followers INTEGER, avg_viewers REAL, engagement REAL, giveaways TEXT,
        codes TEXT, PRIMARY KEY (run_at, login))""")
    now = datetime.now(timezone.utc)
    since = (now - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M")

    runs = [r[0] for r in db.execute(
        "SELECT DISTINCT run_at FROM snapshots WHERE run_at >= ? ORDER BY run_at", (since,))]
    if not runs:
        print("  no scans recorded in the last 7 days - nothing to summarise.")
        return None
    latest = runs[-1]

    before = {r[0] for r in db.execute(
        "SELECT DISTINCT login FROM snapshots WHERE run_at < ?", (since,))}
    week = {r[0]: r for r in db.execute(
        "SELECT login, streamer, fit_score, region, followers, avg_viewers, engagement, "
        "giveaways, codes FROM snapshots WHERE run_at >= ? ORDER BY run_at", (since,))}
    current = {r[0]: r for r in db.execute(
        "SELECT login, streamer, fit_score, region, followers, avg_viewers, engagement, "
        "giveaways, codes FROM snapshots WHERE run_at = ?", (latest,))}

    new_this_week = [v for k, v in week.items() if k not in before]
    signals = [v for v in current.values() if v[7] or v[8]]
    first_scores = {}
    for r in db.execute("SELECT login, fit_score FROM snapshots WHERE run_at >= ? "
                        "ORDER BY run_at", (since,)):
        first_scores.setdefault(r[0], r[1])
    risers = [(c, first_scores[k]) for k, c in current.items()
              if k in first_scores and c[2] - first_scores[k] >= 5]

    L = [f"# Weekly scan summary - week to {now:%d %b %Y}", ""]
    L.append(f"- **{len(current)} streamers on the shortlist** right now")
    L.append(f"- **{len(new_this_week)} new** this week")
    L.append(f"- **{len(signals)}** currently showing giveaway or sponsor activity")
    L.append(f"- {len(runs)} scans completed over the week")
    L.append(f"- {round(sum(c[5] for c in current.values()))} combined live viewers across "
             "the shortlist")
    L.append("")
    reg = Counter(c[3] for c in current.values())
    L.append("**Where they are:** " + ", ".join(f"{k} {v}" for k, v in reg.most_common()))
    L.append("")
    if new_this_week:
        L.append("## Best new finds this week")
        L.append("")
        L.append("| Score | Streamer | Where | Followers | Avg viewers | Engagement |")
        L.append("|---|---|---|---|---|---|")
        for c in sorted(new_this_week, key=lambda c: -c[2])[:15]:
            L.append(f"| {c[2]} | [{c[1]}](https://twitch.tv/{c[0]}) | {c[3]} | "
                     f"{c[4]:,} | {c[5]} | {c[6]}% |")
        L.append("")
    if signals:
        L.append("## Giveaway and sponsor activity")
        for c in signals[:20]:
            what = []
            if c[7]:
                what.append("runs giveaways")
            if c[8]:
                what.append(f"code {c[8]}")
            L.append(f"- **{c[1]}** ({c[3]}, {c[4]:,} followers) - {', '.join(what)}")
        L.append("")
    if risers:
        L.append("## Climbing the list")
        for c, old in sorted(risers, key=lambda x: -(x[0][2] - x[1]))[:10]:
            L.append(f"- {c[1]}: {old} -> {c[2]}")
        L.append("")
    L.append("Full detail in `data/fracton_streamers.xlsx` and `data/fracton_board.html`.")

    text = "\n".join(L)
    with open(os.path.join(BASE, "WEEKLY.md"), "w", encoding="utf-8") as f:
        f.write(text)
    print(f"  WEEKLY.md: {len(current)} shortlisted, {len(new_this_week)} new this week, "
          f"{len(signals)} with signals")
    return text


# ----------------------------- main ---------------------------------
def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "login":
        login()
    elif cmd == "collect":
        api, db = Twitch(), open_db()
        collect(api, db)
        enrich(api, db)
    elif cmd == "loop":
        minutes = int(sys.argv[2]) if len(sys.argv) > 2 else 60
        api, db = Twitch(), open_db()
        while True:
            try:
                collect(api, db)
                enrich(api, db)
            except Exception as e:
                print(f"  Sweep failed ({e}); will retry next round.")
            print(f"  Sleeping {minutes} min. Ctrl+C to stop.\n")
            time.sleep(minutes * 60)
    elif cmd == "report":
        report(open_db())
    elif cmd == "excel":
        excel(open_db())
    elif cmd == "board":
        board(open_db())
    elif cmd == "notes":
        notes(open_db())
    elif cmd == "weekly":
        weekly(open_db())
    elif cmd == "ci":
        # one unattended run: sweep, enrich, write notes and all outputs
        api, db = Twitch(), open_db()
        collect(api, db)
        enrich(api, db)
        keep_days = int(os.environ.get("KEEP_DAYS", "30"))
        cutoff = (datetime.now(timezone.utc) - timedelta(days=keep_days)).strftime("%Y-%m-%d")
        dropped = db.execute("DELETE FROM sightings WHERE seen_at < ?", (cutoff,)).rowcount
        db.commit()
        if dropped:
            db.execute("VACUUM")
            print(f"  pruned {dropped} sightings older than {keep_days} days")
        result = notes(db)
        report(db)
        try:
            excel(db)
        except SystemExit as e:
            print(f"  excel skipped: {e}")
        try:
            board(db)
        except SystemExit as e:
            print(f"  board skipped: {e}")
        summary = os.environ.get("GITHUB_STEP_SUMMARY")
        if summary and os.path.exists(os.path.join(BASE, "LATEST.md")):
            with open(os.path.join(BASE, "LATEST.md")) as src, open(summary, "a") as dst:
                dst.write(src.read())
        notify(result)
    elif cmd == "all":
        db = open_db()
        report(db)
        excel(db)
        board(db)
    else:
        print(__doc__)


if __name__ == "__main__":
    main()