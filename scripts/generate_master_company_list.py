"""
Generate data-raw/pakistan_tech_companies_master.csv (100+ Pakistan tech companies).

Run:
  python scripts/generate_master_company_list.py
  python scripts/generate_master_company_list.py --discover --target 120
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

import uni_data_free_google as udg  # noqa: E402

PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT = PROJECT_ROOT / "data-raw" / "pakistan_tech_companies_master.csv"

# name, city, www (empty www = find during scrape)
MASTER: list[tuple[str, str, str]] = [
    ("Systems Limited", "Lahore", "https://www.systemsltd.com"),
    ("NetSol Technologies", "Lahore", "https://www.netsoltech.com"),
    ("10Pearls", "Karachi", "https://www.10pearls.com"),
    ("Techlogix", "Karachi", "https://www.techlogix.com"),
    ("Arbisoft", "Lahore", "https://arbisoft.com"),
    ("Folio3", "Karachi", "https://www.folio3.com"),
    ("Confiz", "Karachi", "https://www.confiz.com"),
    ("VentureDive", "Karachi", "https://venturedive.com"),
    ("Tkxel", "Lahore", "https://www.tkxel.com"),
    ("Devsinc", "Karachi", "https://www.devsinc.com"),
    ("TRG Pakistan", "Karachi", "https://www.trg.com.pk"),
    ("Ovex Technologies", "Karachi", "https://www.ovextech.com"),
    ("Tintash", "Lahore", "https://tintash.com"),
    ("Contour Software", "Karachi", "https://www.contoursoftware.com"),
    ("Avanza Solutions", "Karachi", "https://www.avanzasolutions.com"),
    ("Nextbridge", "Lahore", "https://nextbridge.com"),
    ("InvoZone", "Lahore", "https://invozone.com"),
    ("Cubix", "Karachi", "https://www.cubix.co"),
    ("CodeNinja", "Lahore", "https://www.codeninja.co"),
    ("Trango Tech", "Karachi", ""),
    ("UrApptech", "Karachi", ""),
    ("Zamratech", "Karachi", ""),
    ("Genetech Solutions", "Karachi", "https://www.genetech.co"),
    ("LinkitSoft", "Karachi", ""),
    ("Bleed AI", "Karachi", ""),
    ("Futurealiti", "Karachi", ""),
    ("Digital Elliptical", "Karachi", ""),
    ("ZAPTA Technologies", "Lahore", ""),
    ("BearPlex", "Lahore", ""),
    ("Square63", "Lahore", "https://square63.co"),
    ("Bridge Zones", "Islamabad", ""),
    ("Esketchers", "Lahore", ""),
    ("Mobizion", "Daska", ""),
    ("3axTech", "Karachi", ""),
    ("5StarDesigners", "Karachi", ""),
    ("Objects", "Karachi", ""),
    ("CodesOrbit", "Islamabad", ""),
    ("Progatix", "Karachi", "https://progatix.com"),
    ("OSITS", "Karachi", "https://osits.pk"),
    ("Arpatech", "Karachi", "https://www.arpatech.com"),
    ("Macrosoft", "Lahore", ""),
    ("Jaffer Business Systems", "Karachi", "https://www.jbs.com.pk"),
    ("Rolustech", "Lahore", "https://www.rolustech.com"),
    ("Pure Logics", "Lahore", "https://purelogics.net"),
    ("Programmers Force", "Lahore", ""),
    ("Softoo", "Islamabad", "https://www.softoo.com"),
    ("TechAbout", "Islamabad", "https://techabout.com"),
    ("FiveRivers Technologies", "Lahore", ""),
    ("NorthBay Solutions", "Lahore", "https://northbaysolutions.com"),
    ("Emumba", "Islamabad", "https://emumba.com"),
    ("KoderLabs", "Lahore", "https://koderlabs.com"),
    ("KalSoft", "Karachi", "https://www.kalsoft.com"),
    ("Cybernet", "Karachi", "https://www.cyber.net.pk"),
    ("Creative Chaos", "Karachi", "https://www.creativechaos.co"),
    ("Clustox", "Karachi", "https://www.clustox.com"),
    ("ResourceInn", "Lahore", ""),
    ("Retailo", "Karachi", ""),
    ("Speridian", "Karachi", ""),
    ("Wavetec", "Karachi", "https://www.wavetec.com"),
    ("Xgrid", "Lahore", "https://xgrid.co"),
    ("Zepto Systems", "Lahore", ""),
    ("AlphaBOLD", "Islamabad", ""),
    ("Alchemy Technologies", "Islamabad", ""),
    ("Gaditek", "Karachi", "https://www.gaditek.com"),
    ("Mindstorm Studios", "Lahore", "https://mindstormstudios.com"),
    ("Ibex Digital", "Lahore", ""),
    ("Hazara IT", "Islamabad", ""),
    ("Intagleo Systems", "Islamabad", ""),
    ("One Byte", "Karachi", ""),
    ("RevCrew", "Karachi", ""),
    ("Sapphire Software Solutions", "Lahore", ""),
    ("SecureTech", "Islamabad", ""),
    ("Sybrid", "Karachi", "https://www.sybrid.com"),
    ("Terafort", "Lahore", ""),
    ("VaporVM", "Karachi", ""),
    ("Visnext", "Lahore", ""),
    ("WeGo Technologies", "Karachi", ""),
    ("Zaeem Technologies", "Lahore", ""),
    ("7Vals", "Lahore", ""),
    ("Abacus Consulting", "Karachi", "https://www.abacus.com.pk"),
    ("Astek Pakistan", "Karachi", ""),
    ("Concept Software", "Karachi", ""),
    ("F3 Technologies", "Islamabad", ""),
    ("Genesis Solutions", "Karachi", ""),
    ("Hexalyze", "Islamabad", ""),
    ("Innovent Engineering", "Lahore", ""),
    ("LMKR", "Islamabad", "https://www.lmkr.com"),
    ("Nexskill", "Islamabad", ""),
    ("Opal Labs", "Lahore", ""),
    ("Rafay Systems", "Islamabad", ""),
    ("Stack Bench", "Karachi", ""),
    ("Afiniti", "Karachi", ""),
    ("i2c Inc", "Lahore", "https://www.i2cinc.com"),
    ("Deloitte Pakistan Technology", "Karachi", ""),
    ("IBM Pakistan", "Karachi", ""),
    ("Ericsson Pakistan", "Islamabad", ""),
    ("M TECHUB LLC", "Rawalpindi", ""),
    ("Sam's Solution", "Islamabad", ""),
    ("System Plus", "Lahore", ""),
    ("Blend IT", "Karachi", ""),
    ("Bolish Technology", "Lahore", ""),
    ("Ciklum Pakistan", "Karachi", ""),
    ("Daewoo Information Technology", "Islamabad", ""),
    ("Eplanet Communications", "Karachi", ""),
    ("Galaxy IT", "Karachi", ""),
    ("Ignite Solutions", "Lahore", ""),
    ("Integra Technologies", "Karachi", ""),
    ("Interloop Limited", "Karachi", "https://www.interloop.com.pk"),
    ("ITSYS", "Karachi", ""),
    ("Learning Pitch", "Lahore", ""),
    ("Metro Infrasys", "Lahore", ""),
    ("Neusol", "Karachi", ""),
    ("Nexsoft", "Islamabad", ""),
    ("OneStack", "Karachi", ""),
    ("Optasia", "Karachi", ""),
    ("PakWheels", "Lahore", "https://www.pakwheels.com"),
    ("PIT Solutions", "Islamabad", ""),
    ("ProGaze", "Karachi", ""),
    ("SlashNext Pakistan", "Karachi", ""),
    ("Switch Communications", "Karachi", ""),
    ("Techloset", "Islamabad", ""),
    ("Triotech Digital", "Karachi", ""),
    ("Uls CJ Solutions", "Lahore", ""),
    ("Viper Technology", "Karachi", ""),
    ("Westech Solutions", "Lahore", ""),
    ("Wynde Solutions", "Karachi", ""),
    ("Xavor Corporation", "Islamabad", "https://www.xavor.com"),
    ("Zurno", "Lahore", ""),
]

DISCOVER_QUERIES = [
    "site:linkedin.com/company Pakistan software house",
    "site:linkedin.com/company Pakistan IT company",
    "site:linkedin.com/company Karachi software",
    "site:linkedin.com/company Lahore technology company",
    "site:linkedin.com/company Islamabad software",
    "site:linkedin.com/company Rawalpindi IT",
    "site:linkedin.com/company Pakistan fintech software",
    "site:linkedin.com/company Pakistan mobile app development",
    "site:linkedin.com/company Pakistan web development agency",
    "site:linkedin.com/company Pakistan SaaS",
    "site:linkedin.com/company Pakistan enterprise software",
    "site:linkedin.com/company Pakistan cloud solutions",
    "site:linkedin.com/company Pakistan AI software",
    "site:linkedin.com/company Pakistan cybersecurity company",
    "site:linkedin.com/company Pakistan ecommerce development",
    "site:linkedin.com/company Pakistan software outsourcing",
    "site:linkedin.com/company Pakistan digital transformation",
    "site:linkedin.com/company Faisalabad software",
    "site:linkedin.com/company Multan IT company",
    "site:linkedin.com/company Peshawar software",
    "site:linkedin.com/company Pakistan game development studio",
    "site:linkedin.com/company Pakistan ERP software",
    "site:linkedin.com/company Pakistan data analytics company",
    "site:linkedin.com/company Pakistan DevOps services",
    "site:linkedin.com/company Pakistan blockchain development",
]

SKIP_NAME_FRAGMENTS = (
    "linkedin",
    "microsoft",
    "google",
    "amazon",
    "facebook",
    "university",
    "college",
    "school",
    "ministry",
    "government",
    "pakistan army",
    "freelancer",
    "self-employed",
    "stealth",
    "confidential",
)


def slug_to_name(slug: str) -> str:
    name = re.sub(r"[-_]+", " ", slug).strip()
    name = re.sub(r"\s+ltd$|\s+limited$|\s+pvt$|\s+pk$", "", name, flags=re.I).strip()
    return name.title() if name else ""


def is_valid_company_name(name: str) -> bool:
    if len(name) < 3 or len(name) > 80:
        return False
    low = name.lower()
    if any(x in low for x in SKIP_NAME_FRAGMENTS):
        return False
    if sum(c.isdigit() for c in name) > 4:
        return False
    return True


def discover_linkedin_companies(session: requests.Session, target: int, engine: str) -> list[tuple[str, str, str]]:
    seen: set[str] = set()
    rows: list[tuple[str, str, str]] = []
    for query in DISCOVER_QUERIES:
        if len(rows) >= target:
            break
        try:
            urls, _ = udg.search_urls(session, query, max_results=25, engine=engine)
        except Exception:
            continue
        for url in urls:
            if "linkedin.com/company/" not in url.lower():
                continue
            slug = urlparse(url).path.strip("/").split("/")[-1]
            name = slug_to_name(slug)
            key = name.lower()
            if not is_valid_company_name(name) or key in seen:
                continue
            seen.add(key)
            city = ""
            if "karachi" in query.lower():
                city = "Karachi"
            elif "lahore" in query.lower():
                city = "Lahore"
            elif "islamabad" in query.lower():
                city = "Islamabad"
            rows.append((name, city, url.split("?")[0]))
            if len(rows) >= target:
                break
        time.sleep(0.35)
    return rows


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--target", type=int, default=100, help="Minimum companies in master list")
    p.add_argument("--discover", action="store_true", help="Add more names from LinkedIn search")
    p.add_argument("--engine", default="auto", choices=("auto", "ddgs", "duckduckgo_html", "jina_ddg"))
    args = p.parse_args()

    records: list[dict[str, str]] = []
    seen: set[str] = set()
    for name, city, www in MASTER:
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        records.append({"name": name, "city": city, "www": www, "industry": "technology"})

    if args.discover:
        session = requests.Session()
        extra = discover_linkedin_companies(session, args.target, args.engine)
        for name, city, li_url in extra:
            key = name.lower()
            if key in seen:
                continue
            seen.add(key)
            records.append(
                {"name": name, "city": city, "www": "", "industry": "technology", "linkedin_seed": li_url}
            )

    df = pd.DataFrame(records)
    if "linkedin_seed" not in df.columns:
        df["linkedin_seed"] = ""
    df = df.drop_duplicates(subset=["name"], keep="first")
    if args.target > 0:
        df = df.head(max(args.target, 100))
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT, index=False)
    print(f"Wrote {len(df)} companies -> {OUTPUT}")
    if len(df) < args.target:
        print(f"Note: only {len(df)} rows. Run with --discover to try adding more from LinkedIn search.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
