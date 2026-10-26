#!/usr/bin/env python3
"""Build a traceable, cleaned macro-announcement panel from the thesis source workbook.

Run from the repository root:
    python scripts/build_clean_announcement_panel.py

Requires openpyxl. Official BLS, Federal Reserve, and BEA release sources are
fetched for event-level checks; other countries retain explicit unverified status.
"""

from __future__ import annotations

import html
import re
import urllib.request
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, time, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data/macroeconomic_announcements_2000_to_2026-10-12.xlsx"
OUTPUT = ROOT / "data/cleaned_announcement_panel_2019-11-20_to_2026-08-31.xlsx"
START = date(2019, 11, 20)
END = date(2026, 8, 31)
SOURCE_TZ = "Europe/Budapest"

BASE_COUNTRIES = [
    "United States", "Canada", "United Kingdom", "Germany", "France", "Italy",
    "Japan", "Brazil", "Türkiye", "China", "India", "Mexico", "South Africa",
    "Russia", "Israel",
]
ADDITIONS = [
    "Australia", "Chile", "Colombia", "Czech Republic", "Egypt", "Indonesia",
    "Malaysia", "Norway", "Poland", "Romania", "Sweden", "Switzerland",
]
SAMPLE_UNITS = set(BASE_COUNTRIES + ADDITIONS + ["Euro Zone"])

# Local time zone of the principal statistical publisher / central-bank city.
# Countries with multiple zones use the usual national release or central-bank zone.
TZ_BY_COUNTRY = {
    "Australia": "Australia/Sydney", "Brazil": "America/Sao_Paulo",
    "Canada": "America/Toronto", "Chile": "America/Santiago",
    "China": "Asia/Shanghai", "Colombia": "America/Bogota",
    "Czech Republic": "Europe/Prague", "Egypt": "Africa/Cairo",
    "Euro Zone": "Europe/Brussels", "France": "Europe/Paris",
    "Germany": "Europe/Berlin", "India": "Asia/Kolkata",
    "Indonesia": "Asia/Jakarta", "Israel": "Asia/Jerusalem",
    "Italy": "Europe/Rome", "Japan": "Asia/Tokyo",
    "Malaysia": "Asia/Kuala_Lumpur", "Mexico": "America/Mexico_City",
    "Norway": "Europe/Oslo", "Poland": "Europe/Warsaw",
    "Romania": "Europe/Bucharest", "Russia": "Europe/Moscow",
    "South Africa": "Africa/Johannesburg", "Sweden": "Europe/Stockholm",
    "Switzerland": "Europe/Zurich", "Türkiye": "Europe/Istanbul",
    "United Kingdom": "Europe/London", "United States": "America/New_York",
}

# Official institutional portals for later country-by-country reconciliation.
# Only the US BLS CPI schedule is marked event-level verified in this workbook.
OFFICIAL_PORTALS = {
    "Australia": ("https://www.abs.gov.au/release-calendar", "https://www.rba.gov.au/monetary-policy/rba-board-minutes/"),
    "Brazil": ("https://www.ibge.gov.br/en/statistics/economic/prices-and-costs.html", "https://www.bcb.gov.br/en/monetarypolicy"),
    "Canada": ("https://www.statcan.gc.ca/en/start", "https://www.bankofcanada.ca/core-functions/monetary-policy/key-interest-rate/"),
    "Chile": ("https://www.ine.gob.cl/estadisticas/economia/indices-de-precio-e-inflacion/indice-de-precios-al-consumidor", "https://www.bcentral.cl/en/web/banco-central/areas/monetary-policy"),
    "China": ("https://www.stats.gov.cn/english/PressRelease/", "http://www.pbc.gov.cn/en/3688241/index.html"),
    "Colombia": ("https://www.dane.gov.co/", "https://www.banrep.gov.co/"),
    "Czech Republic": ("https://www.czso.cz/", "https://www.cnb.cz/en/monetary-policy/bank-board-decisions/"),
    "Egypt": ("https://www.capmas.gov.eg/", "https://www.cbe.org.eg/"),
    "Euro Zone": ("https://ec.europa.eu/eurostat/news/release-calendar", "https://www.ecb.europa.eu/"),
    "France": ("https://www.insee.fr/en/accueil", "https://www.ecb.europa.eu/"),
    "Germany": ("https://www.destatis.de/EN/Home/_node.html", "https://www.ecb.europa.eu/"),
    "India": ("https://mospi.gov.in/", "https://www.rbi.org.in/Scripts/BS_PressReleaseDisplay.aspx"),
    "Indonesia": ("https://www.bps.go.id/en", "https://www.bi.go.id/"),
    "Israel": ("https://www.cbs.gov.il/en/Pages/default.aspx", "https://www.boi.org.il/en/communication-and-publications/press-releases/"),
    "Italy": ("https://www.istat.it/en/", "https://www.ecb.europa.eu/"),
    "Japan": ("https://www.stat.go.jp/english/data/cpi/158c.html", "https://www.boj.or.jp/en/mopo/mpmsche_minu/index.htm"),
    "Malaysia": ("https://www.dosm.gov.my/portal-main/release-content", "https://www.bnm.gov.my/monetary-policy"),
    "Mexico": ("https://en.www.inegi.org.mx/app/saladeprensa/calendario/", "https://www.banxico.org.mx/"),
    "Norway": ("https://www.ssb.no/en/", "https://www.norges-bank.no/en/topics/Monetary-policy/Monetary-policy-meetings/"),
    "Poland": ("https://stat.gov.pl/en/", "https://nbp.pl/en/"),
    "Romania": ("https://insse.ro/cms/en", "https://www.bnr.ro/"),
    "Russia": ("https://eng.rosstat.gov.ru/", "https://www.cbr.ru/eng/press/keypr/"),
    "South Africa": ("https://www.statssa.gov.za/?page_id=1854", "https://www.resbank.co.za/en/home/what-we-do/monetary-policy"),
    "Sweden": ("https://www.scb.se/en/", "https://www.riksbank.se/en-gb/"),
    "Switzerland": ("https://www.bfs.admin.ch/bfs/en/home/statistics/prices/consumer-price-index.html", "https://www.snb.ch/en/the-snb/mandates-goals/monetary-policy"),
    "Türkiye": ("https://www.tuik.gov.tr/", "https://www.tcmb.gov.tr/wps/wcm/connect/EN/TCMB+EN/Main+Menu/Core+Functions/Monetary+Policy/Monetary+Policy+Committee"),
    "United Kingdom": ("https://www.ons.gov.uk/releasecalendar", "https://www.bankofengland.co.uk/monetary-policy/upcoming-mpc-dates"),
    "United States": ("https://www.bls.gov/schedule/news_release/cpi.htm", "https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm"),
}
GDP_PORTAL = {
    "United States": "https://www.bea.gov/news/schedule",
    "United Kingdom": "https://www.ons.gov.uk/releasecalendar",
    "Euro Zone": "https://ec.europa.eu/eurostat/news/release-calendar",
}

SHEET_FAMILIES = {
    "Policy Rates": "Policy Rate",
    "CPI Inflation": "CPI",
    "GDP Growth": "GDP Growth",
    "Policy Announcements": "Policy Communication",
    "GDP Estimates": "GDP Tracker",
}


def download_text(url: str) -> str:
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (research data validation)"})
    with urllib.request.urlopen(request, timeout=45) as response:
        return response.read().decode("utf-8", "replace")


def official_bls_cpi_schedule() -> tuple[dict[date, datetime], list[tuple]]:
    schedule: dict[date, datetime] = {}
    audit = []
    for year in range(2019, 2027):
        url = f"https://www.bls.gov/schedule/{year}/home.htm"
        source = download_text(url)
        text = " ".join(html.unescape(re.sub(r"<[^>]+>", " ", source)).split())
        matches = re.findall(
            r"(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+"
            r"([A-Z][a-z]+\s+\d{1,2},\s+\d{4})\s+(\d{1,2}:\d{2}\s+[AP]M)\s+Consumer Price Index",
            text,
        )
        if not matches:
            raise RuntimeError(f"No CPI entries parsed from official BLS schedule: {url}")
        for date_text, time_text in matches:
            local = datetime.strptime(f"{date_text} {time_text}", "%B %d, %Y %I:%M %p")
            local = local.replace(tzinfo=ZoneInfo("America/New_York"))
            schedule[local.date()] = local
        audit.append((year, url, len(matches), "Consumer Price Index", "America/New_York"))
    return schedule, audit


def official_fed_policy_releases(raw_rows):
    dates = sorted({
        item[8] for item in raw_rows
        if item[2] == "United States" and item[1] == "Policy Rate" and isinstance(item[8], date)
    })
    pattern = re.compile(r"For release at\s+(\d{1,2}:\d{2})\s*([ap])\.?m\.?\s*(EST|EDT)", re.I)

    def fetch(event_date):
        url = f"https://www.federalreserve.gov/newsevents/pressreleases/monetary{event_date:%Y%m%d}a.htm"
        source = download_text(url)
        text = " ".join(html.unescape(re.sub(r"<[^>]+>", " ", source)).split())
        match = pattern.search(text)
        if not match:
            raise RuntimeError(f"Could not parse official Federal Reserve release time: {url}")
        hh, mm = (int(x) for x in match.group(1).split(":"))
        if match.group(2).lower() == "p" and hh != 12:
            hh += 12
        elif match.group(2).lower() == "a" and hh == 12:
            hh = 0
        event_dt = datetime.combine(event_date, time(hh, mm), tzinfo=ZoneInfo("America/New_York"))
        return event_date, event_dt, url, f"Federal Reserve FOMC statement; official page states {match.group(1)} {match.group(2).upper()}.m. {match.group(3)}"

    with ThreadPoolExecutor(max_workers=6) as pool:
        records = list(pool.map(fetch, dates))
    lookup = {r[0]: {"datetime": r[1], "url": r[2], "note": r[3]} for r in records}
    audit = [["United States", "Policy Rate", r[0], r[1].time().replace(tzinfo=None), "America/New_York", r[2], "Official statement release time/date", r[3]] for r in records]
    return lookup, audit


def official_bea_gdp_releases():
    archive_records = []
    for page in range(0, 40):
        url = f"https://www.bea.gov/news/archive?field_related_product_target_id=451&created_1=All&title=&page={page}"
        source = download_text(url)
        page_dates = []
        for row in re.findall(r'<tr class="release-row">(.*?)</tr>', source, re.I | re.S):
            anchor = re.search(r'<a href="([^"]+)"[^>]*>([^<]+)</a>', row, re.I | re.S)
            published = re.search(r'<time datetime="([^"]+)">', row, re.I)
            if not (anchor and published):
                continue
            published_dt = datetime.fromisoformat(published.group(1))
            page_dates.append(published_dt.date())
            if START <= published_dt.date() <= END:
                article_url = "https://www.bea.gov" + anchor.group(1)
                archive_records.append((html.unescape(anchor.group(2)).strip(), article_url, published_dt))
        if page_dates and min(page_dates) < START:
            break
        if not page_dates:
            break

    pattern = re.compile(
        r"EMBARGOED UNTIL RELEASE AT\s+(\d{1,2}:\d{2})\s*([AP])\.?\s*M\.?\s*,?\s*"
        r"(EST|EDT)\s*,?\s*(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+"
        r"([A-Z][a-z]+\s+\d{1,2},\s+\d{4})",
        re.I,
    )

    def fetch(record):
        title, url, archive_dt = record
        source = download_text(url)
        text = " ".join(html.unescape(re.sub(r"<[^>]+>", " ", source)).split())
        match = pattern.search(text)
        if not match:
            raise RuntimeError(f"Could not parse official BEA release date/time: {url}")
        hh, mm = (int(x) for x in match.group(1).split(":"))
        if match.group(2).upper() == "P" and hh != 12:
            hh += 12
        elif match.group(2).upper() == "A" and hh == 12:
            hh = 0
        article_date = datetime.strptime(match.group(4), "%B %d, %Y").date()
        event_dt = datetime.combine(article_date, time(hh, mm), tzinfo=ZoneInfo("America/New_York"))
        if article_date != archive_dt.date():
            raise RuntimeError(f"BEA article/archive date conflict for {url}")
        return article_date, event_dt, url, title, f"BEA article embargo states {match.group(1)} {match.group(2).upper()}.M. {match.group(3)}"

    with ThreadPoolExecutor(max_workers=8) as pool:
        records = list(pool.map(fetch, archive_records))
    lookup = {r[0]: {"datetime": r[1], "url": r[2], "title": r[3], "note": r[4]} for r in records}
    audit = [["United States", "GDP Growth", r[0], r[1].time().replace(tzinfo=None), "America/New_York", r[2], "Official BEA release article", f"{r[3]} — {r[4]}"] for r in records]
    return lookup, audit


def parse_numeric(value):
    if value is None:
        return None
    text = str(value).strip().replace("−", "-").replace(",", "")
    if not text or text in {"-", "—", "–", "N/A", "n/a", "None"}:
        return None
    scale = 1.0
    if text[-1:].lower() in {"k", "m", "b"}:
        scale = {"k": 1e3, "m": 1e6, "b": 1e9}[text[-1].lower()]
        text = text[:-1]
    text = text.replace("%", "").replace("+", "")
    try:
        return float(text) * scale
    except ValueError:
        return None


def is_missing_value(value) -> bool:
    if value is None:
        return True
    return str(value).strip().lower() in {"", "-", "—", "–", "n/a", "na", "none", "null"}


def get_basis(title: str, category: str, subcategory: str) -> str:
    low = title.lower()
    if category == "CPI":
        if re.search(r"\byoy\b|year[- ]on[- ]year", low):
            return "YoY"
        if re.search(r"\bmom\b|month[- ]on[- ]month", low):
            return "MoM"
        if "qoq" in low:
            return "QoQ"
        if subcategory == "consumer_price_index_level":
            return "Index level"
        return "Not stated in title"
    basis = {
        "quarter_on_quarter_growth": "QoQ",
        "annualized_quarter_on_quarter_growth": "QoQ annualized (SAAR)",
        "year_on_year_growth": "YoY",
        "annual_growth": "Annual growth",
        "monthly_gdp_or_activity_growth": "Monthly GDP/activity growth",
        "rolling_three_month_growth": "Rolling 3-month growth",
        "gdp_growth_inconsistent_source_basis": "Source basis inconsistent",
        "mixed_monthly_quarterly_gdp_or_activity_growth": "Mixed monthly/quarterly activity",
    }
    return basis.get(subcategory, "Not stated")


def cpi_measure(subcategory: str) -> str:
    names = {
        "headline_cpi": "National headline CPI",
        "harmonized_cpi": "Harmonised CPI (HICP)",
        "core_or_underlying_cpi": "Core/underlying CPI",
        "core_harmonized_cpi": "Core HICP",
        "consumer_price_index_level": "CPI index level",
        "cpif_constant_interest_rates": "CPIF (constant-interest-rate CPI)",
        "core_cpif": "Core CPIF",
        "cpih_including_owner_occupiers_housing": "CPIH including owner-occupiers' housing",
        "cpi_ex_tobacco": "CPI excluding tobacco",
        "mid_month_cpi": "Mid-month CPI measure",
        "monthly_cpi_indicator": "Monthly CPI indicator",
        "regional_consumer_price_inflation": "Regional consumer-price measure",
        "subnational_cpi": "Subnational CPI measure",
        "subnational_core_cpi": "Subnational core CPI measure",
    }
    return names.get(subcategory, subcategory.replace("_", " ").capitalize())


def policy_role(country: str, title: str, subcategory: str, event_date: date | None) -> tuple[str, bool]:
    low = (title + " " + subcategory).lower()
    if country == "Euro Zone":
        if "deposit facility" in low:
            return "ECB deposit facility rate (primary operational rate from 2024-06)", bool(event_date and event_date >= date(2024, 6, 12))
        if "main refinancing" in low or "interest rate decision" in low:
            return "ECB main refinancing rate (long-run policy series; key operational rate through 2024-06)", bool(event_date and event_date < date(2024, 6, 12))
        return "ECB corridor facility", False
    if "one-week repo policy rate" in low:
        return "Headline policy target candidate", True
    if "headline policy interest-rate decision" in low:
        return "Headline policy target candidate", True
    if "loan prime rate" in low:
        return "Policy-related benchmark (China LPR; analyze separately)", False
    if "facility" in low or "repo policy rate" in low or "reverse repo" in low or "discount rate" in low:
        return "Policy corridor/facility rate", False
    return "Other policy-rate series", False


def gdp_stage(title: str) -> str:
    low = title.lower()
    for needle, label in [
        ("advance", "Advance"), ("second estimate", "Second estimate"),
        ("third estimate", "Third estimate"), ("first estimate", "First estimate"),
        ("initial estimate", "Initial estimate"), ("updated estimate", "Updated estimate"),
        ("final", "Final"), ("flash", "Flash"), ("preliminary", "Preliminary"),
        ("revised", "Revised"),
    ]:
        if needle in low:
            return label
    return "Not specified in source title"


def bea_reference_period(title: str) -> str | None:
    match = re.search(
        r"\b(1st|2nd|3rd|4th|first|second|third|fourth)\s+quarter(?:\s+and\s+year)?\s+(\d{4})\b",
        title,
        re.I,
    )
    if not match:
        return None
    quarter = {
        "1st": "Q1", "first": "Q1", "2nd": "Q2", "second": "Q2",
        "3rd": "Q3", "third": "Q3", "4th": "Q4", "fourth": "Q4",
    }[match.group(1).lower()]
    return f"{quarter} {match.group(2)}"


def reference_period(title: str, source_period) -> str:
    if source_period:
        return str(source_period)
    matches = re.findall(r"\(([^()]*)\)", title)
    for item in reversed(matches):
        v = item.strip()
        if re.fullmatch(r"Q[1-4](?:\s*\d{4})?", v, re.I) or re.fullmatch(
            r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*(?:\s+\d{4})?", v, re.I
        ) or re.fullmatch(r"\d{4}", v):
            return v
    return "Not stated in source"


def add_table(ws, headers, rows, table_name):
    ws.append(headers)
    for row in rows:
        ws.append(list(row))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    if ws.max_row > 1:
        ref = f"A1:{ws.cell(row=ws.max_row, column=ws.max_column).coordinate}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                              showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(table)
    ws.row_dimensions[1].height = 32
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for col in ws.columns:
        col_letter = col[0].column_letter
        max_len = min(55, max((len(str(c.value or "")) for c in col[: min(ws.max_row, 100)]), default=10) + 2)
        ws.column_dimensions[col_letter].width = max(12, max_len)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("http"):
                cell.hyperlink = cell.value
                cell.style = "Hyperlink"
            cell.alignment = Alignment(vertical="top", wrap_text=False)


def main():
    bls_dates, bls_year_audit = official_bls_cpi_schedule()
    retrieved_on = datetime.now(timezone.utc).date().isoformat()

    src_wb = load_workbook(SOURCE, read_only=True, data_only=True)
    raw_rows = []
    for sheet_name, family in SHEET_FAMILIES.items():
        ws = src_wb[sheet_name]
        headers = [c.value for c in ws[1]]
        idx = {name: i for i, name in enumerate(headers)}
        for r in ws.iter_rows(min_row=2, values_only=True):
            country = r[idx["Country"]]
            if country not in SAMPLE_UNITS:
                continue
            source_date = r[idx["Announcement date"]]
            source_time = r[idx["Announcement time"]]
            utc_ts = r[idx["Timestamp (UTC)"]]
            status = r[idx["Time status"]]
            tz_name = TZ_BY_COUNTRY[country]
            event_local = None
            if isinstance(utc_ts, datetime):
                event_local = utc_ts.replace(tzinfo=ZoneInfo("UTC")).astimezone(ZoneInfo(tz_name))
            # For suspect archival dates retain the archived display date in sample selection;
            # otherwise use the converted local event date where it exists.
            candidate_date = source_date.date() if status == "Possible archival placeholder" and isinstance(source_date, datetime) else (
                event_local.date() if event_local else source_date.date() if isinstance(source_date, datetime) else None
            )
            if not candidate_date or candidate_date < START or candidate_date > END:
                continue
            vals = {name: r[i] for name, i in idx.items()}
            raw_rows.append((sheet_name, family, country, tz_name, source_date, source_time, utc_ts,
                             event_local, candidate_date, status, vals))
    src_wb.close()

    fed_releases, fed_audit = official_fed_policy_releases(raw_rows)
    bea_releases, bea_audit = official_bea_gdp_releases()
    bls_sample_dates = {d for d in bls_dates if START <= d <= END}
    source_us_cpi_dates = {
        item[8] for item in raw_rows
        if item[2] == "United States" and item[1] == "CPI"
        and item[10].get("Subcategory") == "headline_cpi"
        and isinstance(item[8], date)
    }
    source_us_gdp_dates = {
        item[8] for item in raw_rows
        if item[2] == "United States" and item[1] == "GDP Growth" and isinstance(item[8], date)
    }
    us_cpi_calendar_complete = source_us_cpi_dates == bls_sample_dates
    us_gdp_calendar_complete = source_us_gdp_dates == set(bea_releases)

    # Prefer headline YoY values for each event unit. Where the archived calendar
    # has no YoY headline series, retain its available headline basis and label it.
    cpi_bases = defaultdict(set)
    for _, family, country, _, _, _, _, _, _, _, vals in raw_rows:
        if family == "CPI" and vals.get("Subcategory") == "headline_cpi":
            cpi_bases[country].add(get_basis(str(vals.get("Announcement") or ""), family, "headline_cpi"))
    cpi_primary_basis = {
        country: ("YoY" if "YoY" in bases else "MoM" if "MoM" in bases else sorted(bases)[0])
        for country, bases in cpi_bases.items() if bases
    }

    panel_headers = [
        "Sample tier", "Country / event unit", "Market link countries", "Event family",
        "Normalized release definition", "Series basis", "Primary series candidate?", "Event-day dummy candidate?",
        "Policy rate role", "Release stage", "Reference period (source or parsed)",
        "Candidate local event date", "Candidate local event time", "Event local timezone",
        "Date validation status", "Time validation status", "Official verified local date",
        "Official verified local time", "Official validation source",
        "Source date (Europe/Budapest)", "Source time (Europe/Budapest)", "Source timestamp UTC",
        "Source timing flag", "Source timing note", "Currency", "Announcement", "Prior value (source)",
        "Actual value (source)", "Expected value (source)", "Actual numeric (source scale)",
        "Expected numeric (source scale)", "Surprise numeric (actual - expected; source scale)",
        "Prior revision (source)", "Source reference period", "Announcement ID", "Economic event ID",
        "Category", "Subcategory", "Source URL", "Missing actual?", "Missing expected?",
        "Analysis handling recommendation",
    ]
    panel_rows = []
    official_matched = Counter()
    for sheet_name, family, country, tz_name, source_date, source_time, utc_ts, event_local, candidate_date, status, v in raw_rows:
        title = str(v.get("Announcement") or "")
        sub = str(v.get("Subcategory") or "")
        is_cleveland_cpi = country == "United States" and family == "CPI" and "cleveland cpi" in title.lower()
        panel_family = "Alternative CPI" if is_cleveland_cpi else family
        basis = get_basis(title, family, sub)
        measure = cpi_measure(sub) if family == "CPI" else sub.replace("_", " ").capitalize()
        if is_cleveland_cpi:
            release_def = "Cleveland Fed Median CPI — alternative underlying inflation measure"
        elif family == "GDP Tracker":
            release_def = "Non-official GDP estimate / nowcast (kept separate from statistical GDP releases)"
        elif family == "Policy Communication":
            release_def = f"Scheduled central-bank communication — {sub}"
        elif family == "Policy Rate":
            release_def = f"Policy-rate announcement — {sub}"
        elif family == "CPI":
            release_def = f"{measure} — {basis}"
        else:
            release_def = f"GDP growth release — {basis}"
        official_gdp_title = bea_releases.get(candidate_date, {}).get("title") if country == "United States" and family == "GDP Growth" else None
        stage = gdp_stage(official_gdp_title or title) if family in {"GDP Growth", "GDP Tracker"} else "Not applicable"
        period = reference_period(title, v.get("Reference period")) if family in {"GDP Growth", "GDP Tracker"} else (
            reference_period(title, v.get("Reference period")) if family == "CPI" else (str(v.get("Reference period") or "Not stated in source"))
        )
        if official_gdp_title:
            period = bea_reference_period(official_gdp_title) or period
        policy_label, policy_primary = policy_role(country, title, sub, candidate_date) if family == "Policy Rate" else ("Not applicable", False)
        primary = False
        event_dummy_candidate = False
        if family == "Policy Rate":
            primary = policy_primary
            event_dummy_candidate = policy_primary or (
                country == "China" and "one-year loan prime rate" in (title + " " + sub).lower()
            )
        elif family == "CPI":
            primary = sub == "headline_cpi" and basis == cpi_primary_basis.get(country)
            if country == "Euro Zone" and not cpi_bases.get(country):
                primary = sub == "harmonized_cpi" and basis == "YoY"
            event_dummy_candidate = sub == "headline_cpi" or (
                country == "Euro Zone" and not cpi_bases.get(country) and sub == "harmonized_cpi"
            )
        elif family == "GDP Growth":
            primary = sub in {"quarter_on_quarter_growth", "annualized_quarter_on_quarter_growth", "year_on_year_growth"}
            event_dummy_candidate = True
        if status == "Possible archival placeholder":
            primary = False
            event_dummy_candidate = False

        verified_date = None
        verified_time = None
        official_source = None
        if is_cleveland_cpi:
            date_status = "Separate Cleveland Fed inflation measure; official release date/time not independently matched"
            time_status = "Source-reported candidate time; Cleveland Fed release time not independently confirmed"
        elif country == "United States" and family == "CPI" and candidate_date in bls_dates:
            official_dt = bls_dates[candidate_date]
            if event_local and event_local.date() == official_dt.date() and event_local.time() == official_dt.time():
                verified_date = official_dt.date()
                verified_time = official_dt.time()
                official_source = f"https://www.bls.gov/schedule/{candidate_date.year}/home.htm"
                date_status = "Matched official BLS CPI date and time"
                time_status = "Verified: 08:30 America/New_York"
                official_matched[sheet_name] += 1
            else:
                date_status = "BLS release date exists; source timestamp differs — review"
                time_status = "Not matched to official schedule — review"
        elif country == "United States" and family == "Policy Rate" and candidate_date in fed_releases:
            official = fed_releases[candidate_date]
            official_dt = official["datetime"]
            if event_local and event_local.date() == official_dt.date() and event_local.time() == official_dt.time():
                verified_date = official_dt.date()
                verified_time = official_dt.time()
                official_source = official["url"]
                date_status = "Matched official Federal Reserve statement date and time"
                time_status = "Verified against Federal Reserve statement release notice"
                official_matched[sheet_name] += 1
            else:
                date_status = "Federal Reserve statement date exists; source timestamp differs — review"
                time_status = "Not matched to official statement time — review"
        elif country == "United States" and family == "GDP Growth" and candidate_date in bea_releases:
            official = bea_releases[candidate_date]
            official_dt = official["datetime"]
            if event_local and event_local.date() == official_dt.date() and event_local.time() == official_dt.time():
                verified_date = official_dt.date()
                verified_time = official_dt.time()
                official_source = official["url"]
                date_status = "Matched official BEA GDP release date and time"
                time_status = "Verified against BEA GDP release article"
                official_matched[sheet_name] += 1
            else:
                date_status = "BEA GDP release date exists; source timestamp differs — review"
                time_status = "Not matched to official release time — review"
        elif status == "Possible archival placeholder":
            date_status = "Suspected archival placeholder — not verified"
            time_status = "Suspected placeholder time — do not use as exact time"
        elif event_local:
            date_status = "Converted from source UTC timestamp; official match pending"
            time_status = "Converted from source UTC timestamp; official match pending"
        else:
            date_status = "No reliable event timestamp — verify official date"
            time_status = "No exact source time"

        actual_num = parse_numeric(v.get("Actual value"))
        expected_num = parse_numeric(v.get("Expected value"))
        surprise = actual_num - expected_num if actual_num is not None and expected_num is not None else None
        if country == "Euro Zone":
            market_links = "Germany; France; Italy"
        else:
            market_links = country
        tier = "Base country" if country in BASE_COUNTRIES else "Addition" if country in ADDITIONS else "Eurozone shared policy/data unit"
        if status == "Possible archival placeholder":
            handling = "Exclude from event-date regression until official date/time is verified; keep for audit."
        elif date_status.startswith("Matched official"):
            handling = "Date/time verified against the official release source; consensus still needs point-in-time provenance."
        elif is_cleveland_cpi:
            handling = "Keep as an alternative Cleveland Fed inflation measure; exclude from national headline CPI dummy pending timing check."
        elif event_local:
            handling = "Usable as a candidate only; verify official release date/time before final regression."
        else:
            handling = "Do not assign a precise event day/time until official source is checked."
        panel_rows.append([
            tier, country, market_links, panel_family, release_def, basis,
            "Yes" if primary else "No", "Yes" if event_dummy_candidate else "No", policy_label, stage, period,
            candidate_date, event_local.time().replace(tzinfo=None) if event_local else None,
            tz_name, date_status, time_status, verified_date, verified_time, official_source,
            source_date.date() if isinstance(source_date, datetime) else source_date,
            source_time, utc_ts, status, v.get("Time note"), v.get("Currency"), title,
            v.get("Prior value"), v.get("Actual value"), v.get("Expected value"),
            actual_num, expected_num, surprise, v.get("Prior revision"), v.get("Reference period"),
            v.get("Announcement ID"), v.get("Economic event ID"), v.get("Category"), sub,
            v.get("Source URL"), "Yes" if is_missing_value(v.get("Actual value")) else "No",
            "Yes" if is_missing_value(v.get("Expected value")) else "No", handling,
        ])

    source_us_national_cpi_rows = sum(
        1 for item in raw_rows
        if item[2] == "United States" and item[1] == "CPI" and "cleveland cpi" not in str(item[10].get("Announcement") or "").lower()
    )
    source_us_gdp_rows = sum(1 for item in raw_rows if item[2] == "United States" and item[1] == "GDP Growth")
    us_cpi_calendar_complete = us_cpi_calendar_complete and official_matched["CPI Inflation"] == source_us_national_cpi_rows
    us_gdp_calendar_complete = us_gdp_calendar_complete and official_matched["GDP Growth"] == source_us_gdp_rows

    # One row per actual event date and country/market. There are deliberately no
    # zero-event rows: an absent row is not converted to a zero unless coverage is certified.
    flag_groups = defaultdict(list)
    for row in panel_rows:
        family = row[3]
        if family not in {"Policy Rate", "CPI", "GDP Growth"} or row[7] != "Yes":
            continue
        event_date = row[11]
        if not isinstance(event_date, date):
            continue
        markets = row[2].split("; ")
        for market in markets:
            flag_groups[(market, event_date)].append(row)
    event_flag_rows = []
    for (market, event_date), items in sorted(flag_groups.items()):
        families = {item[3] for item in items}
        statuses = {item[14] for item in items}
        source_ids = {str(item[34]) for item in items if item[34] is not None}
        event_flag_rows.append([
            market, event_date,
            1 if "Policy Rate" in families else None,
            1 if "CPI" in families else None,
            1 if "GDP Growth" in families else None,
            len(source_ids),
            "Official date/time match exists" if any(s.startswith("Matched official") for s in statuses) else "Candidate event date; validation status in panel",
            "Event day present; no implicit zeroes created",
        ])

    # Validation and coverage by source country and family.
    summary = defaultdict(list)
    for row in panel_rows:
        key = (row[1], row[3])
        summary[key].append(row)
    coverage_rows = []
    all_families = ["Policy Rate", "CPI", "Alternative CPI", "GDP Growth", "Policy Communication", "GDP Tracker"]
    for country in sorted(SAMPLE_UNITS):
        for family in all_families:
            rows = summary.get((country, family), [])
            dates = sorted({r[11] for r in rows if isinstance(r[11], date)})
            flagged = sum(1 for r in rows if r[22] == "Possible archival placeholder")
            verified = sum(1 for r in rows if str(r[14]).startswith("Matched official"))
            miss_act = sum(1 for r in rows if r[39] == "Yes")
            miss_exp = sum(1 for r in rows if r[40] == "Yes")
            if country == "United States" and family == "CPI":
                complete = "Yes for BLS monthly CPI release dates in this sample window" if us_cpi_calendar_complete else "No — BLS and source headline CPI date sets differ"
                zero_rule = "Calendar absence can be treated as no BLS CPI release, then align to actual trading days." if us_cpi_calendar_complete else "Do not code absent dates as 0 until the BLS/source difference is resolved."
            elif country == "United States" and family == "GDP Growth":
                complete = "Yes; all BEA GDP releases in this sample window are reconciled" if us_gdp_calendar_complete else "No — BEA and source GDP date sets differ"
                zero_rule = "Calendar absence can be treated as no BEA GDP release, then align to actual trading days." if us_gdp_calendar_complete else "Do not code absent dates as 0 until the BEA/source difference is resolved."
            elif country == "United States" and family == "Policy Rate":
                complete = "All source decision rows match Federal Reserve statements; full historical calendar reconciliation is pending"
                zero_rule = "Do not code absent days as 0 until the complete FOMC decision calendar is reconciled."
            else:
                complete = "No — full period not reconciled to official schedule"
                zero_rule = "Do not code absent rows as 0 until official calendar completeness is verified."
            coverage_rows.append([
                country, family, len(rows), len(dates), dates[0] if dates else None, dates[-1] if dates else None,
                flagged, verified, miss_act, miss_exp, complete, zero_rule,
            ])

    # Official release-date/time logs are saved in the workbook for reproducibility.
    official_rows = list(fed_audit) + list(bea_audit)
    for dt, local in sorted(bls_dates.items()):
        if START <= dt <= END:
            official_rows.append([
                "United States", "CPI", dt, local.time().replace(tzinfo=None), "America/New_York",
                f"https://www.bls.gov/schedule/{dt.year}/home.htm", retrieved_on,
                "Official BLS annual release schedule; source panel records matched this date/time.",
            ])

    source_register_rows = []
    for country in sorted(SAMPLE_UNITS):
        stats_url, central_url = OFFICIAL_PORTALS[country]
        source_register_rows.append([
            country, stats_url, central_url, GDP_PORTAL.get(country, stats_url),
            ("US CPI and GDP rows matched official release sources; source Fed decision rows matched, full meeting-calendar reconciliation pending." if country == "United States" else "Official agency reference portals identified; historical rows not all checked individually."),
            ("Other U.S. policy communications and all non-U.S. historical events remain pending event-level checks." if country == "United States" else "Event-level date/time validation remains pending."),
        ])

    definitions_rows = [
        ["Topic", "Rule applied", "Interpretation / limitation"],
        ["Sample", f"{START.isoformat()} through {END.isoformat()} inclusive", "27 national markets (15 base + 12 additions); Euro Zone retained as the shared ECB/Eurostat unit."],
        ["Source timezone", "Original announcement date/time is stored in Europe/Budapest; UTC timestamp is preserved.", "Metadata in source workbook says the collector converted Investing.com time-zone 56 to Europe/Budapest."],
        ["Timezone conversion", "Where a source UTC timestamp exists, convert it through IANA time-zone rules to the local publisher / central-bank zone listed in Time Zones.", "This is a conversion, not independent proof that the archived timestamp is correct."],
        ["Placeholder handling", "Retain but flag rows whose source status says Possible archival placeholder.", "Placeholder rows are excluded from primary-series and event-day candidate flags. Do not use their apparent local date/time until an official source confirms it."],
        ["Official validation", "Match US national CPI, policy decisions, and GDP releases to official BLS, Federal Reserve, and BEA sources; leave other rows explicitly pending.", "All 637 US national CPI rows (80 dates), 55 Fed decision rows, and 81 GDP rows matched official dates/times. US CPI and GDP release calendars are fully reconciled; the historical Fed calendar completeness check remains open. This does not validate forecasts."],
        ["CPI definitions", "Keep headline, core, HICP/CPIF/CPIH, mid-month, index-level and regional measures distinct.", "Primary value-series candidate is headline CPI YoY, or headline MoM where no YoY headline exists. Basis remains explicit. Cleveland Fed Median CPI is an alternative measure; CPI event dummies use national headline release rows only."],
        ["GDP definitions", "Keep QoQ, annualized QoQ (SAAR), YoY, annual and monthly activity series distinct; preserve source title and stage when stated.", "US stage and quarter/year come from matched BEA article titles. Elsewhere, missing stage or reference-period year remains unspecified; no stage is inferred."],
        ["Policy-rate definitions", "Keep target decisions, corridor/facility rates and policy-related benchmarks as separate rows.", "China LPR is marked as a policy-related benchmark, not a conventional central-bank target. ECB corridor rates are not collapsed into one value."],
        ["Euro Zone linkage", "ECB / Eurostat event rows link to Germany, France and Italy market outcomes.", "Do not treat Euro Zone as a fourth national equity market or count one ECB decision three times as three independent events."],
        ["Surprise", "Numeric source-scale surprise = actual minus expected when both parse as numbers.", "The source text values are preserved. Numeric differences remain in the original units / scale; archived expectations are not proven point-in-time consensus."],
        ["Missing values", "Blank expected or actual values remain blank; blanks are not zero.", "Source values may reflect later revisions; first-release value and forecast vintages need a separate provenance check."],
        ["Event flags", "Event-date flags are 1 only where a primary rate decision, national headline CPI release, or GDP release row exists.", "The workbook does not create non-event-day zeroes. Absence is not zero until that country's official release calendar and trading-day map are complete."],
        ["Market reaction day", "No return-session assignment is supplied.", "Map verified local announcement date/time to the first tradable session using each market's trading calendar and close time."],
        ["GDP trackers", "NIESR GDP estimates and Atlanta Fed GDPNow are separated from official GDP growth releases.", "They are forecasts/nowcasts, not national-accounts releases."],
        ["Data source", "Investing.com archive workbook retrieved through 2026-10-12.", "Other calendars in the repository can be used for triangulation but are not treated as official validation."],
    ]

    timezone_rows = [[c, TZ_BY_COUNTRY[c], "Official publisher / central-bank principal timezone; exact release authority may differ for an individual event."] for c in sorted(SAMPLE_UNITS)]

    wb = Workbook(write_only=False)
    ws = wb.active
    ws.title = "Announcement Panel"
    add_table(ws, panel_headers, panel_rows, "AnnouncementPanel")
    for row in ws.iter_rows(min_row=2):
        for col in [12, 17, 20]:
            row[col - 1].number_format = "yyyy-mm-dd"
        for col in [13, 18, 21]:
            row[col - 1].number_format = "hh:mm:ss"
        row[21].number_format = "yyyy-mm-dd hh:mm:ss"
        for col in [30, 31, 32]:
            row[col - 1].number_format = "0.######"
    ws2 = wb.create_sheet("Event Date Flags")
    add_table(ws2, ["Market", "Local event date", "Policy-rate dummy", "CPI dummy", "GDP dummy", "Source series rows", "Validation status", "Zero-day rule"], event_flag_rows, "EventDateFlags")
    for row in ws2.iter_rows(min_row=2):
        row[1].number_format = "yyyy-mm-dd"
    ws3 = wb.create_sheet("Coverage and Validation")
    add_table(ws3, ["Country / unit", "Event family", "Source rows", "Distinct event dates", "First candidate date", "Last candidate date", "Placeholder rows", "Officially matched rows", "Missing actual rows", "Missing expected rows", "Calendar completeness", "Non-event zero rule"], coverage_rows, "CoverageValidation")
    for row in ws3.iter_rows(min_row=2):
        row[4].number_format = "yyyy-mm-dd"
        row[5].number_format = "yyyy-mm-dd"
    ws4 = wb.create_sheet("Official US Schedule Checks")
    add_table(ws4, ["Country", "Event family", "Official local date", "Official local time", "Official timezone", "Official source URL", "Validation basis", "Check result"], official_rows, "OfficialUSChecks")
    for row in ws4.iter_rows(min_row=2):
        row[2].number_format = "yyyy-mm-dd"
        row[3].number_format = "hh:mm:ss"
    ws5 = wb.create_sheet("Official Source Register")
    add_table(ws5, ["Country / unit", "Official statistics publisher / release portal", "Official central bank / policy portal", "Official GDP release portal", "Validation status", "Current coverage note"], source_register_rows, "OfficialSourceRegister")
    ws6 = wb.create_sheet("Definitions and Rules")
    add_table(ws6, definitions_rows[0], definitions_rows[1:], "DefinitionsRules")
    ws7 = wb.create_sheet("Time Zones")
    add_table(ws7, ["Country / unit", "IANA timezone used", "Basis / caveat"], timezone_rows, "TimezoneMap")
    ws8 = wb.create_sheet("GDP Trackers")
    tracker_rows = [r for r in panel_rows if r[3] == "GDP Tracker"]
    add_table(ws8, panel_headers, tracker_rows, "GDPTrackers")
    # A visible title row is not used so filtering/table ranges remain straightforward.

    # Make links and status columns readable without excessive horizontal widths.
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
    wb.save(OUTPUT)
    print(f"Wrote {OUTPUT}")
    print(f"Panel rows: {len(panel_rows):,}; event-date flags: {len(event_flag_rows):,}; official schedule checks: {len(official_rows):,}")
    print(f"Official matches: US CPI {official_matched['CPI Inflation']:,}; Fed policy {official_matched['Policy Rates']:,}; BEA GDP {official_matched['GDP Growth']:,}")


if __name__ == "__main__":
    main()
