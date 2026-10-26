# Public source archives

Archived in this repository on 26 October 2026. The new Federal Reserve, ECB and Forex Factory GitHub files were retrieved on that date; the Kaggle, Hugging Face and IMF payloads were retained from the 19 October source audit. These are source files, not a harmonized panel. Keep source names, raw values and source dates intact when creating analysis-ready data.

The main workbook, [`../macroeconomic_announcements_2000_to_2026-10-12.xlsx`](../macroeconomic_announcements_2000_to_2026-10-12.xlsx), was already in the repository and ends on 12 October 2026; it was not regenerated in this update.

## Cleaned analysis panel

[`../cleaned_announcement_panel_2019-11-20_to_2026-08-31.xlsx`](../cleaned_announcement_panel_2019-11-20_to_2026-08-31.xlsx) is a derived workbook for the 27-country sample (15 base countries plus 12 additions), with the Euro Zone retained as a shared ECB/Eurostat event unit. It covers candidate local event dates from 20 November 2019 through 31 August 2026 and includes rate decisions, CPI, GDP growth, policy communications, and a separate GDP tracker sheet.

The workbook converts the source UTC timestamps into the country or publisher's IANA timezone, preserves the original Europe/Budapest date/time and raw values, separates CPI and GDP definitions, calculates source-scale actual-minus-expected values where numeric, and flags suspected archival placeholders. Its `Event Date Flags` sheet contains only observed event dates; it does not create zeroes for days without a row.

Official date/time checks are complete for US national CPI (637 source rows on 80 release dates matched to BLS), US Federal Reserve rate decisions (55 source rows matched to official statements), and US GDP (81 releases matched to BEA articles). The full historical Federal Reserve calendar completeness check remains open. Other sample countries have official publisher portals recorded, but their historical rows are not yet individually reconciled. Cleveland Fed Median CPI is separated from national CPI and its exact release time remains unverified. Archived expectation fields are not proven point-in-time consensus values.

Rebuild with `python scripts/build_clean_announcement_panel.py` from the repository root. The script retrieves current BLS, Federal Reserve, and BEA official source pages to recreate the U.S. checks; the resulting snapshot records the matched dates and article links in the workbook.

## Announcement calendars

| Local path | Source and snapshot | Use and known limits |
|---|---|---|
| [`forex_factory/spoluan/datasets/`](forex_factory/spoluan/datasets/) | [spoluan/forex-factory-scraper](https://github.com/spoluan/forex-factory-scraper), commit [`127e424`](forex_factory/spoluan/source_commit.txt); annual CSVs 2010–2023 | Actual, forecast and previous values for nine currencies. The repository's MIT license covers its software; it does not establish rights to the underlying Forex Factory calendar content. The source timezone is undocumented, and the collector has a year-boundary date error. Treat dates as candidates for validation. |
| [`forex_factory/huggingface/forex_factory_cache_2025-04-07.csv`](forex_factory/huggingface/forex_factory_cache_2025-04-07.csv) | [Ehsanrs2 Forex Factory Calendar](https://huggingface.co/datasets/Ehsanrs2/Forex_Factory_Calendar), fixed revision `5cf3d34732b7995b3bc455139495257e64b2e1b4` | 83,427 rows dated 2007-01-01 to 2025-04-07, with timestamp, currency, event, actual, forecast and previous. The dataset card declares MIT; the underlying calendar originates from Forex Factory. It has no country field, and timestamp offsets do not guarantee correct event-day assignment. |
| [`kaggle/devorvant/economic-calendar-v11.zip`](kaggle/devorvant/economic-calendar-v11.zip) | [devorvant/economic-calendar](https://www.kaggle.com/datasets/devorvant/economic-calendar), version 11; snapshot dated 2021-04-26 | 121,658 rows, 20 countries/regions, dates 2011-01-01 to 2021-04-26. The uploader identifies Investing.com as the source and declares Open Database License / Database Contents (see [ODbL 1.0](https://opendatacommons.org/licenses/odbl/1-0/)). This overlaps the Investing.com workbook and is not independent validation. The archive contains suspicious first-of-month historical dates. |
| [`kaggle/global/global-economic-calendar-v1.zip`](kaggle/global/global-economic-calendar-v1.zip) | [youneseloiarm/global-economic-calendar](https://www.kaggle.com/datasets/youneseloiarm/global-economic-calendar), version 1 | 663,497 raw rows dated 2020-01-01 to 2025-10-01 across 98 country/region labels. 547,221 rows are exact duplicates in the downloaded CSV. The uploader declares [CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/); the card says times may be local or UTC. |

The Kaggle ZIP files are retained as distributed. Extract them with a standard ZIP utility. The Devovant archive contains CSV and Excel copies; the Global archive contains `Calender_data.csv`.

## Policy-event and shock datasets

| Local path | Source | Use and known limits |
|---|---|---|
| [`policy/usmpd/USMPD.xlsx`](policy/usmpd/USMPD.xlsx) and [`policy/usmpd/monetary-policy-surprises.zip`](policy/usmpd/monetary-policy-surprises.zip) | [Federal Reserve Bank of San Francisco, U.S. Monetary Policy Event-Study Database](https://www.frbsf.org/research-and-insights/data-and-indicators/us-monetary-policy-event-study-database/) | US event timestamps in Eastern Time and financial-market changes/shocks. These are market-based surprises, not actual minus survey consensus or announced/prior policy-rate levels. Cite Acosta et al. (2025), as requested by the source. |
| [`policy/ea_mpd/Dataset_EA-MPD.xlsx`](policy/ea_mpd/Dataset_EA-MPD.xlsx) | [ECB Euro Area Monetary Policy Event-Study Database](https://www.ecb.europa.eu/pub/pdf/annex/Dataset_EA-MPD.xlsx) | Euro-area policy-event market responses. Use one shared ECB event for Germany, France and Italy; this is not three national policy announcements. Cite Altavilla et al. (2019), as requested in the workbook. |
| [`policy/imf/`](policy/imf/) | [Bolhuis, Das and Yao (2024), IMF Working Paper 24/224](https://sites.google.com/site/mabolhuis/data) | Public author files for daily monetary-policy shocks. The authors permit academic use without charge when cited. The inspected files stop in December 2022 despite newer file names; the older version preserves original meeting dates useful for checking market-date alignment. |

## Reuse and citation notes

The files are mirrored here for this thesis project with source attribution. A source's public download link or a code-repository license does not necessarily grant a license to the underlying third-party calendar data. In particular, the Forex Factory mirrors and the Investing.com-derived archive carry upstream-provenance limits; the Global Calendar uploader's license also does not independently verify every upstream right. Check the cited source terms before redistributing these raw files outside this thesis repository. Keep the files' source and version metadata with any derived data.

The IMF files are provided for academic use with citation under the author page's stated terms. For the USMPD and EA-MPD, retain the citations and source links above; no separate open-data license was identified in the source files reviewed for this project.

## Reproducibility and interpretation

- These sources are intentionally kept separate. Do not concatenate overlapping records and count them as independent observations.
- A row is not always a unique release: headline/core CPI and advance/final GDP can share a date. Record the event definition and reference period; count unique release dates within each indicator.
- For a daily market study, preserve both the published announcement date and the first tradable market session. Time-zone conversion and after-close/weekend releases can shift the reaction date.
- Compute a calendar surprise as actual minus a genuinely pre-release consensus. Actual minus previous is a change, not a surprise. The archived calendars do not provide a complete point-in-time history of each forecast.
- Validate event dates and first-release values against official sources where possible. For the US, use the [BLS CPI release schedule](https://www.bls.gov/schedule/news_release/cpi.htm) and [BEA release schedule](https://www.bea.gov/news/schedule). ALFRED, OECD revisions and BIS policy-rate data are useful supplementary controls/value checks, not replacements for an announcement calendar; their source pages are [ALFRED](https://alfred.stlouisfed.org/help), [OECD](https://data-explorer.oecd.org/s/da) and [BIS CBPOL](https://data.bis.org/topics/CBPOL).

## Checksums

File versions, retrieval dates and SHA-256 values are recorded in [`manifest.json`](manifest.json) and [`checksums.sha256`](checksums.sha256). The annual Forex Factory archive commit is recorded in [`forex_factory/spoluan/source_commit.txt`](forex_factory/spoluan/source_commit.txt).
