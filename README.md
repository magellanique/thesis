# Thesis data

This repository contains the macroeconomic announcement workbook and source archives for the daily announcement study.

- [Macroeconomic announcements, 2000–2026-10-12](data/macroeconomic_announcements_2000_to_2026-10-12.xlsx)
- [Cleaned announcement panel, 2019-11-20–2026-08-31](data/cleaned_announcement_panel_2019-11-20_to_2026-08-31.xlsx)
- [Source archive inventory, download dates, citations, licenses and data-quality notes](data/sources/README.md)

The cleaned panel contains event-level records, local-time conversions, release definitions, surprise fields, event-date flags, and validation status. US CPI, Federal Reserve policy-decision, and BEA GDP rows are matched to official sources; the other countries remain explicitly flagged for official date/time checks. It does not assign zeroes to unverified no-event days.

The source files are preserved separately. They use different event definitions, date conventions, time zones and revision practices, so review the workbook's validation and definition sheets before using the records in an analysis.
