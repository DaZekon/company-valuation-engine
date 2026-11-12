# Release status

Local build verified November 12, 2026.

- The complete 32-test pytest suite passed using the owner's installed Python 3.15.2 environment.
- `run_demo.py` generated an Excel workbook and Bear/Base/Bull valuation summaries using that same environment.
- The bundled example's saved spreadsheet results passed the independent Python cache verifier.
- The earlier compatible-spreadsheet-engine validation contains 44 comparisons/input-change checks; see `validation.md` and `spreadsheet_checks.json`.
- A GitHub Actions workflow is included, but no hosted Actions run is claimed before repository publication.
- Native Excel opening/recalculation was attempted, but no completed native recalculation is claimed. New workbooks still require an Excel/LibreOffice open, recalculation and save.
- The example is fictional. A reviewed real-company case and live end-to-end SEC download are not included in this release.
- Resume bullets, a LinkedIn Projects entry and a LinkedIn post draft are in `portfolio-kit.md`. Nothing has been posted to LinkedIn.

The project is prepared locally. Remote GitHub publication requires explicit approval of public visibility and successful repository creation/upload.
