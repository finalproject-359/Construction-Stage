# Dashboard Data Troubleshooting (Possible Causes)

This checklist captures likely reasons the dashboard may fail to present complete or accurate data.

## Data source and network
- Invalid or unsupported data-source URL format can abort remote fetch.
- Apps Script / Google Sheet endpoint may return non-OK HTTP status.
- Fetch timeout (20s) can terminate slow responses.
- CORS or browser/network restrictions can block remote fetches.
- Remote payload may not match expected shape (`rows`, `data`, `items`, or array).

## Spreadsheet/schema mismatches
- Header detection may lock onto wrong row if aliases are weak/missing.
- Column names outside recognized aliases can produce missing fields.
- Summary rows (e.g., total/grand total) are intentionally excluded.
- Activity rows with both missing ID and missing activity name are dropped.

## Data filtering and normalization side effects
- Rows are filtered out when planned/actual/EV/progress all evaluate to zero.
- Date normalization can blank invalid/unparsable date strings.
- Numeric parsing strips non-numeric characters; malformed values can collapse to 0.
- Percent normalization treats values <=1 as fractions and >1 as percent numbers.

## Local cache and freshness
- Dashboard warm-start cache can briefly show the last stable rows while the live source is verified immediately; a successful empty live response clears the dashboard so deleted Google Sheet rows do not remain visible.
- Visible dashboard sessions poll the live source every 5 seconds and also refresh on focus, reconnect, page show, visibility changes, and local cost-data storage updates.
- In-flight request guard can delay visible updates when concurrent load triggers occur.
- Cache/localStorage corruption or manual edits can cause inconsistent fallback displays when Google Sheets is unavailable.

## Cross-page/local-storage dependency
- Cost Management now fetches projects and daily costs directly from Apps Script on direct page entry.
- Browser local storage is a warm-start/error fallback cache; if Google Sheets cannot be reached, cached project or daily-cost rows may still appear with a warning.
- Daily costs are remotely persisted through the Apps Script `daily_costs` resource, so cross-device/browser data should appear after sync.

## Merge/conflict behavior
- Remote-first merge precedence can hide newer unsynced local metadata, although Cost Management retains matching local Cost ID/planned-cost overrides while the sheet refreshes.
- Legacy fallback keys can introduce old or duplicate-looking rows if cleanup is absent.
- Orphaned entries (activity/project renamed/removed) may remain in storage.

## Runtime dependency and rendering
- Missing Chart.js disables graphs, reducing perceived dashboard completeness.
- Missing DOM targets or changed element IDs/classes can block render updates.
- Browser extensions can inject noisy runtime errors (some are ignored, others may still interfere).

## User input/data quality issues
- Manual/invalid form submissions can persist out-of-range or invalid values if not revalidated.
- Inconsistent project/activity IDs across sources break joins and aggregates.
- Duplicate IDs across projects can create unexpected dedupe outcomes.
