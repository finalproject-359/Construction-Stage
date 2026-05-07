# Dashboard Data Troubleshooting (Possible Causes)

This checklist captures likely reasons the dashboard may fail to present complete or accurate data.

## Data source and network
- Invalid or unsupported data-source URL format can abort remote fetch.
- Apps Script / Google Sheet endpoint may return non-OK HTTP status.
- Fetch timeout (8s) can terminate slow responses.
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
- Dashboard cache TTL and polling timing can make UI appear stale between refresh cycles.
- In-flight request guard can delay visible updates when concurrent refresh triggers occur.
- Cache/localStorage corruption or manual edits can cause inconsistent displays.

## Cross-page/local-storage dependency
- Some pages rely on local storage already being hydrated by other pages.
- If users land directly on a page without prior sync, required local data can be missing.
- Daily costs are local-only in cost management; cross-device/browser data may not appear.

## Merge/conflict behavior
- Remote-first merge precedence can hide newer unsynced local edits.
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
