# Cost Management Data Handling Review

## Scope
- `cost-management.js`
- `data-service.js`
- `code.gs`
- Related storage helpers from `activities.js` and `projects.js` for context.

## Current Data Flow
1. **Projects source**
   - Cost Management fetches projects directly from the Apps Script `projects` resource when the page boots.
   - If the direct resource request fails, the page attempts the bundled `all` payload exposed by `DataBridge.fetchDashboardBundleFromSource()`.
   - Browser local storage (`constructionStageProjects`) is used only as a warm-start fallback when Google Sheets cannot be reached. An authoritative empty remote project list clears stale cached projects.
2. **Activities source**
   - Primary source: Apps Script `activities` resource or the dashboard row source exposed by `DataBridge`.
   - Local activity/cost metadata keys are retained only for user-maintained Cost ID and planned-cost overrides that still match authoritative activity rows.
3. **Daily actual costs source**
   - Primary source: Apps Script `daily_costs` resource backed by the `DailyCosts` sheet.
   - Creates, updates, and deletes are posted to Apps Script before the UI refreshes local cache state.
   - Browser local storage (`constructionStageDailyCosts`) is a cache for warm starts and offline/error fallback, not the source of truth.
4. **Server-side write integrity**
   - Apps Script serializes project, activity, cost, and daily-cost mutations with a document lock to avoid overlapping writes.
   - Successful mutations flush the spreadsheet and advance realtime metadata so clients bypass old read-cache keys.
5. **Presentation**
   - On page load, remote projects, activities, costs, and daily-cost rows are normalized and cached.
   - Project selection renders KPIs and costing records from authoritative remote rows plus retained matching local metadata.

## Reliability Improvements Already Implemented

### 1) Daily costs are no longer local-only
- Daily cost create/update/delete flows call the Apps Script `daily_costs` resource.
- After each mutation, Cost Management reloads daily costs from Google Sheets for the active project.
- If Google Sheets is unavailable, the page warns the user that cached daily costs are being shown.

### 2) Project lists no longer require prior page hydration
- Direct Cost Management entry now loads projects from the Apps Script `projects` endpoint.
- If that endpoint fails, the bundled `all` payload can still supply projects.
- Cached projects are used only when the remote sheet cannot be reached, not when the sheet returns an authoritative empty list.

### 3) Remote empty states clear stale cache
- An empty remote project response is treated as authoritative.
- An empty remote daily-cost response for the selected project replaces that project slice locally, so deleted sheet rows do not reappear from browser cache.

### 4) Concurrent writes are serialized
- Apps Script wraps mutations in a document lock.
- This reduces lost updates and duplicate row creation when multiple users save costs or daily costs simultaneously.

### 5) Submit-time validation is enforced
- Daily-cost submissions validate project ID, Cost ID, date, working day, activity start/finish bounds, progress, and positive actual cost before posting.
- Apps Script also enforces project, activity, and cost foreign-key existence before saving daily-cost rows.

## Remaining Watch Items

### 1) Offline changes are intentionally not queued
- The current behavior favors reliability over offline editing: if a write cannot reach Google Sheets, the UI does not pretend the save succeeded.
- If offline-first behavior becomes a requirement, add a visible pending-sync queue with retry status instead of silently storing mutations locally.

### 2) Local cost metadata still exists as an override cache
- Cost IDs and planned-cost edits are persisted remotely, but the page keeps local matching overrides to avoid losing in-progress context while remote rows refresh.
- Continue pruning overrides against authoritative activity rows to prevent stale metadata from resurfacing.

### 3) Apps Script deployment remains a critical dependency
- The web app must be deployed with access settings that allow the browser to read and mutate the sheet.
- Deployment URL changes still require updating `DEFAULT_DATA_SOURCE_URL` in `data-service.js`.
