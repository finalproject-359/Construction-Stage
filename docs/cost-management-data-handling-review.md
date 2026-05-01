# Cost Management Data Handling Review

## Scope
- `cost-management.js`
- Related storage and data-source helpers from `activities.js`, `projects.js`, and `data-service.js` for context.

## Current Data Flow
1. **Projects source**
   - Cost Management reads projects from browser local storage key `constructionStageProjects`.
   - No direct remote fetch is performed in this page for projects.
2. **Activities source**
   - Primary local source: `constructionStageActivities`.
   - Legacy fallback source: `constructionStageCostActivities`.
   - Optional remote source: `window.DataBridge.fetchRowsFromSource()`.
3. **Daily actual costs source**
   - Stored only in local storage key `constructionStageDailyCosts`.
   - No remote persistence exists for this dataset.
4. **Presentation**
   - On page load, local and remote activities are merged and deduplicated by `projectId::activityId`.
   - Project selection renders KPIs and costing records from merged activity data + local daily cost entries.

## Findings

### 1) Daily costs are local-only and can be lost across devices/browsers
- `saveDailyCosts` writes exclusively to local storage.
- There is no API call to persist daily costs remotely and no sync path in `cost-management.js`.
- Impact: users on a second browser/device will not see previously entered daily actual costs.

### 2) Project list depends on prior pages to hydrate local storage
- `loadProjects()` reads only local storage and does not fetch projects when empty/stale.
- If a user lands directly on Cost Management before visiting pages that fetch projects, the project list can be empty.
- Impact: false "no projects" state despite existing remote projects.

### 3) Remote activity merge prioritizes remote data over local edits
- Merge order is `[...remoteActivities, ...localActivities]` with first-entry-wins dedupe.
- Impact: if local activity data is newer than remote (e.g., unsynced local change), local updates are hidden.
- This may be intended, but it is not documented in-code as a conflict strategy.

### 4) Daily cost entries are not validated against activity date range in submit handler
- Date inputs include `min` and `max` attributes in the form, which helps normal UI usage.
- However, submit handler does not re-validate date bounds before saving.
- Impact: malformed/manual submissions could insert out-of-range entries.

### 5) Weak integrity checks for orphaned daily costs
- Entries are keyed by `{ projectId, activityId, date }` and upserted by exact string matching.
- No cleanup path exists for stale entries if activities are renamed/re-keyed/removed.
- Impact: local storage may accumulate orphaned rows that never render.

## Practical Recommendations (prioritized)
1. Add remote persistence + fetch for `constructionStageDailyCosts` via existing Apps Script payload pattern used in `activities.js` and `projects.js`.
2. Add project bootstrap fallback in Cost Management: if no local projects, fetch projects from source and cache locally.
3. Define and document merge precedence explicitly (remote-first vs local-first) and align with expected UX.
4. Add submit-time guard: reject daily-cost dates outside `[activity.startDate, activity.finishDate]`.
5. Add periodic cleanup for orphaned daily costs not matching current activity set.

## Notes
- Current parsing/normalization utilities are robust against alias/format variance (`getValueByAliases`, numeric parsing, date normalization).
- Rendering path correctly escapes user-visible strings before HTML injection in key views.
