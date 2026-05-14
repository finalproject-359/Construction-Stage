(function attachDataBridge(global) {
  const DEFAULT_DATA_SOURCE_URL =
    "https://script.google.com/macros/s/AKfycbw-iOWnshHXBXROcFhI3emMKTXh7bAFrhVPyYkGHfg_MShakUWwYtCP86HUyLWBzL6a/exec";

  const EXPECTED_HEADER_ALIASES = [
    "project id",
    "project name",
    "activity id",
    "project id",
    "project",
    "activity",
    "planned start",
    "planned finish",
    "complete",
    "created at",
  ];

  const normalizeHeader = (value) =>
    String(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  const extractRowsFromPayload = (payload) => {
    if (Array.isArray(payload)) return payload;
    if (!payload || typeof payload !== "object") return [];
    if (Array.isArray(payload.rows)) return payload.rows;
    if (Array.isArray(payload.data)) return payload.data;
    if (Array.isArray(payload.items)) return payload.items;
    if (Array.isArray(payload.activities)) return payload.activities;
    if (Array.isArray(payload.costs)) return payload.costs;
    if (Array.isArray(payload.dailyCosts)) return payload.dailyCosts;
    if (Array.isArray(payload.daily_costs)) return payload.daily_costs;
    if (payload.dashboard && typeof payload.dashboard === "object") {
      const dashboardRows = extractRowsFromPayload(payload.dashboard);
      if (dashboardRows.length) return dashboardRows;
    }
    return [];
  };

  const extractResourceRows = (payload, resourceName) => {
    const resourcePayload = payload?.[resourceName];
    if (Array.isArray(resourcePayload)) return resourcePayload;
    if (resourcePayload && typeof resourcePayload === "object") {
      const directRows = extractRowsFromPayload(resourcePayload);
      if (directRows.length) return directRows;
      if (Array.isArray(resourcePayload[resourceName])) return resourcePayload[resourceName];
    }

    const directValue = payload?.[resourceName];
    if (Array.isArray(directValue)) return directValue;
    return [];
  };

  const findHeaderRowIndex = (sheet) => {
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false,
    });

    let bestIndex = 0;
    let bestScore = 0;

    rows.forEach((row, index) => {
      const normalizedCells = row.map((cell) => normalizeHeader(cell)).filter(Boolean);
      const score = EXPECTED_HEADER_ALIASES.reduce((count, alias) => {
        const hasAlias = normalizedCells.some((cell) => cell.includes(alias));
        return count + (hasAlias ? 1 : 0);
      }, 0);

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    return bestIndex;
  };

  const toGoogleSheetCsvUrl = (inputUrl) => {
    if (!inputUrl) return "";
    const trimmed = inputUrl.trim();
    if (!trimmed) return "";
    if (/output=csv/i.test(trimmed)) return trimmed;

    try {
      const parsed = new URL(trimmed);
      const match = parsed.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) return "";
      const sheetId = match[1];
      const gid = parsed.searchParams.get("gid") || "0";
      return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
    } catch {
      return "";
    }
  };

  const isAppsScriptWebAppUrl = (inputUrl) => {
    if (!inputUrl) return false;
    try {
      const parsed = new URL(inputUrl.trim());
      return parsed.hostname === "script.google.com" && /\/macros\/s\/.+\/exec$/.test(parsed.pathname);
    } catch {
      return false;
    }
  };

  const LIVE_SOURCE_FETCH_TIMEOUT_MS = 20000;
  const LIVE_SOURCE_RETRY_DELAYS_MS = [0, 750, 1500];
  const DEFAULT_REALTIME_POLL_MS = 2000;
  const HIDDEN_REALTIME_POLL_MS = 10000;
  const REALTIME_CHANNEL_NAME = "construction-stage-google-sheet-sync";
  const REALTIME_STORAGE_KEY = "constructionStageGoogleSheetVersion";

  const createLiveSourceTimeoutError = () =>
    new Error(`Live data source took longer than ${Math.round(LIVE_SOURCE_FETCH_TIMEOUT_MS / 1000)} seconds to respond.`);

  const isAbortError = (error) =>
    error?.name === "AbortError" || /aborted|abort/i.test(String(error?.message || error || ""));

  const fetchRowsFromSource = async (providedUrl = "") => {
    const rawUrl = providedUrl || DEFAULT_DATA_SOURCE_URL;
    const trimmedUrl = rawUrl.trim();
    const isWebAppSource = isAppsScriptWebAppUrl(trimmedUrl);
    const csvUrl = isWebAppSource ? "" : toGoogleSheetCsvUrl(trimmedUrl);

    if (!isWebAppSource && !csvUrl) {
      throw new Error("Invalid URL. Use a valid Google Sheet or Apps Script Web App URL.");
    }

    const fetchWithTimeout = (urlValue) => fetchLiveSource(urlValue);

    if (isWebAppSource) {
      const response = await fetchWithTimeout(trimmedUrl);
      if (!response.ok) throw new Error(`Unable to fetch Apps Script Web App (HTTP ${response.status})`);

      const contentType = response.headers.get("content-type") || "";
      if (contentType.toLowerCase().includes("application/json")) {
        const payload = await response.json();
        if (payload?.error) throw new Error(payload.error);
        return {
          rows: extractRowsFromPayload(payload),
          sourceName: `Apps Script Web App (${payload?.sheetName || payload?.resource || "sheet"})`,
          generatedAt: payload?.generatedAt || "",
          version: payload?.version || payload?.realtime?.version || "",
        };
      }

      const rawText = await response.text();
      const workbook = XLSX.read(rawText, { type: "string" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const headerRowIndex = findHeaderRowIndex(sheet);
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "", range: headerRowIndex });
      return {
        rows,
        sourceName: `Apps Script Web App CSV "${firstSheetName}"`,
        generatedAt: "",
      };
    }

    const response = await fetchWithTimeout(csvUrl);
    if (!response.ok) throw new Error(`Unable to fetch Google Sheet (HTTP ${response.status})`);

    const csvText = await response.text();
    const workbook = XLSX.read(csvText, { type: "string" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const headerRowIndex = findHeaderRowIndex(sheet);
    const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "", range: headerRowIndex });

    return {
      rows,
      sourceName: `Google Sheet "${firstSheetName}"`,
      generatedAt: "",
    };
  };

  const appendNoCacheParam = (urlValue) => {
    const parsed = new URL(urlValue, window.location.href);
    parsed.searchParams.set("_ts", Date.now().toString());
    parsed.searchParams.set("_nonce", Math.random().toString(36).slice(2));
    return parsed.toString();
  };

  const wait = (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs));

  const fetchLiveSourceOnce = async (urlValue) => {
    const controller = new AbortController();
    let timedOut = false;
    const timeout = setTimeout(() => {
      timedOut = true;
      controller.abort(createLiveSourceTimeoutError());
    }, LIVE_SOURCE_FETCH_TIMEOUT_MS);
    try {
      return await fetch(appendNoCacheParam(urlValue), {
        cache: "no-store",
        // Keep this as a CORS-safelisted request. Apps Script Web Apps do not
        // handle OPTIONS preflight requests, so custom request headers like
        // Cache-Control or Pragma cause the browser to block dashboard loads
        // before doGet can return the JSON payload. Cache busting is handled
        // by appendNoCacheParam instead.
        headers: {
          Accept: "application/json, text/csv, text/plain;q=0.9, */*;q=0.8",
        },
        signal: controller.signal,
      });
    } catch (error) {
      if (timedOut || isAbortError(error)) {
        throw createLiveSourceTimeoutError();
      }
      throw error;
    } finally {
      clearTimeout(timeout);
    }
  };

  const fetchLiveSource = async (urlValue) => {
    let lastError;
    for (const delay of LIVE_SOURCE_RETRY_DELAYS_MS) {
      if (delay) await wait(delay);
      try {
        const response = await fetchLiveSourceOnce(urlValue);
        if (response.ok || response.status < 500) return response;
        lastError = new Error(`Live data source returned HTTP ${response.status}.`);
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError || new Error("Unable to reach live data source.");
  };

  const buildAppsScriptUrl = (baseUrl, params = {}) => {
    const parsed = new URL(baseUrl);
    Object.entries(params).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== "") {
        parsed.searchParams.set(key, String(value));
      }
    });
    return parsed.toString();
  };

  const normalizeSourceVersion = (payload = {}) => {
    const version = String(payload.version || payload.realtime?.version || payload.generatedAt || "").trim();
    return {
      ok: payload.ok !== false,
      version,
      generatedAt: payload.generatedAt || payload.realtime?.updatedAt || "",
      resource: payload.resource || "version",
    };
  };

  const fetchSourceVersion = async (providedUrl = "") => {
    const rawUrl = providedUrl || DEFAULT_DATA_SOURCE_URL;
    const trimmedUrl = rawUrl.trim();
    if (!isAppsScriptWebAppUrl(trimmedUrl)) return null;

    const response = await fetchLiveSource(buildAppsScriptUrl(trimmedUrl, { resource: "version" }));
    if (!response.ok) throw new Error(`Unable to fetch Google Sheet version (HTTP ${response.status})`);
    const payload = await response.json();
    if (payload?.error) throw new Error(payload.error);
    return normalizeSourceVersion(payload);
  };

  const createRealtimeSyncController = () => {
    let timer = null;
    let inFlight = false;
    let lastVersion = "";
    let started = false;
    let channel = null;

    const broadcastChange = (detail) => {
      global.dispatchEvent(new CustomEvent("google-sheet:changed", { detail }));
      try {
        global.localStorage?.setItem(REALTIME_STORAGE_KEY, JSON.stringify(detail));
      } catch (error) {
        console.warn("Unable to store Google Sheet realtime version:", error);
      }
      try {
        channel?.postMessage(detail);
      } catch (error) {
        console.warn("Unable to broadcast Google Sheet realtime version:", error);
      }
    };

    const rememberVersion = (versionInfo, { emit = true } = {}) => {
      if (!versionInfo?.version) return false;
      const previousVersion = lastVersion;
      lastVersion = versionInfo.version;
      const changed = Boolean(previousVersion && previousVersion !== versionInfo.version);
      if (changed && emit) broadcastChange(versionInfo);
      return changed;
    };

    const schedule = (delayMs) => {
      if (!started) return;
      clearTimeout(timer);
      timer = setTimeout(poll, delayMs);
    };

    const getNextDelay = () => (document.visibilityState === "hidden" ? HIDDEN_REALTIME_POLL_MS : DEFAULT_REALTIME_POLL_MS);

    const poll = async ({ emit = true } = {}) => {
      if (inFlight) return;
      inFlight = true;
      try {
        const versionInfo = await fetchSourceVersion(DEFAULT_DATA_SOURCE_URL);
        rememberVersion(versionInfo, { emit });
      } catch (error) {
        console.warn("Google Sheet realtime version check failed:", error);
      } finally {
        inFlight = false;
        schedule(getNextDelay());
      }
    };

    const handleRemoteMessage = (event) => {
      const detail = event?.data || event?.detail || null;
      if (!detail?.version || detail.version === lastVersion) return;
      lastVersion = detail.version;
      global.dispatchEvent(new CustomEvent("google-sheet:changed", { detail }));
    };

    return {
      start() {
        if (started || !isAppsScriptWebAppUrl(DEFAULT_DATA_SOURCE_URL)) return;
        started = true;
        try {
          channel = "BroadcastChannel" in global ? new BroadcastChannel(REALTIME_CHANNEL_NAME) : null;
          channel?.addEventListener("message", handleRemoteMessage);
        } catch (error) {
          console.warn("Google Sheet realtime BroadcastChannel setup failed:", error);
        }
        global.addEventListener("storage", (event) => {
          if (event.key !== REALTIME_STORAGE_KEY || !event.newValue) return;
          try {
            handleRemoteMessage({ data: JSON.parse(event.newValue) });
          } catch {
            // Ignore malformed cross-tab payloads.
          }
        });
        document.addEventListener("visibilitychange", () => {
          if (document.visibilityState === "visible") poll({ emit: true });
        });
        global.addEventListener("online", () => poll({ emit: true }));
        global.addEventListener("focus", () => poll({ emit: true }));
        poll({ emit: false });
      },
      stop() {
        started = false;
        clearTimeout(timer);
        try {
          channel?.close();
        } catch {
          // Ignore channel close failures.
        }
        channel = null;
      },
      pollNow() {
        return poll({ emit: true });
      },
      getLastVersion() {
        return lastVersion;
      },
    };
  };

  const realtimeSyncController = createRealtimeSyncController();

  global.DataBridge = {
    DEFAULT_DATA_SOURCE_URL,
    DEFAULT_REALTIME_POLL_MS,
    fetchRowsFromSource,
    fetchSourceVersion,
    startRealtimeSync: realtimeSyncController.start,
    stopRealtimeSync: realtimeSyncController.stop,
    pollRealtimeSync: realtimeSyncController.pollNow,
    fetchDashboardBundleFromSource: async (providedUrl = "") => {
      const rawUrl = providedUrl || DEFAULT_DATA_SOURCE_URL;
      const trimmedUrl = rawUrl.trim();
      if (!isAppsScriptWebAppUrl(trimmedUrl)) {
        throw new Error("Dashboard bundle fetch requires an Apps Script Web App URL.");
      }

      const withParams = (() => {
        const url = new URL(trimmedUrl);
        url.searchParams.set("resource", "all");
        url.searchParams.set("_ts", Date.now().toString());
        return url.toString();
      })();

      const response = await fetchLiveSource(withParams);
      if (!response.ok) throw new Error(`Unable to fetch Apps Script Web App (HTTP ${response.status})`);
      const payload = await response.json();
      if (payload?.error) throw new Error(payload.error);

      return {
        activities: extractResourceRows(payload, "activities"),
        costs: extractResourceRows(payload, "costs"),
        dailyCosts: extractResourceRows(payload, "daily_costs").length
          ? extractResourceRows(payload, "daily_costs")
          : extractResourceRows(payload, "dailyCosts"),
        dashboardRows: extractResourceRows(payload, "dashboard"),
        generatedAt: payload?.generatedAt || "",
        version: payload?.version || payload?.realtime?.version || "",
        sourceName: "Apps Script Web App (activities + costs + actual costs)",
      };
    },
  };

  realtimeSyncController.start();
})(window);
