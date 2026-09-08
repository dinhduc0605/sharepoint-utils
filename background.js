// ============================================
// Background Service Worker
// 1. Handles extension icon click → inject content script
// 2. Proxies SharePoint API calls via tab injection
// ============================================

// When extension icon is clicked, inject the modal into the active tab
chrome.action.onClicked.addListener(async (tab) => {
    try {
        // Inject CSS first, then JS
        await chrome.scripting.insertCSS({
            target: { tabId: tab.id },
            files: ["content.css"]
        });
        await chrome.scripting.executeScript({
            target: { tabId: tab.id },
            files: ["content.js"]
        });
    } catch (e) {
        console.error("Failed to inject content script:", e);
    }
});

// Handle messages from content script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "fetchSharePoint") {
        handleSharePointFetch(request.url, request.options, sender.tab?.id)
            .then(data => sendResponse({ success: true, data }))
            .catch(err => sendResponse({ success: false, error: err.message }));
        return true;
    }

    if (request.action === "checkConnection") {
        chrome.tabs.query({ url: "https://*.sharepoint.com/*" })
            .then(tabs => sendResponse({ connected: tabs && tabs.length > 0 }))
            .catch(() => sendResponse({ connected: false }));
        return true;
    }
});

async function handleSharePointFetch(url, options = {}, preferredTabId = null) {
    let targetOrigin;
    try {
        targetOrigin = new URL(url).origin;
    } catch (e) {
        throw new Error("INVALID_SHAREPOINT_URL");
    }

    const tabs = await chrome.tabs.query({
        url: targetOrigin + "/*"
    });

    if (!tabs || tabs.length === 0) {
        throw new Error("NO_SP_TAB");
    }

    // Filter out discarded tabs (Chrome unloads tabs after long idle periods).
    // Discarded tabs still appear in tabs.query() but executeScript() will fail on them.
    const activeTabs = tabs.filter(t => !t.discarded);
    const candidates = (activeTabs.length > 0 ? activeTabs : tabs).sort((a, b) => {
        if (preferredTabId !== null) {
            if (a.id === preferredTabId) return -1;
            if (b.id === preferredTabId) return 1;
        }
        return Number(Boolean(b.active)) - Number(Boolean(a.active));
    });

    let lastError = null;

    for (const tab of candidates) {
        try {
            const results = await chrome.scripting.executeScript({
                target: { tabId: tab.id },
                func: async (fetchUrl, fetchOptions) => {
                    const REQUEST_TIMEOUT_MS = 8000;

                    async function fetchWithTimeout(targetUrl, init) {
                        const controller = new AbortController();
                        const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
                        try {
                            return await fetch(targetUrl, {
                                ...init,
                                signal: controller.signal
                            });
                        } finally {
                            clearTimeout(timeoutId);
                        }
                    }

                    async function getRequestDigest(targetUrl) {
                        const existingDigest = document.getElementById("__REQUESTDIGEST")?.value;
                        if (existingDigest) return existingDigest;

                        const parsedUrl = new URL(targetUrl);
                        const apiIndex = parsedUrl.pathname.indexOf("/_api/");
                        const webBasePath = apiIndex >= 0 ? parsedUrl.pathname.slice(0, apiIndex) : "";
                        const contextInfoUrl = parsedUrl.origin + webBasePath + "/_api/contextinfo";
                        const contextResponse = await fetchWithTimeout(contextInfoUrl, {
                            method: "POST",
                            credentials: "same-origin",
                            headers: {
                                "Accept": "application/json;odata=verbose"
                            }
                        });

                        if (!contextResponse.ok) {
                            throw new Error("HTTP " + contextResponse.status + ": " + contextResponse.statusText);
                        }

                        const contextData = await contextResponse.json();
                        const digest = contextData?.d?.GetContextWebInformation?.FormDigestValue;
                        if (!digest) {
                            throw new Error("Could not get SharePoint request digest");
                        }

                        return digest;
                    }

                    try {
                        const method = (fetchOptions?.method || "GET").toUpperCase();
                        const headers = {
                            ...(fetchOptions?.headers || {})
                        };

                        if (method !== "GET" && method !== "HEAD" && !headers["X-RequestDigest"]) {
                            headers["X-RequestDigest"] = await getRequestDigest(fetchUrl);
                        }

                        const response = await fetchWithTimeout(fetchUrl, {
                            ...fetchOptions,
                            headers,
                            credentials: "same-origin"
                        });
                        if (!response.ok) {
                            return { error: "HTTP " + response.status + ": " + response.statusText };
                        }
                        const text = await response.text();
                        let data = null;
                        if (text) {
                            try {
                                data = JSON.parse(text);
                            } catch (e) {
                                data = { raw: text };
                            }
                        }
                        return { data };
                    } catch (e) {
                        if (e.name === "AbortError") {
                            return { error: "REQUEST_TIMEOUT" };
                        }
                        return { error: e.message };
                    }
                },
                args: [url, options]
            });

            const result = results[0]?.result;
            if (!result) {
                lastError = new Error("Failed to execute script in SharePoint tab");
                continue;
            }
            if (result.error) {
                throw new Error(result.error);
            }
            return result.data;
        } catch (e) {
            // If this tab failed due to being discarded/inaccessible, try the next one
            if (e.message && (
                e.message.includes("Cannot access") ||
                e.message.includes("discarded") ||
                e.message.includes("No tab with id")
            )) {
                lastError = e;
                continue;
            }
            // Any other error (e.g. HTTP 403) should propagate immediately
            throw e;
        }
    }

    // All candidate tabs failed
    if (lastError && lastError.message && lastError.message.includes("Cannot access")) {
        throw new Error("SP_TAB_DISCARDED");
    }
    throw lastError || new Error("Failed to execute script in SharePoint tab");
}
