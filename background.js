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
        handleSharePointFetch(request.url, request.options)
            .then(data => sendResponse({ success: true, data }))
            .catch(err => sendResponse({ success: false, error: err.message }));
        return true;
    }

    if (request.action === "checkConnection") {
        chrome.tabs.query({ url: "https://dxcportal.sharepoint.com/*" })
            .then(tabs => sendResponse({ connected: tabs && tabs.length > 0 }))
            .catch(() => sendResponse({ connected: false }));
        return true;
    }
});

async function handleSharePointFetch(url, options = {}) {
    const tabs = await chrome.tabs.query({
        url: "https://dxcportal.sharepoint.com/*"
    });

    if (!tabs || tabs.length === 0) {
        throw new Error("NO_SP_TAB");
    }

    const tabId = tabs[0].id;

    const results = await chrome.scripting.executeScript({
        target: { tabId },
        func: async (fetchUrl, fetchOptions) => {
            try {
                const response = await fetch(fetchUrl, {
                    ...fetchOptions,
                    credentials: "same-origin"
                });
                if (!response.ok) {
                    return { error: "HTTP " + response.status + ": " + response.statusText };
                }
                const data = await response.json();
                return { data };
            } catch (e) {
                return { error: e.message };
            }
        },
        args: [url, options]
    });

    const result = results[0]?.result;

    if (!result) {
        throw new Error("Failed to execute script in SharePoint tab");
    }

    if (result.error) {
        throw new Error(result.error);
    }

    return result.data;
}
