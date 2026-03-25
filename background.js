// ============================================
// Background Service Worker
// Injects fetch into SharePoint tab so cookies are sent
// ============================================

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "fetchSharePoint") {
        handleSharePointFetch(request.url, request.options)
            .then(data => sendResponse({ success: true, data }))
            .catch(err => sendResponse({ success: false, error: err.message }));
        return true; // Keep message channel open for async response
    }
});

async function handleSharePointFetch(url, options = {}) {
    // Find a SharePoint tab that is already logged in
    const tabs = await chrome.tabs.query({
        url: "https://dxcportal.sharepoint.com/*"
    });

    if (!tabs || tabs.length === 0) {
        throw new Error("NO_SP_TAB");
    }

    const tabId = tabs[0].id;

    // Inject fetch into the SharePoint tab context
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
