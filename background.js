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

// ============================================
// Auto-Refresh Logic (Incremental Sync)
// ============================================

chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === "spAutoRefresh") {
        performAutoRefresh();
    }
});

// Create alarm if it doesn't exist
chrome.alarms.get("spAutoRefresh", (alarm) => {
    if (!alarm) {
        chrome.alarms.create("spAutoRefresh", { periodInMinutes: 3 }); // Run every 3 minutes
    }
});

async function performAutoRefresh() {
    try {
        const data = await chrome.storage.local.get(["configUrl", "searchPaths", "fileData", "lastSyncTime"]);
        const { configUrl, searchPaths, fileData, lastSyncTime } = data;
        
        if (!configUrl || !fileData || !lastSyncTime || !searchPaths) return;
        
        // Only fetch changes if we have checked paths
        const checkedPaths = searchPaths.filter(p => p.checked);
        if (checkedPaths.length === 0 && searchPaths.length > 0) return;

        const url = new URL(configUrl);
        const siteUrl = url.origin + '/' + url.pathname.split('/').filter(Boolean).slice(0,2).join('/');
        const folderPath = decodeURIComponent(url.pathname);

        // Calculate time with a 90-minute buffer to account for extremely slow SharePoint indexing
        const syncDate = new Date(lastSyncTime);
        syncDate.setMinutes(syncDate.getMinutes() - 90); 
        const timeFilter = syncDate.toISOString();

        // SharePoint search query to find recently modified documents
        // Format: path:"https://.../Shared Documents/*"
        const queryText = `path:"${configUrl}/*" IsDocument=True LastModifiedTime>="${timeFilter}"`;
        const searchUrl = `${siteUrl}/_api/search/query?querytext='${encodeURIComponent(queryText)}'&selectproperties='Title,Path,Size,LastModifiedTime'&rowlimit=500`;

        const searchData = await handleSharePointFetch(searchUrl, {
            method: "GET",
            headers: { "Accept": "application/json;odata=verbose" }
        });

        const rows = searchData?.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows?.results || [];
        if (rows.length === 0) {
            await chrome.storage.local.set({ lastSyncTime: new Date().toISOString() });
            return;
        }

        let updatedCount = 0;

        rows.forEach(row => {
            const cells = row.Cells.results;
            const getVal = (key) => cells.find(c => c.Key === key)?.Value;
            
            const title = getVal("Title");
            const absolutePath = getVal("Path");
            const size = parseInt(getVal("Size") || "0", 10);
            const modified = getVal("LastModifiedTime");

            if (!absolutePath) return;

            let relativePath = "";
            try {
                relativePath = decodeURIComponent(new URL(absolutePath).pathname);
            } catch(e) { return; }

            let targetFolderKey = folderPath;
            let foundMatch = false;

            const relPathLower = relativePath.toLowerCase();
            const folderPathLower = folderPath.toLowerCase();

            if (checkedPaths.length > 0) {
                for (const p of checkedPaths) {
                    const fullFolder = folderPath + p.path;
                    const fullFolderLower = fullFolder.toLowerCase();
                    if (relPathLower.startsWith(fullFolderLower + "/") || relPathLower === fullFolderLower) {
                        targetFolderKey = fullFolder;
                        foundMatch = true;
                        break; 
                    }
                }
            } else {
                if (relPathLower.startsWith(folderPathLower + "/")) {
                    foundMatch = true;
                }
            }

            if (!foundMatch) return; 

            const cachedFolder = fileData[targetFolderKey];
            if (!cachedFolder || !cachedFolder.files) return;

            const name = relativePath.split('/').pop() || title;

            // Strict case-insensitive deduplication check
            const existingIndex = cachedFolder.files.findIndex(f => f.path.toLowerCase() === relPathLower);
            if (existingIndex >= 0) {
                const oldTime = new Date(cachedFolder.files[existingIndex].modified).getTime();
                const newTime = new Date(modified).getTime();
                
                // Only update if time difference is greater than 1 second (handle formatting differences) or size changed
                if (Math.abs(oldTime - newTime) > 1000 || cachedFolder.files[existingIndex].size !== size) {
                    cachedFolder.files[existingIndex].modified = modified;
                    cachedFolder.files[existingIndex].size = size;
                    updatedCount++;
                }
            } else {
                cachedFolder.files.push({
                    name: name,
                    path: relativePath,
                    size: size,
                    modified: modified
                });
                updatedCount++;
            }
        });

        const storeUpdates = {
            lastSyncTime: new Date().toISOString(),
            lastSyncChanges: updatedCount
        };
        if (updatedCount > 0) {
            storeUpdates.fileData = fileData;
        }
        await chrome.storage.local.set(storeUpdates);

    } catch (e) {
        // Can fail if no SharePoint tab is open, which is fine
        if (e.message !== "NO_SP_TAB") {
            console.error("AutoRefresh failed:", e);
        }
    }
}
