// ============================================
// CONFIG - SharePoint site settings
// ============================================
const CONFIG = {
    siteUrl: "https://dxcportal.sharepoint.com/sites/msteams_268a65",
    folderPath: "/sites/msteams_268a65/Shared Documents"
};

// ============================================
// File Cache (per-path)
// ============================================
const fileCache = new Map(); // key: folderPath, value: { files, time }
const CACHE_DURATION = 10 * 60 * 1000; // 10 minutes

// ============================================
// Search Paths
// ============================================
let searchPaths = []; // user-defined sub-paths like ["/General", "/Reports"]

// ============================================
// DOM Elements
// ============================================
const searchInput = document.getElementById("searchInput");
const searchBtn = document.getElementById("searchBtn");
const clearBtn = document.getElementById("clearBtn");
const refreshBtn = document.getElementById("refreshBtn");
const pathInput = document.getElementById("pathInput");
const addPathBtn = document.getElementById("addPathBtn");
const pathTagsEl = document.getElementById("pathTags");
const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");

// ============================================
// Event Listeners
// ============================================
searchBtn.addEventListener("click", handleSearch);
clearBtn.addEventListener("click", () => {
    searchInput.value = "";
    clearBtn.classList.remove("visible");
    searchInput.focus();
});
refreshBtn.addEventListener("click", async () => {
    fileCache.clear();
    chrome.storage.local.remove("fileCache");
    showStatus("loading", "Refreshing file list...");
    try {
        const files = await loadFilesForPaths();
        showStatus("success", "✅ File list refreshed! (" + files.length + " files)");
    } catch (e) {
        showStatus("error", "❌ Refresh failed: " + e.message);
    }
});

addPathBtn.addEventListener("click", addPath);
pathInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") addPath();
});

// ============================================
// Path Management
// ============================================
function addPath() {
    let path = pathInput.value.trim();
    if (!path) return;

    // Normalize: ensure starts with /
    if (!path.startsWith("/")) path = "/" + path;
    // Remove trailing slash
    if (path.length > 1 && path.endsWith("/")) path = path.slice(0, -1);

    // Check duplicate
    if (searchPaths.includes(path)) {
        pathInput.value = "";
        return;
    }

    searchPaths.push(path);
    pathInput.value = "";
    renderPathTags();
    savePaths();
}

function removePath(path) {
    searchPaths = searchPaths.filter(p => p !== path);
    renderPathTags();
    savePaths();
}

function renderPathTags() {
    pathTagsEl.innerHTML = "";
    if (searchPaths.length === 0) {
        pathTagsEl.innerHTML = '<span class="path-hint">All folders (Shared Documents)</span>';
        return;
    }
    searchPaths.forEach(path => {
        const tag = document.createElement("span");
        tag.className = "path-tag";
        tag.innerHTML = escapeHtml(path) +
            '<button class="path-tag-remove" title="Remove">✕</button>';
        tag.querySelector(".path-tag-remove").addEventListener("click", () => removePath(path));
        pathTagsEl.appendChild(tag);
    });
}

searchInput.addEventListener("input", () => {
    clearBtn.classList.toggle("visible", searchInput.value.length > 0);
});

searchInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") handleSearch();
});

// Auto focus on open
searchInput.focus();

// Restore saved state and check connection on popup open
restoreState();
checkConnection();

// ============================================
// Persistent Storage (chrome.storage.local)
// ============================================
function savePaths() {
    chrome.storage.local.set({ searchPaths });
}

function saveCache() {
    // Convert Map to plain object for storage
    const cacheObj = {};
    fileCache.forEach((value, key) => {
        cacheObj[key] = value;
    });
    chrome.storage.local.set({ fileCache: cacheObj });
}

async function restoreState() {
    try {
        const data = await chrome.storage.local.get(["searchPaths", "fileCache"]);

        // Restore paths
        if (data.searchPaths && Array.isArray(data.searchPaths)) {
            searchPaths = data.searchPaths;
            renderPathTags();
        }

        // Restore cache (only if not expired)
        if (data.fileCache && typeof data.fileCache === "object") {
            const now = Date.now();
            for (const [key, value] of Object.entries(data.fileCache)) {
                if (value.time && (now - value.time < CACHE_DURATION)) {
                    fileCache.set(key, value);
                }
            }
        }
    } catch (e) {
        console.warn("Could not restore state:", e);
    }
}

async function checkConnection() {
    const statusEl = document.getElementById("connectionStatus");
    try {
        const tabs = await chrome.tabs.query({ url: "https://dxcportal.sharepoint.com/*" });
        if (tabs && tabs.length > 0) {
            statusEl.textContent = "✅ Connected to SharePoint";
            statusEl.style.color = "#4caf50";
        } else {
            statusEl.textContent = "⚠ No SharePoint tab found. Please login first.";
            statusEl.style.color = "#ff9800";
        }
    } catch (e) {
        statusEl.textContent = "⚠ Could not check connection";
        statusEl.style.color = "#ff9800";
    }
}

// ============================================
// Main Search Handler
// ============================================
async function handleSearch() {
    const query = searchInput.value.trim();
    if (!query) {
        showStatus("error", "⚠️ Please enter a search keyword!");
        return;
    }

    resultsEl.innerHTML = "";
    searchBtn.disabled = true;

    try {
        // Load files (uses per-path cache)
        showStatus("loading", "Loading file list...");
        const allFiles = await loadFilesForPaths();

        // Filter by filename
        const queryLower = query.toLowerCase();
        const matched = allFiles.filter(f =>
            f.name.toLowerCase().includes(queryLower)
        );
        displayResults(matched, query);
    } catch (error) {
        console.error("Search error:", error);

        if (error.message.includes("NO_SP_TAB")) {
            showStatus(
                "error",
                "🔒 No SharePoint tab found! Please open <a href='https://dxcportal.sharepoint.com/sites/msteams_268a65' target='_blank'>SharePoint</a> in a tab, login, and try again."
            );
        } else if (error.message.includes("403") || error.message.includes("401") || error.message.includes("429")) {
            showStatus(
                "error",
                "🔒 Not authenticated or throttled! Please open SharePoint in a tab and login first."
            );
        } else {
            showStatus("error", "❌ Error: " + error.message);
        }
    } finally {
        searchBtn.disabled = false;
    }
}

// ============================================
// Helper: Send fetch via background service worker
// ============================================
function bgFetch(url, options = {}) {
    return new Promise((resolve, reject) => {
        chrome.runtime.sendMessage(
            { action: "fetchSharePoint", url, options },
            (response) => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                } else if (!response || !response.success) {
                    reject(new Error(response ? response.error : "No response from background"));
                } else {
                    resolve(response.data);
                }
            }
        );
    });
}

// ============================================
// Load All Files (with rate limiting to avoid throttle)
// Fetches once and caches - subsequent searches use cache
// ============================================
const BATCH_SIZE = 3; // Max concurrent API calls
const DELAY_MS = 300; // Delay between batches

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Build the list of root folders to search based on user paths
function getSearchRoots() {
    if (searchPaths.length === 0) {
        return [CONFIG.folderPath];
    }
    return searchPaths.map(p => CONFIG.folderPath + p);
}

// Load files for all user-defined paths, using per-path cache
async function loadFilesForPaths() {
    const roots = getSearchRoots();
    const allFiles = [];
    const now = Date.now();

    // Determine which roots need fresh loading
    const toLoad = [];
    for (const root of roots) {
        const cached = fileCache.get(root);
        if (cached && (now - cached.time < CACHE_DURATION)) {
            allFiles.push(...cached.files);
        } else {
            toLoad.push(root);
        }
    }

    if (toLoad.length === 0) return allFiles;

    // Load uncached paths
    for (const root of toLoad) {
        const files = await loadFilesFromFolder(root);
        fileCache.set(root, { files, time: Date.now() });
        allFiles.push(...files);
    }

    // Persist cache to storage
    saveCache();

    return allFiles;
}

async function loadFilesFromFolder(rootPath) {
    const allFiles = [];
    const queue = [rootPath];

    while (queue.length > 0) {
        const batch = queue.splice(0, BATCH_SIZE);

        const batchResults = await Promise.all(
            batch.map(folderPath => fetchFolderContents(folderPath))
        );

        for (const result of batchResults) {
            allFiles.push(...result.files);
            queue.push(...result.subfolders);
        }

        showStatus("loading", "Loading... (" + allFiles.length + " files found, " + queue.length + " folders remaining)");

        if (queue.length > 0) {
            await delay(DELAY_MS);
        }
    }

    return allFiles;
}

async function fetchFolderContents(folderPath) {
    const files = [];
    const subfolders = [];

    try {
        // Fetch files
        const filesUrl = CONFIG.siteUrl +
            "/_api/web/GetFolderByServerRelativeUrl('" +
            encodeURIComponent(folderPath) +
            "')/Files?$top=5000&$select=Name,ServerRelativeUrl,Length,TimeLastModified";

        const filesData = await bgFetch(filesUrl, {
            method: "GET",
            headers: { "Accept": "application/json;odata=verbose" }
        });

        for (const file of (filesData.d?.results || [])) {
            files.push({
                name: file.Name,
                path: file.ServerRelativeUrl,
                size: parseInt(file.Length || 0),
                modified: file.TimeLastModified || ""
            });
        }

        // Fetch subfolders
        const foldersUrl = CONFIG.siteUrl +
            "/_api/web/GetFolderByServerRelativeUrl('" +
            encodeURIComponent(folderPath) +
            "')/Folders?$select=Name,ServerRelativeUrl";

        const foldersData = await bgFetch(foldersUrl, {
            method: "GET",
            headers: { "Accept": "application/json;odata=verbose" }
        });

        for (const folder of (foldersData.d?.results || [])) {
            if (folder.Name !== "Forms" && !folder.Name.startsWith("_")) {
                subfolders.push(folder.ServerRelativeUrl);
            }
        }
    } catch (e) {
        console.warn("Error fetching: " + folderPath, e);
    }

    return { files, subfolders };
}

// ============================================
// Display Results
// ============================================
function displayResults(files, query) {
    resultsEl.innerHTML = "";

    if (files.length === 0) {
        showStatus("error", "");
        resultsEl.innerHTML =
            '<div class="no-results">' +
            '<div class="no-results-icon">🔍</div>' +
            '<div class="no-results-text">No files found matching "<strong>' +
            escapeHtml(query) +
            '</strong>"</div>' +
            '</div>';
        return;
    }

    showStatus(
        "success",
        '✅ Found ' + files.length + ' file(s) matching "' + escapeHtml(query) + '"'
    );

    // Result header
    const header = document.createElement("div");
    header.className = "result-header";
    header.textContent = "📁 Results (" + files.length + " files)";
    resultsEl.appendChild(header);

    // File items
    files.forEach(function (file) {
        // List Items API returns server-relative path
        const fileUrl = CONFIG.siteUrl.split("/sites/")[0] + file.path;
        const parentFolder =
            file.path
                .replace("/" + file.name, "")
                .replace(CONFIG.folderPath, "") || "/";

        const item = document.createElement("a");
        item.className = "file-item";
        item.href = fileUrl;
        item.target = "_blank";
        item.title = "Open: " + file.name;

        item.innerHTML =
            '<span class="file-icon">' + getFileIcon(file.name) + '</span>' +
            '<div class="file-info">' +
            '<div class="file-name">' + highlightText(escapeHtml(file.name), query) + '</div>' +
            '<div class="file-path">📂 ' + escapeHtml(parentFolder) + '</div>' +
            '<div class="file-size">' + formatFileSize(file.size) + '</div>' +
            '</div>' +
            '<span class="file-arrow">→</span>';

        resultsEl.appendChild(item);
    });
}

// ============================================
// Utility Functions
// ============================================
function showStatus(type, message) {
    statusEl.className = "status " + type;
    if (type === "loading") {
        statusEl.innerHTML = '<span class="spinner"></span> ' + message;
    } else {
        statusEl.innerHTML = message;
    }
    if (!message) {
        statusEl.classList.add("hidden");
    } else {
        statusEl.classList.remove("hidden");
    }
}

function getFileIcon(filename) {
    const ext = filename.split(".").pop().toLowerCase();
    const icons = {
        pdf: "📕", doc: "📘", docx: "📘",
        xls: "📗", xlsx: "📗", csv: "📗",
        ppt: "📙", pptx: "📙",
        txt: "📄", md: "📄", json: "📄", xml: "📄",
        jpg: "🖼️", jpeg: "🖼️", png: "🖼️", gif: "🖼️", svg: "🖼️",
        mp4: "🎬", avi: "🎬", mov: "🎬",
        mp3: "🎵", wav: "🎵",
        zip: "📦", rar: "📦",
        html: "🌐", css: "🎨", js: "⚡",
        py: "🐍", java: "☕", sql: "🗃️"
    };
    return icons[ext] || "📄";
}

function formatFileSize(bytes) {
    if (!bytes || bytes === 0) return "";
    var units = ["B", "KB", "MB", "GB"];
    var i = 0;
    var size = bytes;
    while (size >= 1024 && i < units.length - 1) {
        size /= 1024;
        i++;
    }
    return size.toFixed(1) + " " + units[i];
}

function highlightText(text, query) {
    if (!query) return text;
    var regex = new RegExp("(" + escapeRegex(query) + ")", "gi");
    return text.replace(regex, '<span class="highlight">$1</span>');
}

function escapeHtml(str) {
    var div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
}

function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}