// ============================================
// SharePoint File Search - Content Script
// Injected into the active tab to show a modal overlay
// ============================================
(function () {
    // Prevent double injection
    if (document.getElementById("sp-search-backdrop")) {
        document.getElementById("sp-search-backdrop").remove();
    }

    // ============================================
    // CONFIG
    // ============================================
    const CONFIG = {
        siteUrl: "https://dxcportal.sharepoint.com/sites/msteams_268a65",
        folderPath: "/sites/msteams_268a65/Shared Documents"
    };

    const BATCH_SIZE = 3;
    const DELAY_MS = 300;
    const MIN_QUERY_LENGTH = 3;
    const DEBOUNCE_MS = 300;

    let searchPaths = [];
    const fileCache = new Map();
    let debounceTimer = null;
    let isLoading = false;

    // ============================================
    // Build Modal HTML
    // ============================================
    const backdrop = document.createElement("div");
    backdrop.id = "sp-search-backdrop";
    backdrop.className = "sp-search-backdrop";
    backdrop.innerHTML = `
        <div class="sp-search-modal" id="sp-search-modal">
            <div class="sp-search-header">
                <span class="sp-search-header-icon">📂</span>
                <h2>SharePoint File Search</h2>
                <button class="sp-search-close" id="sp-close-btn" title="Close (Esc)">✕</button>
            </div>
            <div class="sp-path-section">
                <div class="sp-path-row">
                    <label>📂</label>
                    <input type="text" class="sp-path-input" id="sp-path-input" placeholder="/subfolder/path" autocomplete="off">
                    <button class="sp-add-path-btn" id="sp-add-path-btn">+ Add</button>
                </div>
                <div class="sp-path-tags" id="sp-path-tags">
                    <span class="sp-path-hint">All folders (Shared Documents)</span>
                </div>
            </div>
            <div class="sp-search-input-section">
                <div class="sp-search-row">
                    <div class="sp-search-input-wrap">
                        <span class="sp-search-icon">🔍</span>
                        <input type="text" class="sp-search-input" id="sp-search-input" placeholder="Type filename to search..." autocomplete="off">
                        <button class="sp-clear-btn" id="sp-clear-btn" title="Clear">✕</button>
                    </div>
                    <button class="sp-refresh-btn" id="sp-refresh-btn" title="Refresh file list">🔄</button>
                </div>
                <div class="sp-search-hint">Type at least 3 characters to search</div>
            </div>
            <div class="sp-status sp-hidden" id="sp-status"></div>
            <div class="sp-results" id="sp-results"></div>
            <div class="sp-footer">
                <span id="sp-connection-status">⏳ Checking connection...</span>
            </div>
        </div>
    `;

    document.body.appendChild(backdrop);

    // ============================================
    // DOM References
    // ============================================
    const modal = document.getElementById("sp-search-modal");
    const searchInput = document.getElementById("sp-search-input");
    const clearBtn = document.getElementById("sp-clear-btn");
    const refreshBtn = document.getElementById("sp-refresh-btn");
    const pathInput = document.getElementById("sp-path-input");
    const addPathBtn = document.getElementById("sp-add-path-btn");
    const pathTagsEl = document.getElementById("sp-path-tags");
    const statusEl = document.getElementById("sp-status");
    const resultsEl = document.getElementById("sp-results");
    const closeBtn = document.getElementById("sp-close-btn");
    const connectionEl = document.getElementById("sp-connection-status");

    // ============================================
    // Events
    // ============================================
    // Close modal
    closeBtn.addEventListener("click", closeModal);
    backdrop.addEventListener("click", (e) => {
        if (e.target === backdrop) closeModal();
    });
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && document.getElementById("sp-search-backdrop")) {
            closeModal();
        }
    });

    // Prevent modal clicks from propagating
    modal.addEventListener("click", (e) => e.stopPropagation());

    // Search input - live search with debounce
    searchInput.addEventListener("input", () => {
        const val = searchInput.value;
        clearBtn.classList.toggle("sp-visible", val.length > 0);

        clearTimeout(debounceTimer);
        if (val.trim().length >= MIN_QUERY_LENGTH) {
            debounceTimer = setTimeout(() => handleSearch(), DEBOUNCE_MS);
        } else if (val.trim().length === 0) {
            resultsEl.innerHTML = "";
            showStatus("", "");
        }
    });

    clearBtn.addEventListener("click", () => {
        searchInput.value = "";
        clearBtn.classList.remove("sp-visible");
        resultsEl.innerHTML = "";
        showStatus("", "");
        searchInput.focus();
    });

    // Refresh
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

    // Path management
    addPathBtn.addEventListener("click", addPath);
    pathInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") addPath();
    });

    // Focus input
    searchInput.focus();

    // Restore state and check connection
    restoreState();
    checkConnection();

    // ============================================
    // Close Modal
    // ============================================
    function closeModal() {
        const el = document.getElementById("sp-search-backdrop");
        if (el) el.remove();
    }

    // ============================================
    // Path Management
    // ============================================
    function addPath() {
        let path = pathInput.value.trim();
        if (!path) return;
        if (!path.startsWith("/")) path = "/" + path;
        if (path.length > 1 && path.endsWith("/")) path = path.slice(0, -1);
        if (searchPaths.includes(path)) { pathInput.value = ""; return; }

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
            pathTagsEl.innerHTML = '<span class="sp-path-hint">All folders (Shared Documents)</span>';
            return;
        }
        searchPaths.forEach(path => {
            const tag = document.createElement("span");
            tag.className = "sp-path-tag";
            tag.innerHTML = escapeHtml(path) + '<button class="sp-path-tag-remove" title="Remove">✕</button>';
            tag.querySelector(".sp-path-tag-remove").addEventListener("click", () => removePath(path));
            pathTagsEl.appendChild(tag);
        });
    }

    // ============================================
    // Persistent Storage
    // ============================================
    function savePaths() { chrome.storage.local.set({ searchPaths }); }

    function saveCache() {
        const cacheObj = {};
        fileCache.forEach((value, key) => { cacheObj[key] = value; });
        chrome.storage.local.set({ fileCache: cacheObj });
    }

    async function restoreState() {
        try {
            const data = await chrome.storage.local.get(["searchPaths", "fileCache"]);
            if (data.searchPaths && Array.isArray(data.searchPaths)) {
                searchPaths = data.searchPaths;
                renderPathTags();
            }
            if (data.fileCache && typeof data.fileCache === "object") {
                for (const [key, value] of Object.entries(data.fileCache)) {
                    fileCache.set(key, value);
                }
            }
        } catch (e) {
            console.warn("Could not restore state:", e);
        }
    }

    async function checkConnection() {
        try {
            chrome.runtime.sendMessage({ action: "checkConnection" }, (response) => {
                if (chrome.runtime.lastError) {
                    connectionEl.textContent = "⚠ Could not check connection";
                    connectionEl.style.color = "#ff9800";
                    return;
                }
                if (response && response.connected) {
                    connectionEl.textContent = "✅ Connected to SharePoint";
                    connectionEl.style.color = "#4caf50";
                } else {
                    connectionEl.textContent = "⚠ No SharePoint tab found. Please login first.";
                    connectionEl.style.color = "#ff9800";
                }
            });
        } catch (e) {
            connectionEl.textContent = "⚠ Could not check connection";
            connectionEl.style.color = "#ff9800";
        }
    }

    // ============================================
    // Search Handler
    // ============================================
    async function handleSearch() {
        const query = searchInput.value.trim();
        if (query.length < MIN_QUERY_LENGTH) return;
        if (isLoading) return;

        resultsEl.innerHTML = "";

        try {
            isLoading = true;
            showStatus("loading", "Loading file list...");
            const allFiles = await loadFilesForPaths();

            // Re-check query (might have changed during loading)
            const currentQuery = searchInput.value.trim();
            if (currentQuery.length < MIN_QUERY_LENGTH) {
                showStatus("", "");
                return;
            }

            const queryLower = currentQuery.toLowerCase();
            const matched = allFiles.filter(f => f.name.toLowerCase().includes(queryLower));
            displayResults(matched, currentQuery);
        } catch (error) {
            console.error("Search error:", error);
            if (error.message.includes("NO_SP_TAB")) {
                showStatus("error", "🔒 No SharePoint tab found! Please open SharePoint in a tab and login first.");
            } else if (error.message.includes("403") || error.message.includes("401") || error.message.includes("429")) {
                showStatus("error", "🔒 Not authenticated or throttled! Please open SharePoint in a tab and login first.");
            } else {
                showStatus("error", "❌ Error: " + error.message);
            }
        } finally {
            isLoading = false;
        }
    }

    // ============================================
    // Background Fetch Proxy
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
    // File Loading (rate-limited, cached)
    // ============================================
    function delayMs(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

    function getSearchRoots() {
        if (searchPaths.length === 0) return [CONFIG.folderPath];
        return searchPaths.map(p => CONFIG.folderPath + p);
    }

    async function loadFilesForPaths() {
        const roots = getSearchRoots();
        const allFiles = [];
        const toLoad = [];

        for (const root of roots) {
            const cached = fileCache.get(root);
            if (cached) {
                allFiles.push(...cached.files);
            } else {
                toLoad.push(root);
            }
        }

        if (toLoad.length === 0) return allFiles;

        for (const root of toLoad) {
            const files = await loadFilesFromFolder(root);
            fileCache.set(root, { files, time: Date.now() });
            allFiles.push(...files);
        }

        saveCache();
        return allFiles;
    }

    async function loadFilesFromFolder(rootPath) {
        const allFiles = [];
        const queue = [rootPath];

        while (queue.length > 0) {
            const batch = queue.splice(0, BATCH_SIZE);
            const batchResults = await Promise.all(
                batch.map(fp => fetchFolderContents(fp))
            );

            for (const result of batchResults) {
                allFiles.push(...result.files);
                queue.push(...result.subfolders);
            }

            showStatus("loading", "Loading... (" + allFiles.length + " files, " + queue.length + " folders remaining)");

            if (queue.length > 0) await delayMs(DELAY_MS);
        }

        return allFiles;
    }

    async function fetchFolderContents(folderPath) {
        const files = [];
        const subfolders = [];

        try {
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
            showStatus("", "");
            resultsEl.innerHTML =
                '<div class="sp-no-results">' +
                '<div class="sp-no-results-icon">🔍</div>' +
                '<div class="sp-no-results-text">No files found matching "<strong>' +
                escapeHtml(query) + '</strong>"</div></div>';
            return;
        }

        showStatus("success", '✅ Found ' + files.length + ' file(s) matching "' + escapeHtml(query) + '"');

        files.forEach(function (file) {
            const fileUrl = CONFIG.siteUrl.split("/sites/")[0] + file.path;
            const fullPath = file.path.replace(CONFIG.folderPath, "Shared Documents");

            const item = document.createElement("a");
            item.className = "sp-file-item";
            item.href = fileUrl;
            item.target = "_blank";
            item.title = "Open: " + file.name;

            item.innerHTML =
                '<span class="sp-file-icon">' + getFileIcon(file.name) + '</span>' +
                '<div class="sp-file-info">' +
                '<div class="sp-file-name">' + highlightText(escapeHtml(file.name), query) + '</div>' +
                '<div class="sp-file-path">📂 ' + escapeHtml(fullPath) + '</div>' +
                (file.size ? '<div class="sp-file-size">' + formatFileSize(file.size) + '</div>' : '') +
                '</div>' +
                '<span class="sp-file-arrow">→</span>';

            resultsEl.appendChild(item);
        });
    }

    // ============================================
    // Utility Functions
    // ============================================
    function showStatus(type, message) {
        if (!type && !message) { statusEl.className = "sp-status sp-hidden"; return; }
        statusEl.className = "sp-status sp-" + type;
        statusEl.innerHTML = type === "loading"
            ? '<span class="sp-spinner"></span> ' + message
            : message;
        if (!message) statusEl.classList.add("sp-hidden");
        else statusEl.classList.remove("sp-hidden");
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
        while (size >= 1024 && i < units.length - 1) { size /= 1024; i++; }
        return size.toFixed(1) + " " + units[i];
    }

    function highlightText(text, query) {
        if (!query) return text;
        var regex = new RegExp("(" + escapeRegex(query) + ")", "gi");
        return text.replace(regex, '<span class="sp-highlight">$1</span>');
    }

    function escapeHtml(str) {
        var div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    }

    function escapeRegex(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
})();
