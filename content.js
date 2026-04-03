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

    // searchPaths: array of { path: string, checked: boolean }
    let searchPaths = [];
    const fileData = new Map();
    let debounceTimer = null;
    let isLoading = false;
    let lastSearchResults = []; // store last search results for re-filtering
    let lastQuery = "";

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
                <h2>SharePoint Utils</h2>
                <button class="sp-search-close" id="sp-close-btn" title="Close (Esc)">✕</button>
            </div>

            <!-- Tab Bar -->
            <div class="sp-tab-bar">
                <button class="sp-tab-btn sp-active" data-tab="search">🔍 Search</button>
                <button class="sp-tab-btn" data-tab="settings">⚙️ Settings</button>
            </div>

            <!-- Search Tab -->
            <div class="sp-tab-panel sp-active" id="sp-tab-search">
                <div class="sp-search-input-section">
                    <div class="sp-search-row">
                        <div class="sp-search-input-wrap">
                            <span class="sp-search-icon">🔍</span>
                            <input type="text" class="sp-search-input" id="sp-search-input" placeholder="Type filename..." autocomplete="off">
                            <button class="sp-clear-btn" id="sp-clear-btn" title="Clear">✕</button>
                        </div>
                        <select class="sp-filter-select" id="sp-filter-select" title="Filter by subfolder">
                            <option value="">All folders</option>
                        </select>
                    </div>
                    <div class="sp-search-hint">Type at least 3 characters to search</div>
                </div>
                <div class="sp-status sp-hidden" id="sp-status"></div>
                <div class="sp-results" id="sp-results"></div>
            </div>

            <!-- Settings Tab -->
            <div class="sp-tab-panel" id="sp-tab-settings">
                <div class="sp-settings-content">
                    <div class="sp-import-export-section">
                        <div class="sp-settings-label">Import / Export</div>
                        <div class="sp-ie-row">
                            <button class="sp-ie-btn sp-export-btn" id="sp-export-btn">📤 Export CSV</button>
                            <button class="sp-ie-btn sp-import-btn" id="sp-import-btn">📥 Import CSV</button>
                            <input type="file" id="sp-import-file" accept=".csv" style="display:none">
                        </div>
                        <div class="sp-refresh-hint">Export/import paths and cached file data</div>
                    </div>
                    <div class="sp-settings-section">
                        <div class="sp-settings-label">Add Subfolder</div>
                        <div class="sp-path-add-row">
                            <input type="text" class="sp-path-input" id="sp-path-input" placeholder="/subfolder/path" autocomplete="off">
                            <button class="sp-add-path-btn" id="sp-add-path-btn">+ Add</button>
                        </div>
                    </div>
                    <div class="sp-settings-section">
                        <div class="sp-settings-label">Subfolders</div>
                        <div class="sp-path-list" id="sp-path-list">
                            <div class="sp-path-list-empty">No subfolders added. Search will scan all of Shared Documents.</div>
                        </div>
                    </div>
                    <div class="sp-refresh-section">
                        <div class="sp-status sp-hidden" id="sp-settings-status"></div>
                        <button class="sp-refresh-full-btn" id="sp-refresh-btn">🔄 Refresh checked folders</button>
                        <div class="sp-refresh-hint">Re-fetch file list from checked subfolders</div>
                    </div>
                </div>
            </div>

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
    const pathListEl = document.getElementById("sp-path-list");
    const statusEl = document.getElementById("sp-status");
    const resultsEl = document.getElementById("sp-results");
    const closeBtn = document.getElementById("sp-close-btn");
    const connectionEl = document.getElementById("sp-connection-status");
    const settingsStatusEl = document.getElementById("sp-settings-status");
    const filterSelect = document.getElementById("sp-filter-select");
    const exportBtn = document.getElementById("sp-export-btn");
    const importBtn = document.getElementById("sp-import-btn");
    const importFile = document.getElementById("sp-import-file");
    const tabBtns = backdrop.querySelectorAll(".sp-tab-btn");
    const tabPanels = backdrop.querySelectorAll(".sp-tab-panel");

    // ============================================
    // Tab Switching
    // ============================================
    tabBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            const tabId = btn.dataset.tab;
            tabBtns.forEach(b => b.classList.remove("sp-active"));
            tabPanels.forEach(p => p.classList.remove("sp-active"));
            btn.classList.add("sp-active");
            document.getElementById("sp-tab-" + tabId).classList.add("sp-active");
            if (tabId === "search") searchInput.focus();
        });
    });

    // ============================================
    // Events
    // ============================================
    closeBtn.addEventListener("click", closeModal);
    backdrop.addEventListener("click", (e) => {
        if (e.target === backdrop) closeModal();
    });
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && document.getElementById("sp-search-backdrop")) closeModal();
    });
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
            lastSearchResults = [];
            lastQuery = "";
        }
    });

    clearBtn.addEventListener("click", () => {
        searchInput.value = "";
        clearBtn.classList.remove("sp-visible");
        resultsEl.innerHTML = "";
        showStatus("", "");
        lastSearchResults = [];
        lastQuery = "";
        searchInput.focus();
    });

    // Filter dropdown — re-filter existing results instantly
    filterSelect.addEventListener("change", () => {
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(applyFilter(lastSearchResults), lastQuery);
        }
    });

    // Refresh - only re-fetch checked folders
    refreshBtn.addEventListener("click", async () => {
        const checkedRoots = getCheckedRoots();
        checkedRoots.forEach(r => fileData.delete(r));

        // Also remove from storage
        const dataObj = {};
        fileData.forEach((value, key) => { dataObj[key] = value; });
        chrome.storage.local.set({ fileData: dataObj });

        refreshBtn.disabled = true;
        refreshBtn.textContent = "⏳ Refreshing...";
        try {
            const files = await loadFilesForChecked();
            refreshBtn.textContent = "✅ Refreshed! (" + files.length + " files)";
            setTimeout(() => {
                refreshBtn.textContent = "🔄 Refresh checked folders";
                refreshBtn.disabled = false;
            }, 2000);
        } catch (e) {
            refreshBtn.textContent = "❌ Failed: " + e.message;
            setTimeout(() => {
                refreshBtn.textContent = "🔄 Refresh checked folders";
                refreshBtn.disabled = false;
            }, 3000);
        }
    });

    // Path management
    addPathBtn.addEventListener("click", addPath);
    pathInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") addPath();
    });

    // Import / Export
    exportBtn.addEventListener("click", exportCSV);
    importBtn.addEventListener("click", () => importFile.click());
    importFile.addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (file) importCSV(file);
        importFile.value = ""; // reset so same file can be re-imported
    });

    // Focus input & restore state
    searchInput.focus();
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
        if (searchPaths.some(p => p.path === path)) { pathInput.value = ""; return; }

        searchPaths.push({ path, checked: true });
        pathInput.value = "";
        renderPathList();
        savePaths();
        updateFilterOptions();
    }

    function removePath(path) {
        if (!confirm('Remove subfolder "' + path + '" and its cached files?')) return;
        searchPaths = searchPaths.filter(p => p.path !== path);
        fileData.delete(CONFIG.folderPath + path);
        renderPathList();
        savePaths();
        saveData();
        updateFilterOptions();
    }

    function togglePath(path, checked) {
        const item = searchPaths.find(p => p.path === path);
        if (item) item.checked = checked;
        savePaths();
    }

    function renderPathList() {
        pathListEl.innerHTML = "";
        if (searchPaths.length === 0) {
            pathListEl.innerHTML = '<div class="sp-path-list-empty">No subfolders added. Search will scan all of Shared Documents.</div>';
            return;
        }
        searchPaths.forEach((item, index) => {
            const cached = fileData.get(CONFIG.folderPath + item.path);
            const fileCount = cached ? cached.files.length : 0;

            const row = document.createElement("div");
            row.className = "sp-path-item";
            row.innerHTML =
                '<input type="checkbox" class="sp-path-checkbox" ' + (item.checked ? 'checked' : '') + '>' +
                '<span class="sp-path-text">' + escapeHtml(item.path) + '</span>' +
                '<input type="text" class="sp-path-edit-input" value="' + escapeHtml(item.path) + '" style="display:none">' +
                (cached ? '<span class="sp-path-file-count">' + fileCount + ' files</span>' : '') +
                '<button class="sp-path-edit" title="Edit">✏️</button>' +
                '<button class="sp-path-remove" title="Remove">🗑️</button>';

            row.querySelector(".sp-path-checkbox").addEventListener("change", (e) => {
                togglePath(item.path, e.target.checked);
            });

            // Edit functionality
            const textEl = row.querySelector(".sp-path-text");
            const editInput = row.querySelector(".sp-path-edit-input");
            const editBtn = row.querySelector(".sp-path-edit");

            editBtn.addEventListener("click", () => {
                if (editInput.style.display === "none") {
                    // Start editing
                    textEl.style.display = "none";
                    editInput.style.display = "block";
                    editInput.focus();
                    editBtn.textContent = "✅";
                } else {
                    // Save edit
                    let newPath = editInput.value.trim();
                    if (!newPath) { editInput.value = item.path; }
                    else {
                        if (!newPath.startsWith("/")) newPath = "/" + newPath;
                        if (newPath.length > 1 && newPath.endsWith("/")) newPath = newPath.slice(0, -1);
                        // Remove old cache
                        fileData.delete(CONFIG.folderPath + item.path);
                        item.path = newPath;
                        savePaths();
                        saveData();
                    }
                    textEl.textContent = item.path;
                    textEl.style.display = "";
                    editInput.style.display = "none";
                    editBtn.textContent = "✏️";
                }
            });

            editInput.addEventListener("keydown", (e) => {
                if (e.key === "Enter") editBtn.click();
                if (e.key === "Escape") {
                    editInput.value = item.path;
                    textEl.style.display = "";
                    editInput.style.display = "none";
                    editBtn.textContent = "✏️";
                }
            });

            row.querySelector(".sp-path-remove").addEventListener("click", () => removePath(item.path));
            pathListEl.appendChild(row);
        });
    }

    // ============================================
    // Import / Export CSV
    // ============================================
    function csvEscape(val) {
        const s = String(val == null ? "" : val);
        if (s.includes(",") || s.includes('"') || s.includes("\n")) {
            return '"' + s.replace(/"/g, '""') + '"';
        }
        return s;
    }

    function exportCSV() {
        const rows = [];
        // Header
        rows.push("type,path,checked,name,serverRelativeUrl,size,modified,folderKey");

        // PATH rows
        searchPaths.forEach(item => {
            rows.push([
                "PATH", csvEscape(item.path), item.checked ? "true" : "false",
                "", "", "", "", ""
            ].join(","));
        });

        // FILE rows
        fileData.forEach((folderEntry, folderKey) => {
            (folderEntry.files || []).forEach(file => {
                rows.push([
                    "FILE", "", "",
                    csvEscape(file.name), csvEscape(file.path),
                    file.size || 0, csvEscape(file.modified),
                    csvEscape(folderKey)
                ].join(","));
            });
        });

        const csv = rows.join("\n");
        const blob = new Blob([csv], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "sp_search_data.csv";
        a.click();
        URL.revokeObjectURL(url);

        showSettingsStatus("success", "✅ Exported " + searchPaths.length + " paths & " + countStoredFiles() + " files");
        setTimeout(() => showSettingsStatus("", ""), 3000);
    }

    function importCSV(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const text = e.target.result;
                const lines = parseCSVLines(text);
                if (lines.length < 2) throw new Error("Empty or invalid CSV");

                const header = lines[0];
                const newPaths = [];
                const newCache = new Map();

                for (let i = 1; i < lines.length; i++) {
                    const cols = lines[i];
                    if (cols.length < 3) continue;
                    const type = cols[0];

                    if (type === "PATH") {
                        const path = cols[1];
                        const checked = cols[2] === "true";
                        if (path) newPaths.push({ path, checked });
                    } else if (type === "FILE") {
                        const name = cols[3] || "";
                        const serverPath = cols[4] || "";
                        const size = parseInt(cols[5]) || 0;
                        const modified = cols[6] || "";
                        const folderKey = cols[7] || "";

                        if (!folderKey || !name) continue;
                        if (!newCache.has(folderKey)) {
                            newCache.set(folderKey, { files: [], time: Date.now() });
                        }
                        newCache.get(folderKey).files.push({ name, path: serverPath, size, modified });
                    }
                }

                // Apply imported data
                searchPaths = newPaths;
                fileData.clear();
                newCache.forEach((v, k) => fileData.set(k, v));

                savePaths();
                saveData();
                renderPathList();
                updateFilterOptions();

                showSettingsStatus("success", "✅ Imported " + newPaths.length + " paths & " + countStoredFiles() + " files");
                setTimeout(() => showSettingsStatus("", ""), 3000);
            } catch (err) {
                showSettingsStatus("error", "❌ Import failed: " + err.message);
                setTimeout(() => showSettingsStatus("", ""), 5000);
            }
        };
        reader.readAsText(file);
    }

    // Simple CSV line parser (handles quoted fields)
    function parseCSVLines(text) {
        const result = [];
        const lines = text.split("\n");
        for (const line of lines) {
            if (!line.trim()) continue;
            const cols = [];
            let current = "";
            let inQuotes = false;
            for (let i = 0; i < line.length; i++) {
                const ch = line[i];
                if (inQuotes) {
                    if (ch === '"' && line[i + 1] === '"') { current += '"'; i++; }
                    else if (ch === '"') { inQuotes = false; }
                    else { current += ch; }
                } else {
                    if (ch === '"') { inQuotes = true; }
                    else if (ch === ',') { cols.push(current); current = ""; }
                    else { current += ch; }
                }
            }
            cols.push(current);
            result.push(cols);
        }
        return result;
    }

    function countStoredFiles() {
        let count = 0;
        fileData.forEach(entry => { count += (entry.files || []).length; });
        return count;
    }

    // ============================================
    // Persistent Storage
    // ============================================
    function savePaths() { chrome.storage.local.set({ searchPaths }); }

    function saveData() {
        const dataObj = {};
        fileData.forEach((value, key) => { dataObj[key] = value; });
        chrome.storage.local.set({ fileData: dataObj });
    }

    async function restoreState() {
        try {
            const data = await chrome.storage.local.get(["searchPaths", "fileData"]);
            if (data.searchPaths && Array.isArray(data.searchPaths)) {
                // Migration: convert old string[] format to {path, checked}[]
                searchPaths = data.searchPaths.map(p => {
                    if (typeof p === "string") return { path: p, checked: true };
                    return p;
                });
                renderPathList();
                updateFilterOptions();
            }
            if (data.fileData && typeof data.fileData === "object") {
                for (const [key, value] of Object.entries(data.fileData)) {
                    fileData.set(key, value);
                }
                // Re-render to show file counts
                renderPathList();
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

            const currentQuery = searchInput.value.trim();
            if (currentQuery.length < MIN_QUERY_LENGTH) { showStatus("", ""); return; }

            const queryLower = currentQuery.toLowerCase();
            const matched = allFiles.filter(f => f.name.toLowerCase().includes(queryLower));
            lastSearchResults = matched;
            lastQuery = currentQuery;
            displayResults(applyFilter(matched), currentQuery);
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

    // ALL paths - used for search (ignores checkbox)
    function getSearchRoots() {
        if (searchPaths.length === 0) return [CONFIG.folderPath];
        return searchPaths.map(p => CONFIG.folderPath + p.path);
    }

    // Only CHECKED paths - used for refresh
    function getCheckedRoots() {
        const checked = searchPaths.filter(p => p.checked);
        if (checked.length === 0) return [CONFIG.folderPath];
        return checked.map(p => CONFIG.folderPath + p.path);
    }

    // Load only CHECKED folders (used by Refresh)
    async function loadFilesForChecked() {
        const roots = getCheckedRoots();
        const allFiles = [];

        for (const root of roots) {
            const files = await loadFilesFromFolder(root);
            fileData.set(root, { files, time: Date.now() });
            allFiles.push(...files);
        }

        saveData();
        renderPathList();
        return allFiles;
    }

    // Load ALL folders (used by Search)
    async function loadFilesForPaths() {
        const roots = getSearchRoots();
        const allFiles = [];
        const toLoad = [];

        for (const root of roots) {
            const cached = fileData.get(root);
            if (cached) {
                allFiles.push(...cached.files);
            } else {
                toLoad.push(root);
            }
        }

        if (toLoad.length === 0) return allFiles;

        for (const root of toLoad) {
            const files = await loadFilesFromFolder(root);
            fileData.set(root, { files, time: Date.now() });
            allFiles.push(...files);
        }

        saveData();
        // Update file counts in settings
        renderPathList();
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

            showSettingsStatus("loading", "Loading... (" + allFiles.length + " files, " + queue.length + " folders remaining)");

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
            const origin = CONFIG.siteUrl.split("/sites/")[0];
            const parentFolderPath = file.path.replace("/" + file.name, "");
            // Show folder path without "Shared Documents" prefix and without filename
            const displayPath = parentFolderPath.replace(CONFIG.folderPath, "").replace(/^\//, "") || "/";

            const viewFileUrl = getViewFileUrl(origin, file.path, file.name);
            const folderUrl = origin + "/sites/msteams_268a65/Shared Documents/Forms/AllItems.aspx?id=" +
                encodeURIComponent(parentFolderPath);

            const item = document.createElement("div");
            item.className = "sp-file-item";
            item.title = "Open: " + file.name;
            item.style.cursor = "pointer";
            item.addEventListener("click", () => window.open(viewFileUrl, "_blank"));

            item.innerHTML =
                '<span class="sp-file-icon">' + getFileIcon(file.name) + '</span>' +
                '<div class="sp-file-info">' +
                '<div class="sp-file-name">' + highlightText(escapeHtml(file.name), query) + '</div>' +
                '<div class="sp-file-path">📂 ' + escapeHtml(displayPath) + '</div>' +
                (file.size ? '<div class="sp-file-size">' + formatFileSize(file.size) + '</div>' : '') +
                '</div>' +
                '<div class="sp-file-actions">' +
                '<button class="sp-action-btn sp-action-folder" title="Open folder">📂</button>' +
                '</div>';

            item.querySelector(".sp-action-folder").addEventListener("click", (e) => {
                e.stopPropagation();
                window.open(folderUrl, "_blank");
            });

            resultsEl.appendChild(item);
        });
    }

    // ============================================
    // Filter Helpers
    // ============================================
    function applyFilter(files) {
        const filterValue = filterSelect.value;
        if (!filterValue) return files; // "All folders"
        const filterRoot = CONFIG.folderPath + filterValue;
        return files.filter(f => f.path.startsWith(filterRoot));
    }

    function updateFilterOptions() {
        const currentValue = filterSelect.value;
        filterSelect.innerHTML = '<option value="">All folders</option>';
        searchPaths.forEach(item => {
            const opt = document.createElement("option");
            opt.value = item.path;
            opt.textContent = item.path;
            filterSelect.appendChild(opt);
        });
        // Restore previous selection if still valid
        filterSelect.value = currentValue;
    }

    // ============================================
    // URL Helpers
    // ============================================
    function getViewFileUrl(origin, serverRelativePath, fileName) {
        return origin + serverRelativePath + "?web=1";
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

    function showSettingsStatus(type, message) {
        if (!type && !message) { settingsStatusEl.className = "sp-status sp-hidden"; return; }
        settingsStatusEl.className = "sp-status sp-" + type;
        settingsStatusEl.innerHTML = type === "loading"
            ? '<span class="sp-spinner"></span> ' + message
            : message;
        if (!message) settingsStatusEl.classList.add("sp-hidden");
        else settingsStatusEl.classList.remove("sp-hidden");
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
        var i = 0, size = bytes;
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
