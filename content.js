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
    let CONFIG = {
        siteUrl: "https://dxcportal.sharepoint.com/sites/msteams_268a65",
        folderPath: "/sites/msteams_268a65/Shared Documents",
        fullUrl: "https://dxcportal.sharepoint.com/sites/msteams_268a65/Shared Documents"
    };

    const BATCH_SIZE = 3;
    const DELAY_MS = 300;
    const MIN_QUERY_LENGTH = 3;
    const DEBOUNCE_MS = 300;
    const CURRENT_FILE_EXTENSIONS = new Set(["doc", "docx", "xls", "xlsx", "xlsm", "xlsb"]);

    // searchPaths: array of { path: string, checked: boolean }
    let searchPaths = [];
    const fileData = new Map();
    let debounceTimer = null;
    let isLoading = false;
    let lastSearchResults = []; // store last search results for re-filtering
    let lastQuery = "";
    let currentPageFile = null;

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

            <div class="sp-current-file sp-hidden" id="sp-current-file">
                <div class="sp-current-file-info">
                    <div class="sp-current-file-label">Current file</div>
                    <div class="sp-current-file-name" id="sp-current-file-name"></div>
                </div>
                <div class="sp-file-actions">
                    <button class="sp-action-btn sp-action-checkout" id="sp-current-file-checkout-btn" title="Check out file">📝</button>
                    <button class="sp-action-btn sp-action-folder" id="sp-current-file-folder-btn" title="Open folder">📂</button>
                </div>
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
                        <div class="sp-settings-label">Target Document Library URL</div>
                        <div class="sp-path-add-row">
                            <input type="text" class="sp-path-input" id="sp-config-url-input" placeholder="https://..." autocomplete="off">
                            <button class="sp-add-path-btn" id="sp-save-config-btn">Save</button>
                        </div>
                        <div class="sp-refresh-hint">Changing this will clear your subfolders and cached data</div>
                    </div>
                    <div class="sp-settings-section">
                        <div class="sp-settings-label">Add Subfolder</div>
                        <div class="sp-path-add-row">
                            <input type="text" class="sp-path-input" id="sp-path-input" placeholder="/subfolder/path or SharePoint folder URL" autocomplete="off">
                            <button class="sp-add-path-btn" id="sp-add-path-btn">+ Add</button>
                            <button class="sp-add-path-btn" id="sp-current-page-btn">Current Page</button>
                        </div>
                    </div>
                    <div class="sp-settings-section">
                        <div class="sp-settings-label">Subfolders</div>
                        <div class="sp-path-list" id="sp-path-list">
                            <div class="sp-path-list-empty">No subfolders added. Search uses cached files only until you add a subfolder or run a manual refresh.</div>
                        </div>
                    </div>
                    <div class="sp-refresh-section">
                        <div class="sp-status sp-hidden" id="sp-settings-status"></div>
                        <div class="sp-refresh-actions">
                            <button class="sp-refresh-full-btn" id="sp-refresh-btn">🔄 Re-fetch checked folders</button>
                            <button class="sp-refresh-stop-btn sp-hidden" id="sp-stop-refresh-btn">⏹ Stop</button>
                        </div>
                        <div class="sp-refresh-hint">Re-fetch file list from checked subfolders</div>
                    </div>
                </div>
            </div>

            <div class="sp-footer" style="display:flex;">
                <span id="sp-connection-status">⏳ Checking connection...</span>
                <span id="sp-sync-status" style="margin-left:auto;color:#888;font-size:12px;"></span>
            </div>
        </div>
        <div class="sp-inline-modal-backdrop sp-hidden" id="sp-inline-modal-backdrop">
            <div class="sp-inline-modal" id="sp-inline-modal" role="dialog" aria-modal="true" aria-labelledby="sp-inline-modal-title">
                <div class="sp-inline-modal-title" id="sp-inline-modal-title">Check In Comment</div>
                <div class="sp-inline-modal-subtitle" id="sp-inline-modal-subtitle"></div>
                <textarea class="sp-inline-modal-textarea" id="sp-inline-modal-textarea" placeholder="Add a check-in comment"></textarea>
                <div class="sp-inline-modal-actions">
                    <button class="sp-inline-modal-btn sp-inline-modal-btn-secondary" id="sp-inline-modal-cancel-btn">Cancel</button>
                    <button class="sp-inline-modal-btn sp-inline-modal-btn-primary" id="sp-inline-modal-ok-btn">OK</button>
                </div>
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
    const stopRefreshBtn = document.getElementById("sp-stop-refresh-btn");
    const pathInput = document.getElementById("sp-path-input");
    const addPathBtn = document.getElementById("sp-add-path-btn");
    const currentPageBtn = document.getElementById("sp-current-page-btn");
    const pathListEl = document.getElementById("sp-path-list");
    const statusEl = document.getElementById("sp-status");
    const resultsEl = document.getElementById("sp-results");
    const closeBtn = document.getElementById("sp-close-btn");
    const connectionEl = document.getElementById("sp-connection-status");
    const syncStatusEl = document.getElementById("sp-sync-status");
    const settingsStatusEl = document.getElementById("sp-settings-status");
    const currentFileEl = document.getElementById("sp-current-file");
    const currentFileNameEl = document.getElementById("sp-current-file-name");
    const currentFileCheckoutBtn = document.getElementById("sp-current-file-checkout-btn");
    const currentFileFolderBtn = document.getElementById("sp-current-file-folder-btn");
    const inlineModalBackdrop = document.getElementById("sp-inline-modal-backdrop");
    const inlineModal = document.getElementById("sp-inline-modal");
    const inlineModalSubtitleEl = document.getElementById("sp-inline-modal-subtitle");
    const inlineModalTextarea = document.getElementById("sp-inline-modal-textarea");
    const inlineModalCancelBtn = document.getElementById("sp-inline-modal-cancel-btn");
    const inlineModalOkBtn = document.getElementById("sp-inline-modal-ok-btn");
    const DEFAULT_REFRESH_LABEL = "🔄 Re-fetch checked folders";
    const FETCH_STOPPED_ERROR = "FETCH_STOPPED";
    const CHECKOUT_ICON = "📝";
    const CHECKIN_ICON = "✅";
    let activeRefreshController = null;
    let activeInlineModal = null;

    function updateSyncStatusLabel(timeStr, changesCount) {
        if (!syncStatusEl) return;
        if (!timeStr) {
            syncStatusEl.textContent = "";
            return;
        }
        const displayTime = new Date(timeStr).toLocaleTimeString();
        let text = `Auto-synced: ${displayTime}`;
        if (changesCount !== undefined && changesCount > 0) {
             text += ` (${changesCount} changes)`;
        } else if (changesCount === 0) {
             text += ` (0 changes)`;
        }
        syncStatusEl.textContent = text;
    }
    const filterSelect = document.getElementById("sp-filter-select");
    const exportBtn = document.getElementById("sp-export-btn");
    const importBtn = document.getElementById("sp-import-btn");
    const importFile = document.getElementById("sp-import-file");
    const configUrlInput = document.getElementById("sp-config-url-input");
    const saveConfigBtn = document.getElementById("sp-save-config-btn");
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
        if (e.key === "Escape" && activeInlineModal) {
            closeInlineModal(false);
            return;
        }
        if (e.key === "Escape" && document.getElementById("sp-search-backdrop")) closeModal();
    });
    modal.addEventListener("click", (e) => e.stopPropagation());
    inlineModal.addEventListener("click", (e) => e.stopPropagation());
    inlineModalBackdrop.addEventListener("click", () => closeInlineModal(false));
    inlineModalCancelBtn.addEventListener("click", () => closeInlineModal(false));
    inlineModalOkBtn.addEventListener("click", () => submitInlineModal());
    inlineModalTextarea.addEventListener("keydown", (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
            e.preventDefault();
            submitInlineModal();
        }
    });
    currentFileCheckoutBtn.addEventListener("click", async (e) => {
        e.stopPropagation();
        if (!currentPageFile) return;
        await toggleFileCheckout(currentPageFile, currentFileCheckoutBtn);
        renderCurrentPageFile();
    });
    currentFileFolderBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        if (!currentPageFile) return;
        window.open(getFolderUrl(currentPageFile.path), "_blank");
    });

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
        if (activeRefreshController) return;

        activeRefreshController = { cancelled: false };
        refreshBtn.disabled = true;
        refreshBtn.textContent = "⏳ Re-fetching...";
        stopRefreshBtn.classList.remove("sp-hidden");
        stopRefreshBtn.disabled = false;
        try {
            const files = await loadFilesForChecked(activeRefreshController);
            showSettingsStatus("", "");
            refreshBtn.textContent = "✅ Re-fetched! (" + files.length + " files)";
            setTimeout(() => {
                resetRefreshControls();
            }, 2000);
        } catch (e) {
            if (e.message === FETCH_STOPPED_ERROR) {
                showSettingsStatus("error", "⏹ Re-fetch stopped.");
                refreshBtn.textContent = "⏹ Stopped";
                setTimeout(() => {
                    showSettingsStatus("", "");
                    resetRefreshControls();
                }, 1500);
            } else {
                showSettingsStatus("error", "❌ Failed: " + e.message);
                refreshBtn.textContent = "❌ Failed";
                setTimeout(() => {
                    showSettingsStatus("", "");
                    resetRefreshControls();
                }, 3000);
            }
        } finally {
            activeRefreshController = null;
        }
    });

    stopRefreshBtn.addEventListener("click", () => {
        if (!activeRefreshController) return;
        activeRefreshController.cancelled = true;
        stopRefreshBtn.disabled = true;
        stopRefreshBtn.textContent = "⏳ Stopping...";
        showSettingsStatus("loading", "Stopping after the current request finishes...");
    });

    // Path management
    addPathBtn.addEventListener("click", addPath);
    currentPageBtn.addEventListener("click", useCurrentPageUrl);
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

    // Configuration
    saveConfigBtn.addEventListener("click", () => {
        const url = configUrlInput.value.trim();
        if (!url) return;
        if (!confirm("Changing the Document Library URL will reset your subfolders and cached data. Continue?")) return;
        parseAndApplyConfigUrl(url, true);
    });

    // Focus input & restore state
    searchInput.focus();
    restoreState();
    checkConnection();
    detectCurrentPageFile();

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
    function parseAndApplyConfigUrl(inputUrl, resetData) {
        try {
            const url = new URL(inputUrl);
            const origin = url.origin;
            const pathname = decodeURIComponent(url.pathname);
            
            const segments = pathname.split('/').filter(Boolean);
            if (segments.length < 2) throw new Error("Invalid URL form");
            
            CONFIG.siteUrl = origin + '/' + segments[0] + '/' + segments[1];
            CONFIG.folderPath = pathname;
            CONFIG.fullUrl = inputUrl;

            if (resetData) {
                searchPaths = [];
                fileData.clear();
                savePaths();
                saveData();
                chrome.storage.local.set({ configUrl: inputUrl });
                renderPathList();
                updateFilterOptions();
                showSettingsStatus("success", "✅ Configuration saved and data reset.");
                setTimeout(() => showSettingsStatus("", ""), 3000);
            }
        } catch (e) {
            if (resetData) {
                showSettingsStatus("error", "❌ Invalid SharePoint URL format.");
                setTimeout(() => showSettingsStatus("", ""), 3000);
            }
        }
    }

    function addPath() {
        let path;
        try {
            path = normalizeSubfolderInput(pathInput.value);
        } catch (e) {
            showSettingsStatus("error", "❌ " + e.message);
            setTimeout(() => showSettingsStatus("", ""), 3000);
            return;
        }

        if (!path) return;
        if (searchPaths.some(p => p.path === path)) { pathInput.value = ""; return; }

        searchPaths.push({ path, checked: true });
        pathInput.value = "";
        renderPathList();
        savePaths();
        updateFilterOptions();
    }

    function useCurrentPageUrl() {
        pathInput.value = window.location.href;
        pathInput.focus();
    }

    function normalizeSubfolderInput(rawInput) {
        const value = (rawInput || "").trim();
        if (!value) return "";

        let path = value;

        if (/^https?:\/\//i.test(value)) {
            path = extractSubfolderFromUrl(value);
        }

        if (!path.startsWith("/")) path = "/" + path;
        if (path.length > 1 && path.endsWith("/")) path = path.slice(0, -1);
        return path;
    }

    function extractSubfolderFromUrl(inputUrl) {
        let url;
        try {
            url = new URL(inputUrl);
        } catch (e) {
            throw new Error("Invalid subfolder URL.");
        }

        let folderPath = url.searchParams.get("id") || decodeURIComponent(url.pathname);
        folderPath = decodeURIComponent(folderPath);

        if (!folderPath.startsWith(CONFIG.folderPath)) {
            throw new Error("URL is outside the current document library.");
        }

        let relativePath = folderPath.slice(CONFIG.folderPath.length);
        if (!relativePath) {
            throw new Error("The URL points to the document library root, not a subfolder.");
        }

        if (!relativePath.startsWith("/")) {
            relativePath = "/" + relativePath;
        }

        return relativePath;
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
            pathListEl.innerHTML = '<div class="sp-path-list-empty">No subfolders added. Search uses cached files only until you add a subfolder or run a manual refresh.</div>';
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
        rows.push("type,path,checked,name,serverRelativeUrl,size,modified,folderKey,isCheckedOut");

        // URL row
        rows.push([
            "URL", csvEscape(CONFIG.fullUrl), "",
            "", "", "", "", ""
        ].join(","));

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
                    csvEscape(folderKey), file.isCheckedOut ? "true" : "false"
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

                    if (type === "URL") {
                        const importedUrl = cols[1];
                        if (importedUrl) {
                            parseAndApplyConfigUrl(importedUrl, false);
                            configUrlInput.value = importedUrl;
                            chrome.storage.local.set({ configUrl: importedUrl });
                        }
                    } else if (type === "PATH") {
                        const path = cols[1];
                        const checked = cols[2] === "true";
                        if (path) newPaths.push({ path, checked });
                    } else if (type === "FILE") {
                        const name = cols[3] || "";
                        const serverPath = cols[4] || "";
                        const size = parseInt(cols[5]) || 0;
                        const modified = cols[6] || "";
                        const folderKey = cols[7] || "";
                        const isCheckedOut = cols[8] === "true";

                        if (!folderKey || !name) continue;
                        if (!newCache.has(folderKey)) {
                            newCache.set(folderKey, { files: [], time: Date.now() });
                        }
                        newCache.get(folderKey).files.push({ name, path: serverPath, size, modified, isCheckedOut });
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

    function resetRefreshControls() {
        refreshBtn.textContent = DEFAULT_REFRESH_LABEL;
        refreshBtn.disabled = false;
        stopRefreshBtn.textContent = "⏹ Stop";
        stopRefreshBtn.disabled = false;
        stopRefreshBtn.classList.add("sp-hidden");
    }

    // ============================================
    // Persistent Storage
    // ============================================
    function savePaths() { chrome.storage.local.set({ searchPaths }); }

    function buildFileDataObject() {
        const dataObj = {};
        fileData.forEach((value, key) => { dataObj[key] = value; });
        return dataObj;
    }

    function saveFileDataOnly() {
        chrome.storage.local.set({ fileData: buildFileDataObject() });
    }

    function saveData(changesCount = 0) {
        const dataObj = buildFileDataObject();
        const timeNow = new Date().toISOString();
        chrome.storage.local.set({ 
            fileData: dataObj,
            lastSyncTime: timeNow,
            lastSyncChanges: changesCount
        });
        updateSyncStatusLabel(timeNow, changesCount);
    }

    // ============================================
    // Listen for Background Updates
    // ============================================
    chrome.storage.onChanged.addListener((changes, namespace) => {
        if (namespace === "local" && changes.fileData) {
            const newData = changes.fileData.newValue;
            if (newData) {
                // Update local memory cache
                for (const [key, value] of Object.entries(newData)) {
                    fileData.set(key, value);
                }
                
                // Re-render path list to update file counts
                renderPathList();

                // If user is actively searching, re-run search to show new files
                if (searchInput.value.trim().length >= MIN_QUERY_LENGTH) {
                    clearTimeout(debounceTimer);
                    debounceTimer = setTimeout(() => handleSearch(), 300);
                }
            }
        }
        if (namespace === "local" && changes.lastSyncTime) {
            const time = changes.lastSyncTime.newValue;
            // if changes object contains lastSyncChanges, use it. Otherwise, assume 0 for manual sync
            const changesCount = changes.lastSyncChanges ? changes.lastSyncChanges.newValue : 0;
            
            updateSyncStatusLabel(time, changesCount);

            if (settingsStatusEl) {
                showSettingsStatus("success", `🔄 Auto-synced at ${new Date(time).toLocaleTimeString()} (${changesCount} updated)`);
                setTimeout(() => showSettingsStatus("", ""), 4000);
            }
        }
    });

    async function restoreState() {
        try {
            const data = await chrome.storage.local.get(["searchPaths", "fileData", "configUrl", "lastSyncTime", "lastSyncChanges"]);
            
            if (data.lastSyncTime) {
                updateSyncStatusLabel(data.lastSyncTime, data.lastSyncChanges);
            }
            
            if (data.configUrl) {
                parseAndApplyConfigUrl(data.configUrl, false);
                configUrlInput.value = data.configUrl;
            } else {
                configUrlInput.value = CONFIG.fullUrl;
            }

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

    async function detectCurrentPageFile() {
        const file = getCurrentPageFileCandidate();
        if (!file) {
            currentPageFile = null;
            renderCurrentPageFile();
            return;
        }

        try {
            const fileUrl = getCurrentPageFileDetailsUrl(file);
            if (fileUrl) {
                const fileDataResponse = await bgFetch(fileUrl, {
                    method: "GET",
                    headers: { "Accept": "application/json;odata=verbose" }
                });
                const details = fileDataResponse?.d;
                if (details) {
                    file.name = details.Name || file.name;
                    file.path = details.ServerRelativeUrl || file.path;
                    file.size = parseInt(details.Length || 0, 10);
                    file.modified = details.TimeLastModified || "";
                    file.isCheckedOut = isFileCheckedOut(details.CheckOutType);
                }
            }
        } catch (e) {
            console.warn("Could not load current file details:", e);
        }

        currentPageFile = file;
        renderCurrentPageFile();
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
            showStatus("loading", "Loading cached file list...");
            const allFiles = getAllCachedFiles();

            if (allFiles.length === 0) {
                showStatus("error", "ℹ️ No cached files found. Use Re-fetch in Settings first.");
                lastSearchResults = [];
                lastQuery = query;
                resultsEl.innerHTML = "";
                return;
            }

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

    function throwIfRefreshCancelled(controller) {
        if (controller && controller.cancelled) {
            throw new Error(FETCH_STOPPED_ERROR);
        }
    }

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
    async function loadFilesForChecked(controller) {
        const roots = getCheckedRoots();
        const allFiles = [];

        for (const root of roots) {
            throwIfRefreshCancelled(controller);
            const files = await loadFilesFromFolder(root, controller);
            throwIfRefreshCancelled(controller);
            fileData.set(root, { files, time: Date.now() });
            allFiles.push(...files);
        }

        saveData();
        renderPathList();
        return allFiles;
    }

    function getAllCachedFiles() {
        const allFiles = [];
        fileData.forEach(entry => {
            if (entry && Array.isArray(entry.files)) {
                allFiles.push(...entry.files);
            }
        });
        return allFiles;
    }

    async function loadFilesFromFolder(rootPath, controller) {
        const allFiles = [];
        const queue = [rootPath];

        while (queue.length > 0) {
            throwIfRefreshCancelled(controller);
            const batch = queue.splice(0, BATCH_SIZE);
            const batchResults = await Promise.all(
                batch.map(fp => fetchFolderContents(fp))
            );

            for (const result of batchResults) {
                allFiles.push(...result.files);
                queue.push(...result.subfolders);
            }

            showSettingsStatus("loading", "Loading... (" + allFiles.length + " files, " + queue.length + " folders remaining)");

            if (queue.length > 0) {
                await delayMs(DELAY_MS);
                throwIfRefreshCancelled(controller);
            }
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
                "')/Files?$top=5000&$select=Name,ServerRelativeUrl,Length,TimeLastModified,CheckOutType";

            const filesData = await bgFetch(filesUrl, {
                method: "GET",
                headers: { "Accept": "application/json;odata=verbose" }
            });

            for (const file of (filesData.d?.results || [])) {
                files.push({
                    name: file.Name,
                    path: file.ServerRelativeUrl,
                    size: parseInt(file.Length || 0),
                    modified: file.TimeLastModified || "",
                    isCheckedOut: isFileCheckedOut(file.CheckOutType)
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
            const origin = getOriginFromServerRelativePath(file.path) || CONFIG.siteUrl.split("/sites/")[0];
            const displayPath = getDisplayFolderPath(file.path, file.name);
            const viewFileUrl = getViewFileUrl(origin, file.path);
            const folderUrl = getFolderUrl(file.path);

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
                '<button class="sp-action-btn sp-action-checkout" title="' + (file.isCheckedOut ? 'Check in file' : 'Check out file') + '">' + getCheckoutActionIcon(file.isCheckedOut) + '</button>' +
                '<button class="sp-action-btn sp-action-folder" title="Open folder">📂</button>' +
                '</div>';

            item.querySelector(".sp-action-checkout").addEventListener("click", async (e) => {
                e.stopPropagation();
                await toggleFileCheckout(file, item.querySelector(".sp-action-checkout"));
            });

            item.querySelector(".sp-action-folder").addEventListener("click", (e) => {
                e.stopPropagation();
                window.open(folderUrl, "_blank");
            });

            resultsEl.appendChild(item);
        });
    }

    async function toggleFileCheckout(file, buttonEl) {
        if (!buttonEl || buttonEl.disabled) return;

        const nextAction = file.isCheckedOut ? "checkIn" : "checkOut";
        const originalLabel = buttonEl.textContent;
        buttonEl.disabled = true;
        buttonEl.textContent = "⏳";

        try {
            if (nextAction === "checkOut") {
                await sharePointCheckOut(file.path);
                file.isCheckedOut = true;
                buttonEl.textContent = getCheckoutActionIcon(true);
                buttonEl.title = "Check in file";
                showStatus("success", '✅ Checked out "' + escapeHtml(file.name) + '"');
            } else {
                const comment = await promptForCheckInComment(file.name);
                if (comment === null) {
                    buttonEl.textContent = originalLabel;
                    return;
                }
                await sharePointCheckIn(file.path, comment);
                file.isCheckedOut = false;
                buttonEl.textContent = getCheckoutActionIcon(false);
                buttonEl.title = "Check out file";
                showStatus("success", '✅ Checked in "' + escapeHtml(file.name) + '"');
            }
            saveFileDataOnly();
            if (currentPageFile && currentPageFile.path === file.path) {
                currentPageFile.isCheckedOut = file.isCheckedOut;
                renderCurrentPageFile();
            }
        } catch (error) {
            buttonEl.textContent = originalLabel;
            buttonEl.title = file.isCheckedOut ? "Check in file" : "Check out file";
            showStatus("error", "❌ " + error.message);
        } finally {
            buttonEl.disabled = false;
        }
    }

    async function sharePointCheckOut(serverRelativePath) {
        const url = CONFIG.siteUrl +
            "/_api/web/getfilebyserverrelativeurl('" +
            encodeSharePointPath(serverRelativePath) +
            "')/checkOut()";

        await bgFetch(url, {
            method: "POST",
            headers: { "Accept": "application/json;odata=verbose" }
        });
    }

    async function sharePointCheckIn(serverRelativePath, comment) {
        const url = CONFIG.siteUrl +
            "/_api/web/getfilebyserverrelativeurl('" +
            encodeSharePointPath(serverRelativePath) +
            "')/checkIn(comment='" +
            encodeODataString(comment || "") +
            "',checkintype=0)";

        await bgFetch(url, {
            method: "POST",
            headers: { "Accept": "application/json;odata=verbose" }
        });
    }

    function promptForCheckInComment(fileName) {
        if (activeInlineModal) {
            closeInlineModal(false);
        }

        inlineModalSubtitleEl.textContent = fileName ? 'Enter a check-in comment for "' + fileName + '"' : "Enter a check-in comment";
        inlineModalTextarea.value = "";
        inlineModalBackdrop.classList.remove("sp-hidden");
        inlineModalOkBtn.disabled = false;
        activeInlineModal = {};

        setTimeout(() => {
            inlineModalTextarea.focus();
            inlineModalTextarea.select();
        }, 0);

        return new Promise((resolve) => {
            activeInlineModal.resolve = resolve;
        });
    }

    function submitInlineModal() {
        if (!activeInlineModal) return;
        closeInlineModal(true, inlineModalTextarea.value.trim());
    }

    function closeInlineModal(submitted, value) {
        if (!activeInlineModal) return;
        const resolver = activeInlineModal.resolve;
        activeInlineModal = null;
        inlineModalBackdrop.classList.add("sp-hidden");
        inlineModalTextarea.value = "";
        if (resolver) {
            resolver(submitted ? value : null);
        }
    }

    function renderCurrentPageFile() {
        if (!currentFileEl) {
            return;
        }
        if (!currentPageFile) {
            currentFileEl.classList.add("sp-hidden");
            return;
        }

        currentFileNameEl.textContent = currentPageFile.name;
        currentFileCheckoutBtn.textContent = getCheckoutActionIcon(currentPageFile.isCheckedOut);
        currentFileCheckoutBtn.title = currentPageFile.isCheckedOut ? "Check in file" : "Check out file";
        currentFileEl.classList.remove("sp-hidden");
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
    function getViewFileUrl(origin, serverRelativePath) {
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

    function encodeSharePointPath(path) {
        return encodeURIComponent(String(path || "").replace(/'/g, "''"));
    }

    function encodeODataString(value) {
        return encodeURIComponent(String(value || "").replace(/'/g, "''"));
    }

    function getCheckoutActionIcon(isCheckedOut) {
        return isCheckedOut ? CHECKIN_ICON : CHECKOUT_ICON;
    }

    function isFileCheckedOut(checkOutType) {
        const value = parseInt(checkOutType, 10);
        if (Number.isNaN(value)) return false;
        return value !== 2;
    }

    function getCurrentPageFileCandidate() {
        const currentUrl = new URL(window.location.href);
        const rawIdPath = currentUrl.searchParams.get("id");
        const rawFileName = currentUrl.searchParams.get("file");
        const rawSourceDoc = currentUrl.searchParams.get("sourcedoc");
        const serverRelativePath = decodeURIComponent(rawIdPath || "");
        const fileName = decodeSharePointFileName(rawFileName) || (serverRelativePath.split("/").pop() || "");
        const sourceDocId = normalizeSourceDocId(rawSourceDoc);
        const siteUrl = getCurrentPageSiteUrl(currentUrl, serverRelativePath);
        if (!fileName || !isCurrentFileTypeSupported(fileName)) {
            return null;
        }
        return {
            name: fileName,
            path: serverRelativePath,
            size: 0,
            modified: "",
            isCheckedOut: false,
            sourceDocId,
            siteUrl
        };
    }

    function isCurrentFileTypeSupported(fileName) {
        const ext = String(fileName || "").split(".").pop().toLowerCase();
        return CURRENT_FILE_EXTENSIONS.has(ext);
    }

    function getOriginFromServerRelativePath(serverRelativePath) {
        if (!serverRelativePath || !serverRelativePath.startsWith("/")) return "";
        return window.location.origin;
    }

    function getSiteUrlFromServerRelativePath(serverRelativePath) {
        if (!serverRelativePath) return "";
        const segments = String(serverRelativePath).split("/").filter(Boolean);
        if (segments.length < 2) return "";
        return window.location.origin + "/" + segments[0] + "/" + segments[1];
    }

    function getFolderPath(serverRelativePath, fileName) {
        return String(serverRelativePath || "").replace("/" + (fileName || ""), "");
    }

    function getDisplayFolderPath(serverRelativePath, fileName) {
        const parentFolderPath = getFolderPath(serverRelativePath, fileName);
        return parentFolderPath.replace(CONFIG.folderPath, "").replace(/^\//, "") || "/";
    }

    function getFolderUrl(serverRelativePath) {
        const parentFolderPath = getFolderPath(serverRelativePath, String(serverRelativePath || "").split("/").pop() || "");
        return window.location.origin + CONFIG.folderPath + "/Forms/AllItems.aspx?id=" + encodeURIComponent(parentFolderPath);
    }

    function getCurrentPageFileDetailsUrl(file) {
        const select = "?$select=Name,ServerRelativeUrl,Length,TimeLastModified,CheckOutType";
        if (file.sourceDocId && file.siteUrl) {
            return file.siteUrl + "/_api/web/GetFileById(guid'" + encodeODataString(file.sourceDocId) + "')" + select;
        }
        if (file.path) {
            const siteUrl = file.siteUrl || getSiteUrlFromServerRelativePath(file.path);
            if (!siteUrl) return "";
            return siteUrl + "/_api/web/getfilebyserverrelativeurl('" + encodeSharePointPath(file.path) + "')" + select;
        }
        return "";
    }

    function decodeSharePointFileName(rawFileName) {
        if (!rawFileName) return "";
        let value = rawFileName;
        try {
            value = decodeURIComponent(rawFileName);
        } catch (e) {
            value = rawFileName;
        }
        return value.replace(/%u([0-9a-fA-F]{4})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
    }

    function normalizeSourceDocId(rawSourceDoc) {
        if (!rawSourceDoc) return "";
        let value = rawSourceDoc;
        try {
            value = decodeURIComponent(rawSourceDoc);
        } catch (e) {
            value = rawSourceDoc;
        }
        return value.replace(/[{}]/g, "");
    }

    function getCurrentPageSiteUrl(currentUrl, serverRelativePath) {
        const fromPath = getSiteUrlFromServerRelativePath(serverRelativePath);
        if (fromPath) return fromPath;
        const match = currentUrl.pathname.match(/\/(sites|teams)\/[^/]+/);
        if (!match) return "";
        return currentUrl.origin + match[0];
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
