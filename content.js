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
    const TENANT_ORIGIN = "https://dxcportal.sharepoint.com";
    const DEFAULT_SITE = "msteams_268a65";

    const MIN_QUERY_LENGTH = 3;
    const CURRENT_FILE_EXTENSIONS = new Set(["doc", "docx", "xls", "xlsx", "xlsm", "xlsb", "ppt", "pptx", "pptm"]);

    // siteConfigs: array of { site: string, folders: string[] }
    let siteConfigs = [];
    let defaultSite = "";
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
                    <div class="sp-current-file-name" id="sp-current-file-name"></div>
                    <div class="sp-file-path" id="sp-current-file-path"></div>
                </div>
                <div class="sp-file-actions">
                    <button class="sp-action-btn sp-action-desktop" id="sp-current-file-desktop-btn" title="Open in desktop app">🖥️</button>
                    <button class="sp-action-btn sp-action-checkout" id="sp-current-file-checkout-btn" title="Check out file">📝</button>
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
                    </div>
                    <div class="sp-search-filters-row">
                        <div class="sp-filter-group">
                            <label class="sp-filter-label" for="sp-site-filter-select">Site</label>
                            <select class="sp-filter-select" id="sp-site-filter-select" title="Filter by site">
                                <option value="">All sites</option>
                            </select>
                        </div>
                        <div class="sp-filter-group">
                            <label class="sp-filter-label" for="sp-folder-filter-select">Folder</label>
                            <select class="sp-filter-select" id="sp-folder-filter-select" title="Filter by subfolder">
                                <option value="">All folders</option>
                            </select>
                        </div>
                        <div class="sp-filter-group">
                            <label class="sp-filter-label" for="sp-type-filter-select">Type</label>
                            <select class="sp-filter-select" id="sp-type-filter-select" title="Filter by file type">
                                <option value="">All types</option>
                                <option value="excel">Excel</option>
                                <option value="word">Word</option>
                                <option value="powerpoint">PowerPoint</option>
                                <option value="pdf">PDF</option>
                                <option value="other">Other</option>
                            </select>
                        </div>
                    </div>
                    <div class="sp-search-hint">Type at least 3 characters to search</div>
                </div>
                <div class="sp-status sp-hidden" id="sp-status"></div>
                <div class="sp-results" id="sp-results"></div>
            </div>

            <!-- Settings Tab -->
            <div class="sp-tab-panel" id="sp-tab-settings">
                <div class="sp-settings-content">
                    <div class="sp-settings-section">
                        <div class="sp-settings-label">Add Site</div>
                        <div class="sp-path-add-row">
                            <input type="text" class="sp-path-input" id="sp-site-input" placeholder="site-name" autocomplete="off">
                            <button class="sp-add-path-btn" id="sp-add-site-btn">+ Add</button>
                            <button class="sp-add-path-btn" id="sp-current-site-btn">Current Site</button>
                        </div>
                        <div class="sp-refresh-hint">Library URL format: ${TENANT_ORIGIN}/sites/{site}/Shared Documents</div>
                    </div>
                    <div class="sp-settings-section">
                        <div class="sp-settings-label">Sites</div>
                        <div class="sp-path-list" id="sp-site-list">
                            <div class="sp-path-list-empty">No sites added yet.</div>
                        </div>
                    </div>
                    <div class="sp-refresh-hint" id="sp-selected-library-hint"></div>
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
                            <div class="sp-path-list-empty">No subfolders added yet.</div>
                        </div>
                    </div>
                    <div class="sp-status sp-hidden" id="sp-settings-status"></div>
                </div>
            </div>

            <div class="sp-footer" style="display:flex;">
                <span id="sp-connection-status">⏳ Checking connection...</span>
            </div>
        </div>
        <div class="sp-inline-modal-backdrop sp-hidden" id="sp-inline-modal-backdrop">
            <div class="sp-inline-modal" id="sp-inline-modal" role="dialog" aria-modal="true" aria-labelledby="sp-inline-modal-title">
                <div class="sp-inline-modal-title" id="sp-inline-modal-title">Check In Comment</div>
                <div class="sp-inline-modal-subtitle" id="sp-inline-modal-subtitle"></div>
                <div class="sp-inline-modal-warning" id="sp-inline-modal-warning">
                    ⚠️ Make sure Excel shows <strong>"Saved"</strong> before checking in — checking in too early may cause unsaved changes to be lost.
                </div>
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
    const siteInput = document.getElementById("sp-site-input");
    const addSiteBtn = document.getElementById("sp-add-site-btn");
    const currentSiteBtn = document.getElementById("sp-current-site-btn");
    const siteListEl = document.getElementById("sp-site-list");
    const selectedLibraryHintEl = document.getElementById("sp-selected-library-hint");
    const pathInput = document.getElementById("sp-path-input");
    const addPathBtn = document.getElementById("sp-add-path-btn");
    const currentPageBtn = document.getElementById("sp-current-page-btn");
    const pathListEl = document.getElementById("sp-path-list");
    const statusEl = document.getElementById("sp-status");
    const resultsEl = document.getElementById("sp-results");
    const closeBtn = document.getElementById("sp-close-btn");
    const connectionEl = document.getElementById("sp-connection-status");
    const settingsStatusEl = document.getElementById("sp-settings-status");
    const currentFileEl = document.getElementById("sp-current-file");
    const currentFileNameEl = document.getElementById("sp-current-file-name");
    const currentFilePathEl = document.getElementById("sp-current-file-path");
    const currentFileDesktopBtn = document.getElementById("sp-current-file-desktop-btn");
    const currentFileCheckoutBtn = document.getElementById("sp-current-file-checkout-btn");
    const inlineModalBackdrop = document.getElementById("sp-inline-modal-backdrop");
    const inlineModal = document.getElementById("sp-inline-modal");
    const inlineModalSubtitleEl = document.getElementById("sp-inline-modal-subtitle");
    const inlineModalTextarea = document.getElementById("sp-inline-modal-textarea");
    const inlineModalCancelBtn = document.getElementById("sp-inline-modal-cancel-btn");
    const inlineModalOkBtn = document.getElementById("sp-inline-modal-ok-btn");
    const CHECKOUT_ICON = "📝";
    const CHECKIN_ICON = "✅";
    let activeInlineModal = null;

    const siteFilterSelect = document.getElementById("sp-site-filter-select");
    const folderFilterSelect = document.getElementById("sp-folder-filter-select");
    const fileTypeFilterSelect = document.getElementById("sp-type-filter-select");
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
    currentFileDesktopBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        if (!currentPageFile) return;
        openFileInDesktopApp(currentPageFile);
    });

    // Search input - press Enter to search
    searchInput.addEventListener("input", () => {
        const val = searchInput.value;
        clearBtn.classList.toggle("sp-visible", val.length > 0);
        if (val.trim().length === 0) {
            resultsEl.innerHTML = "";
            showStatus("", "");
            lastSearchResults = [];
            lastQuery = "";
        }
    });

    searchInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            handleSearch();
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

    // Filter dropdowns — re-filter existing results instantly
    siteFilterSelect.addEventListener("change", () => {
        updateFolderFilterOptions();
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(applyFilter(lastSearchResults), lastQuery);
        }
    });

    folderFilterSelect.addEventListener("change", () => {
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(applyFilter(lastSearchResults), lastQuery);
        }
    });

    fileTypeFilterSelect.addEventListener("change", () => {
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(applyFilter(lastSearchResults), lastQuery);
        }
    });

    // Site management
    addSiteBtn.addEventListener("click", addSite);
    currentSiteBtn.addEventListener("click", useCurrentSiteUrl);
    siteInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") addSite();
    });

    // Folder management
    addPathBtn.addEventListener("click", addPath);
    currentPageBtn.addEventListener("click", useCurrentPageUrl);
    pathInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") addPath();
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
    // Site & Path Management
    // ============================================
    function normalizeSiteInput(rawInput) {
        const value = String(rawInput || "").trim();
        if (!value) return "";

        if (/^https?:\/\//i.test(value)) {
            try {
                const url = new URL(value);
                const match = decodeURIComponent(url.pathname).match(/^\/sites\/([^/]+)/i);
                return match ? match[1] : "";
            } catch (e) {
                return "";
            }
        }

        return value
            .replace(/^\/+/, "")
            .replace(/^sites\//i, "")
            .replace(/\/.*$/, "")
            .trim();
    }

    function buildSiteUrl(site) {
        return TENANT_ORIGIN + "/sites/" + site;
    }

    function buildLibraryPath(site) {
        return "/sites/" + site + "/Shared Documents";
    }

    function buildLibraryUrl(site) {
        return TENANT_ORIGIN + buildLibraryPath(site);
    }

    function getSiteConfig(site) {
        return siteConfigs.find(function (item) { return item.site === site; }) || null;
    }

    function ensureDefaultSite() {
        if (defaultSite && getSiteConfig(defaultSite)) return;
        defaultSite = siteConfigs[0]?.site || "";
    }

    function addSite() {
        const site = normalizeSiteInput(siteInput.value);
        if (!site) {
            showSettingsStatus("error", "❌ Invalid site name.");
            setTimeout(() => showSettingsStatus("", ""), 3000);
            return;
        }
        if (getSiteConfig(site)) {
            defaultSite = site;
            siteInput.value = "";
            renderSiteList();
            renderPathList();
            updateSearchFilterOptions();
            updateSelectedLibraryHint();
            return;
        }

        siteConfigs.push({ site, folders: [] });
        defaultSite = site;
        siteInput.value = "";
        saveSiteConfigs();
        renderSiteList();
        renderPathList();
        updateSearchFilterOptions();
        updateSelectedLibraryHint();
    }

    function removeSite(site) {
        if (!confirm('Remove site "' + site + '" and all of its folders?')) return;
        siteConfigs = siteConfigs.filter(function (item) { return item.site !== site; });
        ensureDefaultSite();
        saveSiteConfigs();
        renderSiteList();
        renderPathList();
        updateSearchFilterOptions();
        updateSelectedLibraryHint();
    }

    function setDefaultSite(site) {
        if (!getSiteConfig(site)) return;
        defaultSite = site;
        saveSiteConfigs();
        renderSiteList();
        renderPathList();
        updateSearchFilterOptions();
        siteFilterSelect.value = site;
        updateFolderFilterOptions();
        updateSelectedLibraryHint();
    }

    function renderSiteList() {
        siteListEl.innerHTML = "";
        if (siteConfigs.length === 0) {
            siteListEl.innerHTML = '<div class="sp-path-list-empty">No sites added yet.</div>';
            return;
        }

        siteConfigs.forEach(function (siteConfig) {
            const row = document.createElement("div");
            row.className = "sp-path-item" + (siteConfig.site === defaultSite ? " sp-path-item-active" : "");
            row.innerHTML =
                '<button class="sp-site-select-btn" title="Select site">' + escapeHtml(siteConfig.site) + '</button>' +
                '<input type="radio" class="sp-site-default-radio" name="sp-default-site" title="Default site" ' + (siteConfig.site === defaultSite ? "checked" : "") + '>' +
                '<button class="sp-path-remove" title="Remove site">🗑️</button>';

            row.querySelector(".sp-site-select-btn").addEventListener("click", function () {
                setDefaultSite(siteConfig.site);
            });

            row.querySelector(".sp-site-default-radio").addEventListener("change", function () {
                setDefaultSite(siteConfig.site);
            });

            row.querySelector(".sp-path-remove").addEventListener("click", function () {
                removeSite(siteConfig.site);
            });

            siteListEl.appendChild(row);
        });
    }

    function updateSelectedLibraryHint() {
        if (!defaultSite) {
            selectedLibraryHintEl.textContent = "Add a site to manage its folders.";
            return;
        }
        selectedLibraryHintEl.textContent = "Default site: " + buildLibraryUrl(defaultSite);
    }

    function migrateLegacyState(data) {
        const fallbackSite = parseSiteFromLibraryUrl(data.configUrl) || DEFAULT_SITE;
        const folders = Array.isArray(data.searchPaths)
            ? data.searchPaths.map(function (item) {
                return typeof item === "string" ? item : item?.path;
            }).filter(Boolean)
            : [];
        siteConfigs = [{ site: fallbackSite, folders }];
        defaultSite = fallbackSite;
    }

    function parseSiteFromLibraryUrl(inputUrl) {
        if (!inputUrl) return "";
        try {
            const pathname = decodeURIComponent(new URL(inputUrl).pathname);
            const match = pathname.match(/^\/sites\/([^/]+)\/Shared Documents/i);
            return match ? match[1] : "";
        } catch (e) {
            return "";
        }
    }

    function addPath() {
        if (!defaultSite) {
            showSettingsStatus("error", "❌ Add a site first.");
            setTimeout(() => showSettingsStatus("", ""), 3000);
            return;
        }

        let path;
        try {
            path = normalizeSubfolderInput(pathInput.value, defaultSite);
        } catch (e) {
            showSettingsStatus("error", "❌ " + e.message);
            setTimeout(() => showSettingsStatus("", ""), 3000);
            return;
        }

        const siteConfig = getSiteConfig(defaultSite);
        if (!siteConfig) return;
        if (!path) return;
        if (siteConfig.folders.includes(path)) { pathInput.value = ""; return; }

        siteConfig.folders.push(path);
        pathInput.value = "";
        saveSiteConfigs();
        renderPathList();
        updateSearchFilterOptions();
    }

    function useCurrentPageUrl() {
        if (!defaultSite) {
            showSettingsStatus("error", "❌ Set a default site first.");
            setTimeout(() => showSettingsStatus("", ""), 3000);
            return;
        }
        pathInput.value = window.location.href;
        pathInput.focus();
    }

    function useCurrentSiteUrl() {
        const site = normalizeSiteInput(window.location.href);
        if (!site) {
            showSettingsStatus("error", "❌ Could not detect a SharePoint site from the current URL.");
            setTimeout(() => showSettingsStatus("", ""), 3000);
            return;
        }
        siteInput.value = site;
        siteInput.focus();
    }

    function normalizeSubfolderInput(rawInput, site) {
        const value = (rawInput || "").trim();
        if (!value) return "";

        let path = value;

        if (/^https?:\/\//i.test(value)) {
            path = extractSubfolderFromUrl(value, site);
        } else if (value.startsWith("/sites/")) {
            path = extractSubfolderFromServerPath(value, site);
        }

        if (!path.startsWith("/")) path = "/" + path;
        if (path.length > 1 && path.endsWith("/")) path = path.slice(0, -1);
        return path;
    }

    function extractSubfolderFromServerPath(folderPath, site) {
        const libraryPath = buildLibraryPath(site);
        if (!folderPath.startsWith(libraryPath)) {
            throw new Error("Path is outside the selected site library.");
        }

        let relativePath = folderPath.slice(libraryPath.length);
        if (!relativePath) {
            throw new Error("The path points to the document library root, not a subfolder.");
        }
        if (!relativePath.startsWith("/")) {
            relativePath = "/" + relativePath;
        }

        return relativePath;
    }

    function extractSubfolderFromUrl(inputUrl, site) {
        let url;
        try {
            url = new URL(inputUrl);
        } catch (e) {
            throw new Error("Invalid subfolder URL.");
        }

        let folderPath = url.searchParams.get("id") || decodeURIComponent(url.pathname);
        folderPath = decodeURIComponent(folderPath);

        return extractSubfolderFromServerPath(folderPath, site);
    }

    function removePath(path) {
        const siteConfig = getSiteConfig(defaultSite);
        if (!siteConfig) return;
        if (!confirm('Remove subfolder "' + path + '"?')) return;
        siteConfig.folders = siteConfig.folders.filter(function (item) { return item !== path; });
        saveSiteConfigs();
        renderPathList();
        updateSearchFilterOptions();
    }

    function renderPathList() {
        pathListEl.innerHTML = "";
        const siteConfig = getSiteConfig(defaultSite);
        if (!siteConfig) {
            pathListEl.innerHTML = '<div class="sp-path-list-empty">Set a default site to manage its folders.</div>';
            return;
        }
        if (siteConfig.folders.length === 0) {
            pathListEl.innerHTML = '<div class="sp-path-list-empty">No subfolders added for this site yet.</div>';
            return;
        }
        siteConfig.folders.forEach((path) => {
            const row = document.createElement("div");
            row.className = "sp-path-item";
            row.innerHTML =
                '<span class="sp-path-text">' + escapeHtml(path) + '</span>' +
                '<input type="text" class="sp-path-edit-input" value="' + escapeHtml(path) + '" style="display:none">' +
                '<button class="sp-path-edit" title="Edit">✏️</button>' +
                '<button class="sp-path-remove" title="Remove">🗑️</button>';

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
                    if (!newPath) { editInput.value = path; }
                    else {
                        newPath = normalizeSubfolderInput(newPath, defaultSite);
                        if (newPath !== path && siteConfig.folders.includes(newPath)) {
                            showSettingsStatus("error", "❌ Folder already exists for this site.");
                            setTimeout(() => showSettingsStatus("", ""), 3000);
                            editInput.value = path;
                            return;
                        }
                        const pathIndex = siteConfig.folders.indexOf(path);
                        if (pathIndex >= 0) {
                            siteConfig.folders[pathIndex] = newPath;
                        }
                        saveSiteConfigs();
                        updateSearchFilterOptions();
                        path = newPath;
                    }
                    textEl.textContent = path;
                    textEl.style.display = "";
                    editInput.style.display = "none";
                    editBtn.textContent = "✏️";
                }
            });

            editInput.addEventListener("keydown", (e) => {
                if (e.key === "Enter") editBtn.click();
                if (e.key === "Escape") {
                    editInput.value = path;
                    textEl.style.display = "";
                    editInput.style.display = "none";
                    editBtn.textContent = "✏️";
                }
            });

            row.querySelector(".sp-path-remove").addEventListener("click", () => removePath(path));
            pathListEl.appendChild(row);
        });
    }

    // ============================================
    // Persistent Storage
    // ============================================
    function saveSiteConfigs() { chrome.storage.local.set({ siteConfigs, defaultSite }); }

    async function restoreState() {
        try {
            const data = await chrome.storage.local.get(["siteConfigs", "searchPaths", "configUrl", "defaultSite"]);

            if (Array.isArray(data.siteConfigs) && data.siteConfigs.length > 0) {
                siteConfigs = data.siteConfigs
                    .map(function (item) {
                        return {
                            site: normalizeSiteInput(item?.site),
                            folders: Array.isArray(item?.folders) ? item.folders.filter(Boolean) : []
                        };
                    })
                    .filter(function (item) { return item.site; });
            } else {
                migrateLegacyState(data);
            }

            if (siteConfigs.length === 0) {
                siteConfigs = [{ site: DEFAULT_SITE, folders: [] }];
            }

            defaultSite = normalizeSiteInput(data.defaultSite) || "";
            ensureDefaultSite();
            saveSiteConfigs();
            renderSiteList();
            renderPathList();
            updateSearchFilterOptions();
            updateSelectedLibraryHint();
            chrome.storage.local.remove("fileData");
            chrome.storage.local.remove("searchPaths");
            chrome.storage.local.remove("configUrl");
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
            showStatus("loading", 'Searching SharePoint for "' + escapeHtml(query) + '"...');

            const selectedSite = siteFilterSelect.value;
            const matched = await searchFilesByName(query, selectedSite);

            const currentQuery = searchInput.value.trim();
            if (currentQuery.length < MIN_QUERY_LENGTH) { showStatus("", ""); return; }

            lastSearchResults = matched;
            lastQuery = currentQuery;
            displayResults(applyFilter(matched), currentQuery);
        } catch (error) {
            console.error("Search error:", error);
            if (error.message.includes("NO_SP_TAB")) {
                showStatus("error", "🔒 No SharePoint tab found! Please open SharePoint in a tab and login first.");
            } else if (error.message.includes("SP_TAB_DISCARDED")) {
                showStatus("error", "🔄 SharePoint tab was unloaded after long inactivity. Please click on the SharePoint tab to reload it, then try again.");
            } else if (error.message.includes("403") || error.message.includes("401") || error.message.includes("429")) {
                showStatus("error", "🔒 Not authenticated or throttled! Please open SharePoint in a tab and login first.");
            } else {
                showStatus("error", "❌ Error: " + error.message);
            }
        } finally {
            isLoading = false;
        }
    }

    async function searchFilesByName(query, selectedSite) {
        const rawQuery = String(query || "").replace(/["']/g, "").trim();
        const normalized = normalizeSearchInput(rawQuery);
        const searchTerms = normalized
            .split(/\s+/)
            .filter(Boolean);

        if (searchTerms.length === 0) {
            return [];
        }

        let activeConfigs = siteConfigs;
        if (selectedSite) {
            activeConfigs = siteConfigs.filter(c => c.site === selectedSite);
        }

        const pathClauses = activeConfigs.map(function (siteConfig) {
            return 'path:"' + buildLibraryUrl(siteConfig.site) + '"';
        });

        // Exact logic from dist extension "title" mode
        const termClauses = searchTerms.map(function (term) {
            return "(Title:" + term + " OR Title:" + term + "* OR Filename:" + term + " OR Filename:" + term + "* OR FileLeafRef:" + term + " OR FileLeafRef:" + term + "*)";
        });

        let queryText = termClauses.join(" ");
        if (pathClauses.length > 0) {
            queryText += " " + (pathClauses.length === 1 ? pathClauses[0] : "(" + pathClauses.join(" OR ") + ")");
        }

        // Use selected site URL as API base if available, otherwise fallback
        const baseSite = selectedSite || (siteConfigs.length > 0 ? siteConfigs[0].site : null);
        const siteUrl = baseSite ? buildSiteUrl(baseSite) : TENANT_ORIGIN;

        const rows = await fetchSearchRows(queryText, siteUrl);
        const uniqueFiles = new Map();

        rows.forEach(function (row) {
            const cells = row.Cells?.results || [];
            const getVal = function (key) {
                const cell = cells.find(function (item) { return item.Key === key; });
                return cell ? cell.Value : "";
            };

            const absolutePath = getVal("Path");
            if (!absolutePath) return;

            let relativePath = "";
            try {
                relativePath = decodeURIComponent(new URL(absolutePath).pathname);
            } catch (e) {
                return;
            }

            const matchedSiteConfig = getSiteConfigByServerRelativePath(relativePath);
            if (!matchedSiteConfig) {
                return;
            }

            const isContainer = getVal("IsContainer") === "true";
            if (isContainer) return; // Keep folder filtering for clean UI

            const name = getVal("Filename") || getVal("FileLeafRef") || getVal("Title") || relativePath.split("/").pop() || "";

            const fileKey = relativePath.toLowerCase();
            if (uniqueFiles.has(fileKey)) {
                return;
            }

            uniqueFiles.set(fileKey, {
                name,
                path: relativePath,
                site: matchedSiteConfig.site,
                siteUrl: buildSiteUrl(matchedSiteConfig.site),
                size: parseInt(getVal("Size") || "0", 10),
                modified: getVal("LastModifiedTime") || "",
                isCheckedOut: false
            });
        });

        return Array.from(uniqueFiles.values());
    }

    async function fetchSearchRows(queryText, apiBase) {
        const rows = [];
        const rowLimit = 500;
        const maxRows = 500;
        const baseUrl = apiBase || TENANT_ORIGIN;

        for (let startRow = 0; startRow < maxRows; startRow += rowLimit) {
            const searchUrl = baseUrl +
                "/_api/search/query?querytext='" + encodeURIComponent(queryText) +
                "'&startrow=" + startRow +
                "&rowlimit=" + rowLimit +
                "&selectproperties='Title,Path,Filename,FileLeafRef,LastModifiedTime,Size,EditorOWSUSER,ContentClass,FileExtension,IsDocument,IsContainer'" +
                "&sortlist='IsContainer:descending,Rank:descending'";

            const searchData = await bgFetch(searchUrl, {
                method: "GET",
                headers: { "Accept": "application/json;odata=verbose" }
            });

            const pageRows = searchData?.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows?.results || [];
            rows.push(...pageRows);
            if (pageRows.length < rowLimit) {
                break;
            }
        }

        return rows;
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
    // Search Scope
    // ============================================
    function getSearchQueryPaths() {
        return siteConfigs.map(function (siteConfig) {
            return buildLibraryUrl(siteConfig.site);
        });
    }

    function getSiteConfigByServerRelativePath(serverRelativePath) {
        return siteConfigs.find(function (siteConfig) {
            const libraryPath = buildLibraryPath(siteConfig.site);
            return serverRelativePath === libraryPath || serverRelativePath.startsWith(libraryPath + "/");
        }) || null;
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
            const viewFileUrl = getViewFileUrl(file.path);

            const item = document.createElement("div");
            item.className = "sp-file-item";
            item.title = "Open: " + file.name;
            item.style.cursor = "pointer";
            item.addEventListener("click", () => window.open(viewFileUrl, "_blank"));

            item.innerHTML =
                '<span class="sp-file-icon">' + getFileIcon(file.name) + '</span>' +
                '<div class="sp-file-info">' +
                '<div class="sp-file-name">' + highlightText(escapeHtml(file.name), query) + '</div>' +
                '<div class="sp-file-path"></div>' +
                (file.size ? '<div class="sp-file-size">' + formatFileSize(file.size) + '</div>' : '') +
                '</div>' +
                '<div class="sp-file-actions">' +
                (canOpenInDesktopApp(file) ? '<button class="sp-action-btn sp-action-desktop" title="Open in desktop app">🖥️</button>' : '') +
                '<button class="sp-action-btn sp-action-checkout" title="' + (file.isCheckedOut ? 'Check in file' : 'Check out file') + '">' + getCheckoutActionIcon(file.isCheckedOut) + '</button>' +
                '</div>';

            const desktopBtn = item.querySelector(".sp-action-desktop");
            if (desktopBtn) {
                desktopBtn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    openFileInDesktopApp(file);
                });
            }

            item.querySelector(".sp-action-checkout").addEventListener("click", async (e) => {
                e.stopPropagation();
                await toggleFileCheckout(file, item.querySelector(".sp-action-checkout"));
            });

            const pathEl = item.querySelector(".sp-file-path");
            pathEl.appendChild(buildFolderBreadcrumb(file));

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
                await sharePointCheckOut(file);
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
                await sharePointCheckIn(file, comment);
                file.isCheckedOut = false;
                buttonEl.textContent = getCheckoutActionIcon(false);
                buttonEl.title = "Check out file";
                showStatus("success", '✅ Checked in "' + escapeHtml(file.name) + '"');
            }
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

    async function sharePointCheckOut(file) {
        const siteUrl = file?.siteUrl || getSiteUrlFromServerRelativePath(file?.path);
        const serverRelativePath = file?.path || "";
        const url = siteUrl +
            "/_api/web/getfilebyserverrelativeurl('" +
            encodeSharePointPath(serverRelativePath) +
            "')/checkOut()";

        await bgFetch(url, {
            method: "POST",
            headers: { "Accept": "application/json;odata=verbose" }
        });
    }

    async function sharePointCheckIn(file, comment) {
        const siteUrl = file?.siteUrl || getSiteUrlFromServerRelativePath(file?.path);
        const serverRelativePath = file?.path || "";
        const url = siteUrl +
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
        currentFilePathEl.innerHTML = "";
        currentFilePathEl.appendChild(buildFolderBreadcrumb(currentPageFile));
        currentFileDesktopBtn.classList.toggle("sp-hidden", !canOpenInDesktopApp(currentPageFile));
        currentFileCheckoutBtn.textContent = getCheckoutActionIcon(currentPageFile.isCheckedOut);
        currentFileCheckoutBtn.title = currentPageFile.isCheckedOut ? "Check in file" : "Check out file";
        currentFileEl.classList.remove("sp-hidden");
    }

    // ============================================
    // Filter Helpers
    // ============================================
    function getFileTypeCategory(fileName) {
        const ext = String(fileName || "").split(".").pop().toLowerCase();
        if (["xls", "xlsx", "xlsm", "xlsb", "csv"].includes(ext)) return "excel";
        if (["doc", "docx", "dot", "dotx"].includes(ext)) return "word";
        if (["ppt", "pptx", "pptm", "pps", "ppsx"].includes(ext)) return "powerpoint";
        if (ext === "pdf") return "pdf";
        return "other";
    }

    function applyFilter(files) {
        const selectedSite = siteFilterSelect.value;
        const selectedFolderKey = folderFilterSelect.value;
        const selectedType = fileTypeFilterSelect.value;

        return files.filter(function (file) {
            if (selectedSite && file.site !== selectedSite) {
                return false;
            }
            if (selectedFolderKey) {
                const parts = selectedFolderKey.split("|");
                const folderSite = parts[0];
                const folderPath = parts.slice(1).join("|");
                const folderRoot = buildLibraryPath(folderSite) + folderPath;
                if (!(file.site === folderSite && (file.path === folderRoot || file.path.startsWith(folderRoot + "/")))) {
                    return false;
                }
            }
            if (selectedType && getFileTypeCategory(file.name) !== selectedType) {
                return false;
            }
            return true;
        });
    }

    function updateSearchFilterOptions() {
        const currentSiteValue = siteFilterSelect.value;
        siteFilterSelect.innerHTML = '<option value="">All sites</option>';
        siteConfigs.forEach(function (siteConfig) {
            const opt = document.createElement("option");
            opt.value = siteConfig.site;
            opt.textContent = siteConfig.site;
            siteFilterSelect.appendChild(opt);
        });
        siteFilterSelect.value = currentSiteValue || defaultSite || "";
        if (siteFilterSelect.value !== (currentSiteValue || defaultSite || "")) {
            siteFilterSelect.value = defaultSite || "";
        }
        updateFolderFilterOptions();
    }

    function updateFolderFilterOptions() {
        const selectedSite = siteFilterSelect.value;
        const currentValue = folderFilterSelect.value;
        folderFilterSelect.innerHTML = '<option value="">All folders</option>';

        siteConfigs.forEach(function (siteConfig) {
            if (selectedSite && siteConfig.site !== selectedSite) {
                return;
            }
            siteConfig.folders.forEach(function (path) {
                const opt = document.createElement("option");
                opt.value = siteConfig.site + "|" + path;
                opt.textContent = selectedSite ? path : siteConfig.site + " " + path;
                folderFilterSelect.appendChild(opt);
            });
        });

        folderFilterSelect.value = currentValue;
        if (folderFilterSelect.value !== currentValue) {
            folderFilterSelect.value = "";
        }
    }

    // ============================================
    // URL Helpers
    // ============================================
    function getViewFileUrl(serverRelativePath) {
        return getFileWebUrl(serverRelativePath) + "?web=1";
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

    function normalizeSearchInput(value) {
        return String(value || "")
            .replace(/[()（）{}[\]【】「」『』.·。_\-]/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    function sanitizeSearchTerm(term) {
        return term.replace(/["'*]/g, "").trim();
    }

    function matchesAllSearchTerms(text, searchTerms) {
        // Normalize using NFKC to convert full-width alphanumeric to half-width,
        // and half-width katakana to full-width, ensuring robust Japanese string matching.
        const value = String(text || "").normalize('NFKC').toLowerCase();
        return searchTerms.every(function (term) {
            return value.includes(term.normalize('NFKC').toLowerCase());
        });
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

    function getSiteUrlFromServerRelativePath(serverRelativePath) {
        const site = getSiteSlugFromServerRelativePath(serverRelativePath);
        return site ? buildSiteUrl(site) : "";
    }

    function getSiteSlugFromServerRelativePath(serverRelativePath) {
        if (!serverRelativePath) return "";
        const match = String(serverRelativePath).match(/^\/sites\/([^/]+)/i);
        return match ? match[1] : "";
    }

    function getFolderPath(serverRelativePath, fileName) {
        return String(serverRelativePath || "").replace("/" + (fileName || ""), "");
    }

    function buildFolderBreadcrumb(file) {
        const wrapper = document.createElement("div");
        const serverRelativePath = file?.path || "";
        const site = file?.site || getSiteConfigByServerRelativePath(serverRelativePath)?.site || getSiteSlugFromServerRelativePath(serverRelativePath);
        const libraryPath = site ? buildLibraryPath(site) : "";
        const parentFolderPath = getFolderPath(serverRelativePath, file?.name || "");
        const relativePath = libraryPath && parentFolderPath.startsWith(libraryPath)
            ? parentFolderPath.slice(libraryPath.length)
            : "";
        const segments = relativePath.split("/").filter(Boolean);

        wrapper.appendChild(document.createTextNode(site || ""));

        if (segments.length === 0) {
            wrapper.appendChild(document.createTextNode(" / "));
            wrapper.appendChild(createFolderLink("/", getFolderUrlFromFolderPath(parentFolderPath)));
            return wrapper;
        }

        let currentPath = "";
        segments.forEach(function (segment, index) {
            currentPath += "/" + segment;
            wrapper.appendChild(document.createTextNode(index === 0 ? " / " : " / "));
            wrapper.appendChild(createFolderLink(segment, getFolderUrlFromFolderPath(libraryPath + currentPath)));
        });

        return wrapper;
    }

    function createFolderLink(label, href) {
        const link = document.createElement("a");
        link.className = "sp-file-path-link";
        link.href = href;
        link.textContent = label;
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        link.addEventListener("click", function (e) {
            e.stopPropagation();
        });
        return link;
    }

    function getFolderUrl(serverRelativePath) {
        const parentFolderPath = getFolderPath(serverRelativePath, String(serverRelativePath || "").split("/").pop() || "");
        return getFolderUrlFromFolderPath(parentFolderPath);
    }

    function getFolderUrlFromFolderPath(folderPath) {
        const libraryRootUrl = getLibraryRootUrlByPath(folderPath);
        return libraryRootUrl + "/Forms/AllItems.aspx?id=" + encodeURIComponent(folderPath);
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

    function canOpenInDesktopApp(file) {
        return !!getDesktopOpenUrl(file);
    }

    function openFileInDesktopApp(file) {
        const desktopUrl = getDesktopOpenUrl(file);
        if (!desktopUrl) {
            showStatus("error", "❌ Desktop open is only available for Word, Excel, and PowerPoint files");
            return;
        }
        window.location.href = desktopUrl;
    }

    function getDesktopOpenUrl(file) {
        const protocol = getOfficeDesktopProtocol(file && file.name);
        const webUrl = getRawFileUrl(file);
        if (!protocol || !webUrl) return "";
        return protocol + encodeURI(webUrl);
    }

    function getOfficeDesktopProtocol(fileName) {
        const ext = String(fileName || "").split(".").pop().toLowerCase();
        if (ext === "doc" || ext === "docx") return "ms-word:ofe|u|";
        if (ext === "xls" || ext === "xlsx" || ext === "xlsm" || ext === "xlsb") return "ms-excel:ofe|u|";
        if (ext === "ppt" || ext === "pptx" || ext === "pptm") return "ms-powerpoint:ofe|u|";
        return "";
    }

    function getRawFileUrl(file) {
        if (!file || !file.path) return "";
        return getFileWebUrl(file.path);
    }

    function getFileWebUrl(serverRelativePath) {
        const path = String(serverRelativePath || "");
        if (!path) return "";

        const siteConfig = getSiteConfigByServerRelativePath(path);
        const libraryRootUrl = getLibraryRootUrlByPath(path);
        if (libraryRootUrl && siteConfig) {
            const relativePath = path.slice(buildLibraryPath(siteConfig.site).length);
            return libraryRootUrl + relativePath;
        }

        try {
            return TENANT_ORIGIN + path;
        } catch (e) {
            return path;
        }
    }

    function getLibraryRootUrlByPath(serverRelativePath) {
        const siteConfig = getSiteConfigByServerRelativePath(serverRelativePath);
        const site = siteConfig?.site || getSiteSlugFromServerRelativePath(serverRelativePath);
        return site ? buildLibraryUrl(site) : "";
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
