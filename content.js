// ============================================
// SharePoint Utils - Content Script
// Injected into the active tab to show a modal overlay
// ============================================
(function () {
    const previouslyFocusedElement = document.activeElement;

    // Prevent double injection
    if (document.getElementById("sp-search-backdrop")) {
        document.getElementById("sp-search-backdrop").remove();
    }

    // ============================================
    // CONFIG
    // ============================================
    const TENANT_ORIGIN = "https://dxcportal.sharepoint.com";
    const DEFAULT_LIBRARY_NAME = "Shared Documents";
    const DEFAULT_SITE_PATH_PREFIX = "sites";
    const DEFAULT_SITE = "msteams_268a65";

    const MIN_QUERY_LENGTH = 3;
    const CURRENT_FILE_METADATA_TIMEOUT_MS = 9000;
    const SEARCH_PAGE_SIZE = 100;
    const CURRENT_FILE_EXTENSIONS = new Set(["doc", "docx", "xls", "xlsx", "xlsm", "xlsb", "ppt", "pptx", "pptm"]);
    const SP_UTILS = globalThis.SharePointUtils;
    if (!SP_UTILS) throw new Error("SharePoint Utils helpers were not loaded");
    const {
        encodeODataString,
        encodeSharePointPath,
        escapeHtml,
        escapeRegex,
        formatDate,
        formatEditorName,
        formatFileSize,
        getFileIcon,
        getFileTypeCategory,
        getOfficeDesktopProtocol,
        isFileCheckedOut,
        normalizeSearchInput
    } = SP_UTILS;

    // siteConfigs: array of { site: string, folders: string[] }
    let siteConfigs = [];
    let defaultSite = "";
    let isLoading = false;
    let lastSearchResults = []; // store last search results for re-filtering
    let lastQuery = "";
    let searchState = null;
    let favoriteFiles = [];
    let recentFiles = [];
    let recentSearches = [];
    let savedSearches = [];
    let syncEnabled = false;
    let currentPageFile = null;

    // ============================================
    // Build Modal HTML
    // ============================================
    const backdrop = document.createElement("div");
    backdrop.id = "sp-search-backdrop";
    backdrop.className = "sp-search-backdrop";
    backdrop.innerHTML = `
        <div class="sp-search-modal" id="sp-search-modal" role="dialog" aria-modal="true" aria-labelledby="sp-modal-title">
            <div class="sp-search-header">
                <img src="${chrome.runtime.getURL('icons/icon32.png')}" class="sp-search-header-icon" alt="S">
                <h2 id="sp-modal-title">SharePoint Utils</h2>
                <button type="button" class="sp-save-search-btn" id="sp-save-search-btn" title="Save current search" aria-label="Save current search">🔖</button>
                <button type="button" class="sp-settings-icon-btn" id="sp-settings-icon-btn" title="Settings" aria-label="Open settings">⚙️</button>
                <button type="button" class="sp-search-close" id="sp-close-btn" title="Close (Esc)" aria-label="Close SharePoint Utils">✕</button>
            </div>

            <div class="sp-current-file sp-hidden" id="sp-current-file" aria-live="polite">
                <div class="sp-current-file-info">
                    <div class="sp-current-file-name" id="sp-current-file-name"></div>
                    <div class="sp-file-path" id="sp-current-file-path"></div>
                    <div class="sp-file-meta" id="sp-current-file-meta"></div>
                </div>
                <div class="sp-file-actions">
                    <button type="button" class="sp-action-btn sp-action-desktop" id="sp-current-file-desktop-btn" title="Open in desktop app" aria-label="Open in desktop app">🖥️</button>
                    <button type="button" class="sp-action-btn sp-action-checkout" id="sp-current-file-checkout-btn" title="Check out file" aria-label="Check out file">📝</button>
                    <button type="button" class="sp-action-btn sp-action-share" id="sp-current-file-share-btn" title="Copy share link" aria-label="Copy share link">🔗</button>
                </div>
            </div>

            <!-- Search Panel (always visible) -->
            <div class="sp-search-panel" id="sp-tab-search">
                <div class="sp-search-input-section">
                    <div class="sp-search-row">
                        <div class="sp-search-input-wrap">
                            <span class="sp-search-icon">🔍</span>
                            <input type="text" class="sp-search-input" id="sp-search-input" placeholder="Type filename and press Enter…" autocomplete="off" aria-label="Search SharePoint files">
                            <button type="button" class="sp-clear-btn" id="sp-clear-btn" title="Clear" aria-label="Clear search">✕</button>
                        </div>
                        <div class="sp-inline-filters">
                            <select class="sp-filter-select sp-filter-inline" id="sp-site-filter-select" title="Filter by site" aria-label="Filter by site"></select>
                            <select class="sp-filter-select sp-filter-inline" id="sp-folder-filter-select" title="Filter by subfolder" aria-label="Filter by subfolder">
                                <option value="">All folders</option>
                            </select>
                            <select class="sp-filter-select sp-filter-inline" id="sp-type-filter-select" title="Filter by file type" aria-label="Filter by file type">
                                <option value="">All types</option>
                                <option value="excel">Excel</option>
                                <option value="word">Word</option>
                                <option value="powerpoint">PowerPoint</option>
                                <option value="pdf">PDF</option>
                                <option value="other">Other</option>
                            </select>
                            <select class="sp-filter-select sp-filter-inline" id="sp-sort-filter-select" title="Sort results" aria-label="Sort results">
                                <option value="relevance">Relevance</option>
                                <option value="nameAsc">Name A–Z</option>
                                <option value="nameDesc">Name Z–A</option>
                                <option value="modifiedDesc">Newest</option>
                                <option value="modifiedAsc">Oldest</option>
                                <option value="sizeDesc">Largest</option>
                                <option value="sizeAsc">Smallest</option>
                            </select>
                            <input type="date" class="sp-filter-date" id="sp-modified-after-filter" title="Modified after" aria-label="Show files modified after">
                        </div>
                    </div>
                </div>
                <div class="sp-results" id="sp-results">
                    <div class="sp-search-empty-state" id="sp-search-empty-state">
                        <div class="sp-search-empty-icon">🔍</div>
                        <div class="sp-search-empty-title">Search for a file</div>
                        <div class="sp-search-empty-hint">Type a filename and press <kbd>Enter</kbd> to search across your SharePoint sites</div>
                        <div class="sp-empty-extras" id="sp-empty-extras"></div>
                    </div>
                </div>
            </div>

            <!-- Settings Modal Overlay (inside the main modal) -->
            <div class="sp-settings-overlay sp-hidden" id="sp-settings-overlay" aria-hidden="true">
                <div class="sp-settings-overlay-header">
                    <span class="sp-settings-overlay-title">⚙️ Settings</span>
                    <button type="button" class="sp-search-close" id="sp-settings-close-btn" title="Close Settings" aria-label="Close settings">✕</button>
                </div>
                <div class="sp-settings-content">
                    <div class="sp-settings-section sp-settings-backup-section">
                        <div class="sp-settings-label">Backup & Sync</div>
                        <label class="sp-sync-toggle">
                            <input type="checkbox" id="sp-sync-config-toggle">
                            <span>Sync sites and folders across Chrome</span>
                        </label>
                        <div class="sp-backup-actions">
                            <button type="button" class="sp-add-path-btn sp-btn-secondary" id="sp-export-config-btn">Export JSON</button>
                            <button type="button" class="sp-add-path-btn sp-btn-secondary" id="sp-import-config-btn">Import JSON</button>
                            <input type="file" id="sp-import-config-input" accept="application/json,.json" class="sp-file-input-hidden">
                        </div>
                        <div class="sp-refresh-hint">If synced data already exists, enabling asks whether to keep Local or Chrome profile. Later opens use the synced configuration; local changes are copied back automatically.</div>
                        <div class="sp-refresh-hint">Export creates a portable backup of sites, libraries, folders, and the default site.</div>
                    </div>
                    <div class="sp-settings-section">
                        <div class="sp-settings-label">Add Site</div>
                        <div class="sp-path-add-row">
                            <input type="text" class="sp-path-input" id="sp-site-input" placeholder="site-name" autocomplete="off">
                            <button class="sp-add-path-btn" id="sp-add-site-btn">+ Add</button>
                            <button class="sp-add-path-btn" id="sp-current-site-btn">Current Site</button>
                        </div>
                        <div class="sp-refresh-hint">Enter a site name or full library URL, e.g. ${TENANT_ORIGIN}/sites/{site}/Shared Documents</div>
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
                    <div class="sp-status sp-hidden" id="sp-settings-status" role="status" aria-live="polite"></div>
                </div>
                <div class="sp-footer" style="display:flex; align-items:center; justify-content:space-between;">
                    <span id="sp-settings-connection-status"></span>
                    <span class="sp-version">v1.1</span>
                </div>
            </div>

            <div class="sp-footer" style="display:flex; align-items:center; justify-content:space-between;">
                    <span id="sp-connection-status" aria-live="polite">⏳ Checking connection...</span>
                <span class="sp-version">v1.1</span>
            </div>
        </div>
        <div class="sp-inline-modal-backdrop sp-hidden" id="sp-inline-modal-backdrop" aria-hidden="true">
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
        <div class="sp-inline-modal-backdrop sp-hidden" id="sp-share-modal-backdrop" aria-hidden="true">
            <div class="sp-inline-modal" id="sp-share-modal" role="dialog" aria-modal="true" aria-labelledby="sp-share-modal-title">
                <div class="sp-inline-modal-title" id="sp-share-modal-title">Share Options</div>
                <div class="sp-inline-modal-subtitle" id="sp-share-modal-filename" style="word-break: break-all;"></div>
                <div class="sp-inline-modal-actions" style="display:flex; flex-direction:column; gap:8px; margin-top:16px;">
                    <button class="sp-inline-modal-btn sp-inline-modal-btn-secondary" id="sp-share-link-btn" style="width:100%; text-align:left; justify-content:flex-start;">🔗 Share Link (URL only)</button>
                    <button class="sp-inline-modal-btn sp-inline-modal-btn-secondary" id="sp-share-backlog-btn" style="width:100%; text-align:left; justify-content:flex-start;">📝 Backlog Format [Name](URL)</button>
                    <button class="sp-inline-modal-btn sp-inline-modal-btn-primary" id="sp-share-cancel-btn" style="width:100%; margin-top:8px; justify-content:center; display:flex;">Cancel</button>
                </div>
            </div>
        </div>
        <div class="sp-inline-modal-backdrop sp-hidden" id="sp-sync-modal-backdrop" aria-hidden="true">
            <div class="sp-inline-modal" id="sp-sync-modal" role="dialog" aria-modal="true" aria-labelledby="sp-sync-modal-title" aria-describedby="sp-sync-modal-subtitle sp-sync-modal-warning">
                <div class="sp-inline-modal-title" id="sp-sync-modal-title">Choose sync source</div>
                <div class="sp-inline-modal-subtitle" id="sp-sync-modal-subtitle"></div>
                <div class="sp-inline-modal-warning" id="sp-sync-modal-warning">
                    Synced configuration already exists. Choose which side to keep; the other side will be overwritten.
                </div>
                <div class="sp-inline-modal-actions sp-sync-modal-actions">
                    <button class="sp-inline-modal-btn sp-inline-modal-btn-secondary" id="sp-sync-use-local-btn">Keep Local → Upload to profile</button>
                    <button class="sp-inline-modal-btn sp-inline-modal-btn-primary" id="sp-sync-use-profile-btn">Use Chrome profile → Replace local</button>
                    <button class="sp-inline-modal-btn sp-inline-modal-btn-secondary" id="sp-sync-cancel-btn">Cancel</button>
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
    const syncConfigToggle = document.getElementById("sp-sync-config-toggle");
    const exportConfigBtn = document.getElementById("sp-export-config-btn");
    const importConfigBtn = document.getElementById("sp-import-config-btn");
    const importConfigInput = document.getElementById("sp-import-config-input");
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
    const currentFileMetaEl = document.getElementById("sp-current-file-meta");
    const currentFileShareBtn = document.getElementById("sp-current-file-share-btn");
    const currentFileDesktopBtn = document.getElementById("sp-current-file-desktop-btn");
    const currentFileCheckoutBtn = document.getElementById("sp-current-file-checkout-btn");
    const inlineModalBackdrop = document.getElementById("sp-inline-modal-backdrop");
    const inlineModal = document.getElementById("sp-inline-modal");
    const inlineModalSubtitleEl = document.getElementById("sp-inline-modal-subtitle");
    const inlineModalTextarea = document.getElementById("sp-inline-modal-textarea");
    const inlineModalCancelBtn = document.getElementById("sp-inline-modal-cancel-btn");
    const inlineModalOkBtn = document.getElementById("sp-inline-modal-ok-btn");
    
    const shareModalBackdrop = document.getElementById("sp-share-modal-backdrop");
    const shareModal = document.getElementById("sp-share-modal");
    const shareModalFilename = document.getElementById("sp-share-modal-filename");
    const shareLinkBtn = document.getElementById("sp-share-link-btn");
    const shareBacklogBtn = document.getElementById("sp-share-backlog-btn");
    const shareCancelBtn = document.getElementById("sp-share-cancel-btn");
    const syncModalBackdrop = document.getElementById("sp-sync-modal-backdrop");
    const syncModal = document.getElementById("sp-sync-modal");
    const syncModalSubtitleEl = document.getElementById("sp-sync-modal-subtitle");
    const syncUseLocalBtn = document.getElementById("sp-sync-use-local-btn");
    const syncUseProfileBtn = document.getElementById("sp-sync-use-profile-btn");
    const syncCancelBtn = document.getElementById("sp-sync-cancel-btn");
    let activeShareFile = null;
    let activeShareUrl = null;

    const CHECKOUT_ICON = "📝";
    const CHECKIN_ICON = "✅";
    let activeInlineModal = null;
    let activeSyncConflict = null;

    const siteFilterSelect = document.getElementById("sp-site-filter-select");
    const folderFilterSelect = document.getElementById("sp-folder-filter-select");
    const fileTypeFilterSelect = document.getElementById("sp-type-filter-select");
    const sortFilterSelect = document.getElementById("sp-sort-filter-select");
    const modifiedAfterInput = document.getElementById("sp-modified-after-filter");
    const settingsIconBtn = document.getElementById("sp-settings-icon-btn");
    const saveSearchBtn = document.getElementById("sp-save-search-btn");
    const settingsOverlay = document.getElementById("sp-settings-overlay");
    const settingsCloseBtn = document.getElementById("sp-settings-close-btn");

    // ============================================
    // Settings Overlay Open/Close
    // ============================================
    function openSettingsOverlay() {
        settingsOverlay.classList.remove("sp-hidden");
        settingsOverlay.setAttribute("aria-hidden", "false");
        settingsIconBtn.classList.add("sp-active");
        const settingsConnEl = document.getElementById("sp-settings-connection-status");
        if (settingsConnEl) settingsConnEl.textContent = connectionEl.textContent;
    }
    function closeSettingsOverlay() {
        settingsOverlay.classList.add("sp-hidden");
        settingsOverlay.setAttribute("aria-hidden", "true");
        settingsIconBtn.classList.remove("sp-active");
        searchInput.focus();
    }
    settingsIconBtn.addEventListener("click", () => {
        if (settingsOverlay.classList.contains("sp-hidden")) {
            openSettingsOverlay();
        } else {
            closeSettingsOverlay();
        }
    });
    settingsCloseBtn.addEventListener("click", closeSettingsOverlay);
    saveSearchBtn.addEventListener("click", saveCurrentSearch);
    syncConfigToggle.addEventListener("change", toggleConfigSync);
    exportConfigBtn.addEventListener("click", exportConfig);
    importConfigBtn.addEventListener("click", () => importConfigInput.click());
    importConfigInput.addEventListener("change", importConfigFromFile);

    // ============================================
    // Events
    // ============================================
    closeBtn.addEventListener("click", closeModal);
    backdrop.addEventListener("click", (e) => {
        if (e.target === backdrop) closeModal();
    });
    backdrop.addEventListener("keydown", (e) => {
        if (e.key === "Tab" && document.getElementById("sp-search-backdrop")) {
            trapModalFocus(e);
            return;
        }
        if (e.key === "Escape" && activeSyncConflict) {
            closeSyncConflict(null);
            return;
        }
        if (e.key === "Escape" && activeInlineModal) {
            closeInlineModal(false);
            return;
        }
        if (e.key === "Escape" && activeShareFile) {
            closeShareModal();
            return;
        }
        if (e.key === "Escape" && !settingsOverlay.classList.contains("sp-hidden")) {
            closeSettingsOverlay();
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

    shareModal.addEventListener("click", (e) => e.stopPropagation());
    shareModalBackdrop.addEventListener("click", () => closeShareModal());
    shareCancelBtn.addEventListener("click", () => closeShareModal());
    shareLinkBtn.addEventListener("click", async () => {
        if (!activeShareUrl) return;
        try {
            const encodedUrl = encodeURI(activeShareUrl);
            await navigator.clipboard.writeText(encodedUrl);
            showStatus("success", '🟢 Copied link to clipboard');
        } catch (err) {
            showStatus("error", '❌ Failed to copy link');
        }
        closeShareModal();
    });
    shareBacklogBtn.addEventListener("click", async () => {
        if (!activeShareUrl || !activeShareFile) return;
        try {
            const encodedUrl = encodeURI(activeShareUrl);
            const text = `[${activeShareFile.name}](${encodedUrl})`;
            await navigator.clipboard.writeText(text);
            showStatus("success", '🟢 Copied backlog format to clipboard');
        } catch (err) {
            showStatus("error", '❌ Failed to copy backlog format');
        }
        closeShareModal();
    });

    syncModal.addEventListener("click", (e) => e.stopPropagation());
    syncModalBackdrop.addEventListener("click", () => closeSyncConflict(null));
    syncUseLocalBtn.addEventListener("click", () => closeSyncConflict("local"));
    syncUseProfileBtn.addEventListener("click", () => closeSyncConflict("profile"));
    syncCancelBtn.addEventListener("click", () => closeSyncConflict(null));

    currentFileShareBtn.addEventListener("click", async (e) => {
        e.stopPropagation();
        if (!currentPageFile) return;
        openShareModal(currentPageFile, getCurrentPageFileShareUrl(currentPageFile));
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
            showEmptyState();
            showStatus("", "");
            lastSearchResults = [];
            lastQuery = "";
            searchState = null;
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
        showEmptyState();
        showStatus("", "");
        lastSearchResults = [];
        lastQuery = "";
        searchState = null;
        searchInput.focus();
    });

    // Filter dropdowns — re-filter existing results instantly
    siteFilterSelect.addEventListener("change", () => {
        updateFolderFilterOptions();
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(getVisibleResults(), lastQuery);
        }
    });

    folderFilterSelect.addEventListener("change", () => {
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(getVisibleResults(), lastQuery);
        }
    });

    fileTypeFilterSelect.addEventListener("change", () => {
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(getVisibleResults(), lastQuery);
        }
    });

    sortFilterSelect.addEventListener("change", () => {
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(getVisibleResults(), lastQuery);
        }
    });

    modifiedAfterInput.addEventListener("change", () => {
        if (lastSearchResults.length > 0 && lastQuery) {
            displayResults(getVisibleResults(), lastQuery);
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
    restoreState().finally(() => detectCurrentPageFile());
    checkConnection();

    // ============================================
    // Close Modal
    // ============================================
    function closeModal() {
        const el = document.getElementById("sp-search-backdrop");
        if (el) el.remove();
        if (previouslyFocusedElement && typeof previouslyFocusedElement.focus === "function") {
            previouslyFocusedElement.focus();
        }
    }

    function trapModalFocus(event) {
        let focusRoot = modal;
        if (activeSyncConflict) focusRoot = syncModal;
        else if (activeInlineModal) focusRoot = inlineModal;
        else if (activeShareFile) focusRoot = shareModal;
        else if (!settingsOverlay.classList.contains("sp-hidden")) focusRoot = settingsOverlay;

        const focusable = Array.from(focusRoot.querySelectorAll(
            'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), a[href], [tabindex]:not([tabindex="-1"])'
        )).filter(function (element) {
            return element.getClientRects().length > 0;
        });
        if (!focusable.length) return;

        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (event.shiftKey && (document.activeElement === first || !focusRoot.contains(document.activeElement))) {
            event.preventDefault();
            last.focus();
        } else if (!event.shiftKey && (document.activeElement === last || !focusRoot.contains(document.activeElement))) {
            event.preventDefault();
            first.focus();
        }
    }

    // ============================================
    // Site & Path Management
    // ============================================
    function normalizeOrigin(rawOrigin) {
        try {
            const url = new URL(String(rawOrigin || "").trim());
            if (!/^https?:$/i.test(url.protocol)) return "";
            return url.origin;
        } catch (e) {
            return "";
        }
    }

    function parseSiteLocation(rawInput) {
        const value = String(rawInput || "").trim();
        if (!value) return null;

        if (/^https?:\/\//i.test(value)) {
            try {
                const url = new URL(value);
                const parts = decodeURIComponent(url.pathname).split("/").filter(Boolean);
                const markerIndex = parts.findIndex(part => /^(sites|teams)$/i.test(part));
                if (markerIndex < 0 || !parts[markerIndex + 1]) return null;
                return {
                    origin: url.origin,
                    site: parts[markerIndex + 1],
                    sitePathPrefix: parts[markerIndex].toLowerCase(),
                    library: parts[markerIndex + 2] || DEFAULT_LIBRARY_NAME
                };
            } catch (e) {
                return null;
            }
        }

        const site = value
            .replace(/^\/+/, "")
            .replace(/^(sites|teams)\//i, "")
            .replace(/\/.*$/, "")
            .trim();
        return site ? {
            origin: TENANT_ORIGIN,
            site,
            sitePathPrefix: DEFAULT_SITE_PATH_PREFIX,
            library: DEFAULT_LIBRARY_NAME
        } : null;
    }

    function normalizeSiteInput(rawInput) {
        return parseSiteLocation(rawInput)?.site || "";
    }

    function normalizeSiteConfig(rawConfig) {
        const input = rawConfig || {};
        const parsed = parseSiteLocation(input.libraryUrl || input.site || input.url);
        const site = normalizeSiteInput(input.site || input.libraryUrl || input.url);
        if (!site) return null;
        const origin = normalizeOrigin(input.origin) || parsed?.origin || TENANT_ORIGIN;
        const sitePathPrefix = /^(sites|teams)$/i.test(input.sitePathPrefix || "")
            ? input.sitePathPrefix.toLowerCase()
            : (parsed?.sitePathPrefix || DEFAULT_SITE_PATH_PREFIX);
        const library = String(input.library || parsed?.library || DEFAULT_LIBRARY_NAME).trim() || DEFAULT_LIBRARY_NAME;
        return {
            site,
            folders: Array.isArray(input.folders) ? input.folders.filter(Boolean) : [],
            origin,
            sitePathPrefix,
            library
        };
    }

    function buildSitePath(siteOrConfig) {
        const config = typeof siteOrConfig === "object"
            ? siteOrConfig
            : getSiteConfig(siteOrConfig);
        const site = config?.site || siteOrConfig || "";
        const prefix = config?.sitePathPrefix || DEFAULT_SITE_PATH_PREFIX;
        return "/" + prefix + "/" + site;
    }

    function buildSiteUrl(siteOrConfig) {
        const config = typeof siteOrConfig === "object"
            ? siteOrConfig
            : getSiteConfig(siteOrConfig);
        const origin = config?.origin || TENANT_ORIGIN;
        return origin + buildSitePath(config || siteOrConfig);
    }

    function buildLibraryPath(siteOrConfig) {
        const config = typeof siteOrConfig === "object"
            ? siteOrConfig
            : getSiteConfig(siteOrConfig);
        const library = config?.library || DEFAULT_LIBRARY_NAME;
        return buildSitePath(config || siteOrConfig) + "/" + library;
    }

    function buildLibraryUrl(siteOrConfig) {
        const config = typeof siteOrConfig === "object"
            ? siteOrConfig
            : getSiteConfig(siteOrConfig);
        const origin = config?.origin || TENANT_ORIGIN;
        return origin + buildLibraryPath(config || siteOrConfig);
    }

    function getSiteConfig(site) {
        return siteConfigs.find(function (item) { return item.site === site; }) || null;
    }

    function ensureDefaultSite() {
        if (defaultSite && getSiteConfig(defaultSite)) return;
        defaultSite = siteConfigs[0]?.site || "";
    }

    function addSite() {
        const parsed = parseSiteLocation(siteInput.value);
        if (!parsed?.site) {
            showSettingsStatus("error", "❌ Invalid site name.");
            setTimeout(() => showSettingsStatus("", ""), 3000);
            return;
        }
        const site = parsed.site;
        if (getSiteConfig(site)) {
            defaultSite = site;
            siteInput.value = "";
            renderSiteList();
            renderPathList();
            updateSearchFilterOptions();
            updateSelectedLibraryHint();
            return;
        }

        siteConfigs.push({
            site,
            folders: [],
            origin: parsed.origin,
            sitePathPrefix: parsed.sitePathPrefix,
            library: parsed.library
        });
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
            row.title = buildLibraryUrl(siteConfig);
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
        siteConfigs = [{
            site: fallbackSite,
            folders,
            origin: TENANT_ORIGIN,
            sitePathPrefix: DEFAULT_SITE_PATH_PREFIX,
            library: DEFAULT_LIBRARY_NAME
        }];
        defaultSite = fallbackSite;
    }

    function parseSiteFromLibraryUrl(inputUrl) {
        return parseSiteLocation(inputUrl)?.site || "";
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
        const location = parseSiteLocation(window.location.href);
        if (!location?.site) {
            showSettingsStatus("error", "❌ Could not detect a SharePoint site from the current URL.");
            setTimeout(() => showSettingsStatus("", ""), 3000);
            return;
        }
        siteInput.value = location.origin + "/" + location.sitePathPrefix + "/" + location.site;
        siteInput.focus();
    }

    function normalizeSubfolderInput(rawInput, site) {
        const value = (rawInput || "").trim();
        if (!value) return "";

        let path = value;

        if (/^https?:\/\//i.test(value)) {
            path = extractSubfolderFromUrl(value, site);
        } else if (/^\/(sites|teams)\//i.test(value)) {
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
    const CONFIG_SYNC_KEYS = ["siteConfigs", "defaultSite", "syncEnabled"];

    function getConfigPayload() {
        return { siteConfigs, defaultSite };
    }

    function saveSiteConfigs() {
        chrome.storage.local.set({ ...getConfigPayload(), syncEnabled });
        if (syncEnabled) {
            chrome.storage.sync.set({ ...getConfigPayload(), syncEnabled: true }).catch(function (error) {
                console.warn("Could not sync SharePoint configuration:", error);
                showSettingsStatus("error", "❌ Could not sync configuration");
            });
        }
    }

    function saveUserPreferences() {
        chrome.storage.local.set({ favoriteFiles, recentFiles, recentSearches, savedSearches });
    }

    function getStoredConfig(data) {
        const configs = Array.isArray(data?.siteConfigs)
            ? data.siteConfigs.map(normalizeSiteConfig).filter(Boolean)
            : [];
        if (configs.length === 0) return null;

        const requestedDefault = normalizeSiteInput(data.defaultSite) || "";
        return {
            siteConfigs: configs,
            defaultSite: configs.some(item => item.site === requestedDefault)
                ? requestedDefault
                : configs[0].site
        };
    }

    async function toggleConfigSync() {
        const shouldEnable = syncConfigToggle.checked;
        const previousSyncEnabled = syncEnabled;
        syncConfigToggle.disabled = true;
        try {
            if (!shouldEnable) {
                syncEnabled = false;
                await chrome.storage.sync.set({ syncEnabled: false });
                await chrome.storage.local.set({ syncEnabled: false });
                showSettingsStatus("success", "🟢 Configuration sync disabled");
                return;
            }

            const syncedData = await chrome.storage.sync.get(CONFIG_SYNC_KEYS);
            const syncedConfig = getStoredConfig(syncedData);
            let selectedConfig = { ...getConfigPayload(), siteConfigs: siteConfigs.slice() };
            let selectedSource = "local";

            if (syncedConfig) {
                const choice = await promptSyncConflict(selectedConfig, syncedConfig);
                if (!choice) {
                    syncConfigToggle.checked = false;
                    showSettingsStatus("", "");
                    return;
                }

                if (choice === "profile") {
                    selectedSource = "profile";
                    selectedConfig = syncedConfig;
                    siteConfigs = selectedConfig.siteConfigs;
                    defaultSite = selectedConfig.defaultSite;
                    renderSiteList();
                    renderPathList();
                    updateSearchFilterOptions();
                    updateSelectedLibraryHint();
                }
            }

            syncEnabled = true;
            await chrome.storage.local.set({ ...selectedConfig, syncEnabled: true });
            await chrome.storage.sync.set({ ...selectedConfig, syncEnabled: true });
            showSettingsStatus("success", "🟢 Configuration sync enabled using " + selectedSource + " configuration");
        } catch (error) {
            syncEnabled = previousSyncEnabled;
            syncConfigToggle.checked = syncEnabled;
            chrome.storage.local.set({ syncEnabled });
            showSettingsStatus("error", "❌ Could not update configuration sync");
        } finally {
            syncConfigToggle.disabled = false;
        }
    }

    function exportConfig() {
        const payload = {
            schemaVersion: 1,
            exportedAt: new Date().toISOString(),
            ...getConfigPayload()
        };
        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "sharepoint-utils-config.json";
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 0);
        showSettingsStatus("success", "🟢 Configuration exported");
    }

    async function importConfigFromFile(event) {
        const file = event.target.files?.[0];
        event.target.value = "";
        if (!file) return;

        try {
            const payload = JSON.parse(await file.text());
            if (!Array.isArray(payload.siteConfigs) || payload.siteConfigs.length === 0) {
                throw new Error("The file does not contain any site configurations");
            }

            const importedConfigs = payload.siteConfigs
                .slice(0, 100)
                .map(normalizeSiteConfig)
                .filter(Boolean);
            if (importedConfigs.length === 0) {
                throw new Error("No valid site configurations were found");
            }

            siteConfigs = importedConfigs;
            defaultSite = normalizeSiteInput(payload.defaultSite) || "";
            ensureDefaultSite();
            saveSiteConfigs();
            renderSiteList();
            renderPathList();
            updateSearchFilterOptions();
            updateSelectedLibraryHint();
            showSettingsStatus("success", "🟢 Imported " + importedConfigs.length + " site configuration(s)");
        } catch (error) {
            showSettingsStatus("error", "❌ Import failed: " + escapeHtml(error.message));
        }
    }

    async function restoreState() {
        try {
            const data = await chrome.storage.local.get([
                "siteConfigs", "searchPaths", "configUrl", "defaultSite",
                "favoriteFiles", "recentFiles", "recentSearches", "savedSearches", "syncEnabled"
            ]);

            syncEnabled = data.syncEnabled === true;
            try {
                const syncedData = await chrome.storage.sync.get(CONFIG_SYNC_KEYS);
                syncEnabled = syncEnabled || syncedData.syncEnabled === true;
                if (syncEnabled && Array.isArray(syncedData.siteConfigs) && syncedData.siteConfigs.length > 0) {
                    data.siteConfigs = syncedData.siteConfigs;
                    data.defaultSite = syncedData.defaultSite;
                }
            } catch (error) {
                console.warn("Could not restore synced SharePoint configuration:", error);
            }

            favoriteFiles = Array.isArray(data.favoriteFiles)
                ? data.favoriteFiles.filter(file => file?.path && file?.name).slice(0, 50)
                : [];
            recentFiles = Array.isArray(data.recentFiles)
                ? data.recentFiles.filter(file => file?.path && file?.name).slice(0, 20)
                : [];
            recentSearches = Array.isArray(data.recentSearches)
                ? data.recentSearches.filter(Boolean).map(String).slice(0, 10)
                : [];
            savedSearches = Array.isArray(data.savedSearches)
                ? data.savedSearches.filter(Boolean).map(String).slice(0, 20)
                : [];

            if (Array.isArray(data.siteConfigs) && data.siteConfigs.length > 0) {
                siteConfigs = data.siteConfigs
                    .map(normalizeSiteConfig)
                    .filter(Boolean);
            } else {
                migrateLegacyState(data);
            }

            if (siteConfigs.length === 0) {
                siteConfigs = [{
                    site: DEFAULT_SITE,
                    folders: [],
                    origin: TENANT_ORIGIN,
                    sitePathPrefix: DEFAULT_SITE_PATH_PREFIX,
                    library: DEFAULT_LIBRARY_NAME
                }];
            }

            defaultSite = normalizeSiteInput(data.defaultSite) || "";
            ensureDefaultSite();
            saveSiteConfigs();
            syncConfigToggle.checked = syncEnabled;
            renderSiteList();
            renderPathList();
            updateSearchFilterOptions();
            updateSelectedLibraryHint();
            renderEmptyStateExtras();
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
                    showStatus("error", "⚠ Could not check connection");
                    return;
                }
                if (response && response.connected) {
                    showStatus("", "");
                } else {
                    showStatus("error", "⚠ No SharePoint tab found. Please login first.");
                }
            });
        } catch (e) {
            showStatus("error", "⚠ Could not check connection");
        }
    }

    async function detectCurrentPageFile() {
        const file = getCurrentPageFileCandidate();
        if (!file) {
            currentPageFile = null;
            renderCurrentPageFile();
            return;
        }

        // Render the URL-derived candidate immediately. Metadata enriches the
        // banner in the background and must not delay showing the file actions.
        currentPageFile = file;
        renderCurrentPageFile();

        try {
            const fileUrl = getCurrentPageFileDetailsUrl(file);
            if (fileUrl) {
                const fileDataResponse = await bgFetchWithTimeout(fileUrl, {
                    method: "GET",
                    headers: { "Accept": "application/json;odata=verbose" }
                }, CURRENT_FILE_METADATA_TIMEOUT_MS);
                const details = fileDataResponse?.d;
                if (details) {
                    file.name = details.Name || file.name;
                    file.path = details.ServerRelativeUrl || file.path;
                    file.size = parseInt(details.Length || 0, 10);
                    file.modified = details.TimeLastModified || "";
                    file.isCheckedOut = isFileCheckedOut(details.CheckOutType);
                    file.editor = details.ModifiedBy?.Title || "";
                    file.metadataLoaded = true;
                } else {
                    file.metadataError = true;
                }
            } else {
                file.metadataError = true;
            }
        } catch (e) {
            file.metadataError = true;
            console.warn("Could not load current file details:", e);
        }

        file.metadataLoading = false;
        if (currentPageFile === file) {
            renderCurrentPageFile();
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
            showStatus("loading", 'Searching SharePoint for "' + escapeHtml(query) + '"...');

            const selectedSite = siteFilterSelect.value;
            const searchPage = await searchFilesByName(query, selectedSite, 0);

            const currentQuery = searchInput.value.trim();
            if (currentQuery.length < MIN_QUERY_LENGTH) { showStatus("", ""); return; }

            lastSearchResults = searchPage.files;
            lastQuery = currentQuery;
            rememberRecentSearch(currentQuery);
            searchState = {
                query: currentQuery,
                selectedSite,
                sitePages: searchPage.sitePages,
                hasMore: searchPage.hasMore,
                errors: searchPage.errors
            };
            displayResults(getVisibleResults(), currentQuery);
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

    async function searchFilesByName(query, selectedSite, previousSitePages = []) {
        const rawQuery = String(query || "").replace(/["']/g, "").trim();
        const normalized = normalizeSearchInput(rawQuery);
        const searchTerms = normalized
            .split(/\s+/)
            .filter(Boolean);

        if (searchTerms.length === 0) {
            return { files: [], sitePages: [], hasMore: false, errors: [] };
        }

        let activeConfigs = siteConfigs;
        if (selectedSite) {
            activeConfigs = siteConfigs.filter(c => c.site === selectedSite);
        }

        // Exact logic from dist extension "title" mode
        const termClauses = searchTerms.map(function (term) {
            return "(Title:" + term + " OR Title:" + term + "* OR Filename:" + term + " OR Filename:" + term + "* OR FileLeafRef:" + term + " OR FileLeafRef:" + term + "*)";
        });
        const termQuery = termClauses.join(" ");
        const priorPages = Array.isArray(previousSitePages) ? previousSitePages : [];
        const requests = activeConfigs.map(function (siteConfig) {
            const previousPage = priorPages.find(page => page.site === siteConfig.site);
            const startRow = previousPage?.nextStartRow || 0;
            const queryText = termQuery + ' path:"' + buildLibraryUrl(siteConfig) + '"';
            return fetchSearchRows(queryText, buildSiteUrl(siteConfig), startRow)
                .then(page => ({ siteConfig, page }));
        });
        const settled = await Promise.allSettled(requests);
        const errors = [];
        const sitePages = [];
        const uniqueFiles = new Map();

        settled.forEach(function (result, index) {
            if (result.status === "rejected") {
                errors.push({ site: activeConfigs[index].site, error: result.reason?.message || "Unknown error" });
                return;
            }
            const siteConfig = result.value.siteConfig;
            const page = result.value.page;
            sitePages.push({
                site: siteConfig.site,
                nextStartRow: page.nextStartRow,
                hasMore: page.hasMore
            });

            page.rows.forEach(function (row) {
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

            const matchedSiteConfig = getSiteConfigByServerRelativePath(relativePath) || siteConfig;
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
                editor: getVal("EditorOWSUSER") || "",
                isCheckedOut: false
            });
            });
        });

        if (activeConfigs.length > 0 && errors.length === activeConfigs.length) {
            throw new Error(errors[0].error);
        }

        return {
            files: Array.from(uniqueFiles.values()),
            sitePages,
            hasMore: sitePages.some(page => page.hasMore),
            errors
        };
    }

    async function fetchSearchRows(queryText, apiBase, startRow = 0) {
        const rowLimit = SEARCH_PAGE_SIZE;
        const baseUrl = apiBase || TENANT_ORIGIN;
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

        const rows = searchData?.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows?.results || [];
        return {
            rows,
            nextStartRow: startRow + rows.length,
            hasMore: rows.length >= rowLimit
        };
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

    function bgFetchWithTimeout(url, options = {}, timeoutMs = CURRENT_FILE_METADATA_TIMEOUT_MS) {
        let timeoutId = null;
        const timeout = new Promise((resolve, reject) => {
            timeoutId = setTimeout(() => reject(new Error("REQUEST_TIMEOUT")), timeoutMs);
        });

        return Promise.race([bgFetch(url, options), timeout]).finally(() => {
            if (timeoutId) clearTimeout(timeoutId);
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
        const candidatePath = String(serverRelativePath || "");
        const lowerPath = candidatePath.toLowerCase();
        return siteConfigs.find(function (siteConfig) {
            const libraryPath = buildLibraryPath(siteConfig);
            const lowerLibraryPath = libraryPath.toLowerCase();
            return lowerPath === lowerLibraryPath || lowerPath.startsWith(lowerLibraryPath + "/");
        }) || null;
    }

    // ============================================
    // Display Results
    // ============================================
    function showEmptyState() {
        resultsEl.innerHTML = '';
        const el = document.createElement('div');
        el.className = 'sp-search-empty-state';
        el.id = 'sp-search-empty-state';
        el.innerHTML = '<div class="sp-search-empty-icon">🔍</div>' +
            '<div class="sp-search-empty-title">Search for a file</div>' +
            '<div class="sp-search-empty-hint">Type a filename and press <kbd>Enter</kbd> to search across your SharePoint sites</div>' +
            '<div class="sp-empty-extras" id="sp-empty-extras"></div>';
        resultsEl.appendChild(el);
        renderEmptyStateExtras();
    }
    function hideEmptyState() {
        const el = document.getElementById('sp-search-empty-state');
        if (el) el.remove();
    }

    function getFavoriteKey(file) {
        return String(file?.siteUrl || "").toLowerCase() + "|" + String(file?.path || "").toLowerCase();
    }

    function toStoredFile(file, url) {
        return {
            name: String(file?.name || "File"),
            path: String(file?.path || ""),
            site: String(file?.site || ""),
            siteUrl: String(file?.siteUrl || ""),
            url: String(url || getViewFileUrl(file?.path || "", file?.siteUrl || "")),
            size: Number(file?.size) || 0,
            modified: String(file?.modified || "")
        };
    }

    function isFavoriteFile(file) {
        const key = getFavoriteKey(file);
        return favoriteFiles.some(item => getFavoriteKey(item) === key);
    }

    function updateFavoriteButton(button, file) {
        if (!button) return;
        const favorite = isFavoriteFile(file);
        button.textContent = favorite ? "⭐" : "☆";
        button.title = favorite ? "Remove from favorites" : "Add to favorites";
        button.setAttribute("aria-label", button.title);
    }

    function toggleFavorite(file, button) {
        const key = getFavoriteKey(file);
        const existingIndex = favoriteFiles.findIndex(item => getFavoriteKey(item) === key);
        if (existingIndex >= 0) {
            favoriteFiles.splice(existingIndex, 1);
            showStatus("success", '🟢 Removed "' + escapeHtml(file.name) + '" from favorites');
        } else {
            favoriteFiles.unshift(toStoredFile(file, getViewFileUrl(file.path, file.siteUrl)));
            favoriteFiles = favoriteFiles.slice(0, 50);
            showStatus("success", '🟢 Added "' + escapeHtml(file.name) + '" to favorites');
        }
        saveUserPreferences();
        updateFavoriteButton(button, file);
        renderEmptyStateExtras();
    }

    function rememberRecentFile(file, url) {
        const storedFile = toStoredFile(file, url);
        const key = getFavoriteKey(storedFile);
        recentFiles = [storedFile].concat(recentFiles.filter(item => getFavoriteKey(item) !== key)).slice(0, 20);
        saveUserPreferences();
    }

    function rememberRecentSearch(query) {
        const value = String(query || "").trim();
        if (!value) return;
        recentSearches = [value].concat(recentSearches.filter(item => item.toLowerCase() !== value.toLowerCase())).slice(0, 10);
        saveUserPreferences();
    }

    function saveCurrentSearch() {
        const value = searchInput.value.trim();
        if (value.length < MIN_QUERY_LENGTH) {
            showStatus("error", "❌ Enter at least " + MIN_QUERY_LENGTH + " characters before saving a search");
            return;
        }
        savedSearches = [value].concat(savedSearches.filter(item => item.toLowerCase() !== value.toLowerCase())).slice(0, 20);
        saveUserPreferences();
        renderEmptyStateExtras();
        showStatus("success", '🟢 Saved search "' + escapeHtml(value) + '"');
    }

    function createEmptyStateButton(label, className, onClick) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = className;
        button.textContent = label;
        button.addEventListener("click", onClick);
        return button;
    }

    function renderEmptyStateExtras() {
        const extras = document.getElementById("sp-empty-extras");
        if (!extras) return;
        extras.innerHTML = "";

        const addSection = function (title, items, className, onClick, removeAction) {
            if (!items.length) return;
            const section = document.createElement("div");
            section.className = "sp-empty-section";
            const heading = document.createElement("div");
            heading.className = "sp-empty-section-title";
            heading.textContent = title;
            section.appendChild(heading);
            const list = document.createElement("div");
            list.className = "sp-empty-section-list";
            items.forEach(function (item, index) {
                const row = document.createElement("div");
                row.className = "sp-empty-item-row";
                const button = createEmptyStateButton(
                    typeof item === "string" ? item : item.name,
                    className,
                    () => onClick(item)
                );
                row.appendChild(button);
                if (removeAction) {
                    const removeButton = createEmptyStateButton("✕", "sp-empty-item-remove", () => removeAction(item, index));
                    removeButton.title = "Remove";
                    row.appendChild(removeButton);
                }
                list.appendChild(row);
            });
            section.appendChild(list);
            extras.appendChild(section);
        };

        addSection("Recent searches", recentSearches, "sp-empty-item-btn", function (query) {
            searchInput.value = query;
            clearBtn.classList.add("sp-visible");
            handleSearch();
        });
        addSection("Saved searches", savedSearches, "sp-empty-item-btn", function (query) {
            searchInput.value = query;
            clearBtn.classList.add("sp-visible");
            handleSearch();
        }, function (query) {
            savedSearches = savedSearches.filter(item => item !== query);
            saveUserPreferences();
            renderEmptyStateExtras();
        });
        addSection("Recent files", recentFiles, "sp-empty-item-btn", function (file) {
            const url = file.url || getViewFileUrl(file.path);
            if (url) window.open(url, "_blank");
        });
        addSection("Favorites", favoriteFiles, "sp-empty-item-btn", function (file) {
            const url = file.url || getViewFileUrl(file.path);
            if (url) window.open(url, "_blank");
        });
    }

    function displayResults(files, query) {
        hideEmptyState();
        resultsEl.innerHTML = "";

        if (files.length === 0) {
            showStatus("", "");
            resultsEl.innerHTML =
                '<div class="sp-no-results">' +
                '<div class="sp-no-results-icon">🔍</div>' +
                '<div class="sp-no-results-text">No files found matching "<strong>' +
                escapeHtml(query) + '</strong>"</div></div>';
            appendLoadMoreButton();
            return;
        }

        const failedSiteCount = searchState?.errors?.length || 0;
        const partialMessage = failedSiteCount
            ? ' ⚠ ' + failedSiteCount + ' site(s) could not be searched.'
            : '';
        showStatus("success", '🟢 Found ' + files.length + ' file(s) matching "' + escapeHtml(query) + '"' + partialMessage);

        files.forEach(function (file) {
            const viewFileUrl = getViewFileUrl(file.path, file.siteUrl);

            const item = document.createElement("div");
            item.className = "sp-file-item";
            item.title = "Open: " + file.name;
            item.style.cursor = "pointer";
            item.setAttribute("role", "button");
            item.setAttribute("tabindex", "0");
            const openResult = function () {
                rememberRecentFile(file, viewFileUrl);
                window.open(viewFileUrl, "_blank");
            };
            item.addEventListener("click", openResult);
            item.addEventListener("keydown", function (event) {
                if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    openResult();
                }
            });

            item.innerHTML =
                '<span class="sp-file-icon">' + getFileIcon(file.name) + '</span>' +
                '<div class="sp-file-info">' +
                '<div class="sp-file-name">' + highlightText(escapeHtml(file.name), query) + '</div>' +
                '<div class="sp-file-path"></div>' +
                '<div class="sp-file-meta">' +
                (file.size ? '<span class="sp-file-size">💾 ' + formatFileSize(file.size) + '</span>' : '') +
                (file.editor ? '<span class="sp-file-editor">👤 ' + escapeHtml(formatEditorName(file.editor)) + '</span>' : '') +
                (file.modified ? '<span class="sp-file-modified">🕒 ' + escapeHtml(formatDate(file.modified)) + '</span>' : '') +
                '</div>' +
                '</div>' +
                '<div class="sp-file-actions">' +
                (canOpenInDesktopApp(file) ? '<button type="button" class="sp-action-btn sp-action-desktop" title="Open in desktop app" aria-label="Open in desktop app">🖥️</button>' : '') +
                '<button type="button" class="sp-action-btn sp-action-checkout" title="' + (file.isCheckedOut ? 'Check in file' : 'Check out file') + '" aria-label="' + (file.isCheckedOut ? 'Check in file' : 'Check out file') + '">' + getCheckoutActionIcon(file.isCheckedOut) + '</button>' +
                '<button type="button" class="sp-action-btn sp-action-share" title="Copy share link" aria-label="Copy share link">🔗</button>' +
                '<button type="button" class="sp-action-btn sp-action-favorite" title="' + (isFavoriteFile(file) ? 'Remove from favorites' : 'Add to favorites') + '" aria-label="' + (isFavoriteFile(file) ? 'Remove from favorites' : 'Add to favorites') + '">' + (isFavoriteFile(file) ? "⭐" : "☆") + '</button>' +
                '</div>';

            item.querySelector(".sp-action-share").addEventListener("click", async (e) => {
                e.stopPropagation();
                openShareModal(file, viewFileUrl);
            });

            const favoriteBtn = item.querySelector(".sp-action-favorite");
            favoriteBtn.addEventListener("click", (e) => {
                e.stopPropagation();
                toggleFavorite(file, favoriteBtn);
            });
            updateFavoriteButton(favoriteBtn, file);

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

        appendLoadMoreButton();
    }

    function appendLoadMoreButton() {
        if (!searchState?.hasMore) return;
        const loadMoreBtn = document.createElement("button");
        loadMoreBtn.type = "button";
        loadMoreBtn.className = "sp-load-more-btn";
        loadMoreBtn.textContent = "Load more results";
        loadMoreBtn.addEventListener("click", loadMoreResults);
        resultsEl.appendChild(loadMoreBtn);
    }

    async function loadMoreResults() {
        if (isLoading || !searchState?.hasMore || !lastQuery) return;

        const queryAtStart = searchInput.value.trim();
        if (queryAtStart !== searchState.query) return;

        isLoading = true;
        showStatus("loading", 'Loading more results for "' + escapeHtml(searchState.query) + '"...');
        try {
            const searchPage = await searchFilesByName(
                searchState.query,
                searchState.selectedSite,
                searchState.sitePages
            );
            const existingPaths = new Set(lastSearchResults.map(file => file.path.toLowerCase()));
            searchPage.files.forEach(file => {
                if (!existingPaths.has(file.path.toLowerCase())) {
                    lastSearchResults.push(file);
                    existingPaths.add(file.path.toLowerCase());
                }
            });
            searchState.sitePages = searchPage.sitePages;
            searchState.hasMore = searchPage.hasMore;
            searchState.errors = (searchState.errors || []).concat(searchPage.errors || []);
            displayResults(getVisibleResults(), searchState.query);
        } catch (error) {
            console.error("Load more error:", error);
            showStatus("error", "❌ Could not load more results: " + error.message);
        } finally {
            isLoading = false;
        }
    }

    async function toggleFileCheckout(file, buttonEl) {
        if (!buttonEl || buttonEl.disabled) return;

        const originalLabel = buttonEl.textContent;
        buttonEl.disabled = true;
        buttonEl.textContent = "⏳";

        try {
            const nextAction = file.isCheckedOut ? "checkIn" : "checkOut";
            if (nextAction === "checkOut") {
                await sharePointCheckOut(file);
                file.isCheckedOut = true;
                buttonEl.textContent = getCheckoutActionIcon(true);
                buttonEl.title = "Check in file";
                buttonEl.setAttribute("aria-label", "Check in file");
                showStatus("success", '🟢 Checked out "' + escapeHtml(file.name) + '"');
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
                buttonEl.setAttribute("aria-label", "Check out file");
                showStatus("success", '🟢 Checked in "' + escapeHtml(file.name) + '"');
            }
            if (currentPageFile && currentPageFile.path === file.path) {
                currentPageFile.isCheckedOut = file.isCheckedOut;
                renderCurrentPageFile();
            }
        } catch (error) {
            buttonEl.textContent = originalLabel;
            buttonEl.title = file.isCheckedOut ? "Check in file" : "Check out file";
            buttonEl.setAttribute("aria-label", buttonEl.title);
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
        inlineModalBackdrop.setAttribute("aria-hidden", "false");
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
        inlineModalBackdrop.setAttribute("aria-hidden", "true");
        inlineModalTextarea.value = "";
        if (resolver) {
            resolver(submitted ? value : null);
        }
    }

    function promptSyncConflict(localConfig, syncedConfig) {
        syncModalSubtitleEl.textContent = "Local: " + localConfig.siteConfigs.length + " site(s) · Chrome profile: " + syncedConfig.siteConfigs.length + " site(s)";
        syncModalBackdrop.classList.remove("sp-hidden");
        syncModalBackdrop.setAttribute("aria-hidden", "false");
        activeSyncConflict = {};

        setTimeout(() => syncUseLocalBtn.focus(), 0);

        return new Promise((resolve) => {
            activeSyncConflict.resolve = resolve;
        });
    }

    function closeSyncConflict(choice) {
        if (!activeSyncConflict) return;
        const resolver = activeSyncConflict.resolve;
        activeSyncConflict = null;
        syncModalBackdrop.classList.add("sp-hidden");
        syncModalBackdrop.setAttribute("aria-hidden", "true");
        if (resolver) resolver(choice);
    }

    function openShareModal(file, url) {
        activeShareFile = file;
        activeShareUrl = url;
        shareModalFilename.textContent = file.name;
        shareModalBackdrop.classList.remove("sp-hidden");
        shareModalBackdrop.setAttribute("aria-hidden", "false");
    }

    function closeShareModal() {
        activeShareFile = null;
        activeShareUrl = null;
        shareModalBackdrop.classList.add("sp-hidden");
        shareModalBackdrop.setAttribute("aria-hidden", "true");
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
        if (currentPageFile.path) {
            currentFilePathEl.appendChild(buildFolderBreadcrumb(currentPageFile));
        } else {
            currentFilePathEl.textContent = currentPageFile.metadataLoading
                ? "Loading file details..."
                : "File path unavailable";
        }
        
        currentFileMetaEl.innerHTML = 
            (currentPageFile.size ? '<span class="sp-file-size">💾 ' + formatFileSize(currentPageFile.size) + '</span>' : '') +
            (currentPageFile.editor ? '<span class="sp-file-editor">👤 ' + escapeHtml(formatEditorName(currentPageFile.editor)) + '</span>' : '') +
            (currentPageFile.modified ? '<span class="sp-file-modified">🕒 ' + escapeHtml(formatDate(currentPageFile.modified)) + '</span>' : '') +
            (currentPageFile.metadataLoading && !currentPageFile.size && !currentPageFile.modified
                ? '<span class="sp-file-modified">⏳ Loading file details...</span>'
                : '') +
            (currentPageFile.metadataError ? '<span class="sp-file-modified">⚠ File details unavailable</span>' : '');

        currentFileDesktopBtn.classList.toggle("sp-hidden", !canOpenInDesktopApp(currentPageFile));
        currentFileShareBtn.disabled = !getCurrentPageFileShareUrl(currentPageFile);
        currentFileCheckoutBtn.disabled = !currentPageFile.path || !currentPageFile.metadataLoaded;
        currentFileCheckoutBtn.textContent = getCheckoutActionIcon(currentPageFile.isCheckedOut);
        currentFileCheckoutBtn.title = currentPageFile.metadataLoaded && currentPageFile.path
            ? (currentPageFile.isCheckedOut ? "Check in file" : "Check out file")
            : currentPageFile.metadataLoading ? "Loading file details..." : "File details unavailable";
        currentFileCheckoutBtn.setAttribute("aria-label", currentFileCheckoutBtn.title);
        currentFileEl.classList.remove("sp-hidden");
    }

    // ============================================
    // Filter Helpers
    // ============================================
    function applyFilter(files) {
        const selectedSite = siteFilterSelect.value;
        const selectedFolderKey = folderFilterSelect.value;
        const selectedType = fileTypeFilterSelect.value;
        const modifiedAfter = modifiedAfterInput.value
            ? new Date(modifiedAfterInput.value + "T00:00:00")
            : null;

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
            if (modifiedAfter && !Number.isNaN(modifiedAfter.getTime())) {
                const fileModified = new Date(file.modified || "");
                if (Number.isNaN(fileModified.getTime()) || fileModified < modifiedAfter) {
                    return false;
                }
            }
            return true;
        });
    }

    function sortResults(files) {
        const sortMode = sortFilterSelect.value || "relevance";
        if (sortMode === "relevance") return files;

        return files.slice().sort(function (a, b) {
            if (sortMode === "nameAsc" || sortMode === "nameDesc") {
                const direction = sortMode === "nameAsc" ? 1 : -1;
                return direction * String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base" });
            }

            if (sortMode === "sizeAsc" || sortMode === "sizeDesc") {
                const direction = sortMode === "sizeAsc" ? 1 : -1;
                return direction * ((Number(a.size) || 0) - (Number(b.size) || 0));
            }

            const aTime = new Date(a.modified || "").getTime();
            const bTime = new Date(b.modified || "").getTime();
            const safeATime = Number.isNaN(aTime) ? 0 : aTime;
            const safeBTime = Number.isNaN(bTime) ? 0 : bTime;
            const direction = sortMode === "modifiedAsc" ? 1 : -1;
            return direction * (safeATime - safeBTime);
        });
    }

    function getVisibleResults() {
        return sortResults(applyFilter(lastSearchResults));
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
    function getViewFileUrl(serverRelativePath, siteUrl = "") {
        return getFileWebUrl(serverRelativePath, siteUrl) + "?web=1";
    }

    function getCurrentPageFileShareUrl(file) {
        if (file?.path) {
            return getViewFileUrl(file.path, file.siteUrl);
        }
        // A sourcedoc-only viewer URL is still a useful share target while
        // SharePoint metadata is loading or unavailable.
        return window.location.href;
    }

    // ============================================
    // Utility Functions
    // ============================================
    let statusTimeout = null;
    function showStatus(type, message) {
        if (statusTimeout) {
            clearTimeout(statusTimeout);
            statusTimeout = null;
        }

        if (!type && !message) { 
            connectionEl.innerHTML = ""; 
            connectionEl.style.color = "var(--sp-text-secondary)";
            return; 
        }
        
        let color = "var(--sp-text-secondary)";
        if (type === "error") color = "#f87171";
        else if (type === "success") color = "#4ade80";
        else if (type === "loading") color = "var(--sp-text-primary)";
        
        connectionEl.style.color = color;
        connectionEl.innerHTML = type === "loading"
            ? '<span class="sp-spinner" style="margin-right: 8px; width: 10px; height: 10px; border-width: 2px;"></span>' + message
            : message;
            
        if (type !== "loading" && message) {
            statusTimeout = setTimeout(() => {
                connectionEl.innerHTML = "";
                connectionEl.style.color = "var(--sp-text-secondary)";
            }, 5000);
        }
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

    function getCheckoutActionIcon(isCheckedOut) {
        return isCheckedOut ? CHECKIN_ICON : CHECKOUT_ICON;
    }

    function getCurrentPageFileCandidate() {
        const currentUrl = new URL(window.location.href);
        const rawIdPath = currentUrl.searchParams.get("id");
        const rawFileName = currentUrl.searchParams.get("file");
        const rawSourceDoc = currentUrl.searchParams.get("sourcedoc");
        const serverRelativePath = decodeSharePointPath(rawIdPath) || getCurrentPageServerRelativePath(currentUrl);
        const sourceDocId = normalizeSourceDocId(rawSourceDoc);
        const fileName = decodeSharePointFileName(rawFileName) ||
            (serverRelativePath.split("/").pop() || "") ||
            getFileNameFromDocumentTitle();
        const siteUrl = getCurrentPageSiteUrl(currentUrl, serverRelativePath);
        if ((!fileName && !sourceDocId) || (!sourceDocId && !isCurrentFileTypeSupported(fileName))) {
            return null;
        }
        return {
            name: fileName || "Current file",
            path: serverRelativePath,
            size: 0,
            modified: "",
            isCheckedOut: false,
            sourceDocId,
            siteUrl,
            metadataLoading: true,
            metadataLoaded: false,
            metadataError: false
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
        const match = String(serverRelativePath).match(/^\/(sites|teams)\/([^/]+)/i);
        return match ? match[2] : "";
    }

    function getFolderPath(serverRelativePath, fileName) {
        return String(serverRelativePath || "").replace("/" + (fileName || ""), "");
    }

    function buildFolderBreadcrumb(file) {
        const wrapper = document.createElement("div");
        const serverRelativePath = file?.path || "";
        const site = file?.site || getSiteConfigByServerRelativePath(serverRelativePath)?.site || getSiteSlugFromServerRelativePath(serverRelativePath);
        const siteUrl = file?.siteUrl || "";
        const libraryPath = site ? buildLibraryPath(site) : "";
        const parentFolderPath = getFolderPath(serverRelativePath, file?.name || "");
        if (!libraryPath || !parentFolderPath.startsWith(libraryPath)) {
            wrapper.textContent = parentFolderPath || "/";
            return wrapper;
        }
        const relativePath = libraryPath && parentFolderPath.startsWith(libraryPath)
            ? parentFolderPath.slice(libraryPath.length)
            : "";
        const segments = relativePath.split("/").filter(Boolean);

        if (segments.length === 0) {
            wrapper.appendChild(createFolderLink("/", getFolderUrlFromFolderPath(parentFolderPath, siteUrl)));
            return wrapper;
        }

        let currentPath = "";
        segments.forEach(function (segment, index) {
            currentPath += "/" + segment;
            if (index > 0) {
                wrapper.appendChild(document.createTextNode(" / "));
            }
            wrapper.appendChild(createFolderLink(segment, getFolderUrlFromFolderPath(libraryPath + currentPath, siteUrl)));
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

    function getFolderUrlFromFolderPath(folderPath, siteUrl = "") {
        let libraryRootUrl = getLibraryRootUrlByPath(folderPath);
        if (!libraryRootUrl && siteUrl) {
            try {
                const parsedSiteUrl = new URL(siteUrl);
                const sitePath = parsedSiteUrl.pathname.replace(/\/$/, "");
                const relativeLibrary = folderPath.startsWith(sitePath + "/")
                    ? folderPath.slice(sitePath.length + 1).split("/")[0]
                    : "";
                if (relativeLibrary) {
                    libraryRootUrl = parsedSiteUrl.origin + sitePath + "/" + relativeLibrary;
                }
            } catch (e) {
                // Use the configured library URL when available.
            }
        }
        return libraryRootUrl + "/Forms/AllItems.aspx?id=" + encodeURIComponent(folderPath);
    }

    function getCurrentPageFileDetailsUrl(file) {
        const select = "?$select=Name,ServerRelativeUrl,Length,TimeLastModified,CheckOutType,ModifiedBy/Title&$expand=ModifiedBy";
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

    function decodeSharePointPath(rawPath) {
        if (!rawPath) return "";
        try {
            const decoded = decodeURIComponent(rawPath);
            try {
                const absoluteUrl = new URL(decoded);
                return decodeURIComponent(absoluteUrl.pathname);
            } catch (e) {
                // The value is a relative server path, which is expected here.
            }
            return decoded;
        } catch (e) {
            return rawPath;
        }
    }

    function getCurrentPageServerRelativePath(currentUrl) {
        const path = decodeSharePointPath(currentUrl.pathname);
        const match = path.match(/\/(sites|teams)\/[^/]+\/.*\.(?:docx?|xlsx?|xlsm|xlsb|pptx?|pptm)$/i);
        return match ? match[0] : "";
    }

    function getFileNameFromDocumentTitle() {
        const title = String(document.title || "");
        const match = title.match(/[^|\u2013\u2014-]+\.(?:docx?|xlsx?|xlsm|xlsb|pptx?|pptm)\b/i);
        return match ? match[0].trim() : "";
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
        if (match) return currentUrl.origin + match[0];

        const pageContextUrl = window._spPageContextInfo?.webAbsoluteUrl;
        if (pageContextUrl) {
            try {
                const contextUrl = new URL(pageContextUrl, currentUrl.origin);
                if (contextUrl.origin === currentUrl.origin) {
                    return contextUrl.href.replace(/\/$/, "");
                }
            } catch (e) {
                // Fall through to the configured site when page context is invalid.
            }
        }

        const configuredSite = defaultSite || siteConfigs[0]?.site;
        return configuredSite ? buildSiteUrl(configuredSite) : "";
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

    function getRawFileUrl(file) {
        if (!file || !file.path) return "";
        return getFileWebUrl(file.path, file.siteUrl);
    }

    function getFileWebUrl(serverRelativePath, siteUrl = "") {
        const path = String(serverRelativePath || "");
        if (!path) return "";

        const siteConfig = getSiteConfigByServerRelativePath(path);
        const libraryRootUrl = getLibraryRootUrlByPath(path);
        if (libraryRootUrl && siteConfig) {
            const relativePath = path.slice(buildLibraryPath(siteConfig.site).length);
            return libraryRootUrl + relativePath;
        }

        try {
            const currentSiteUrl = siteUrl ? new URL(siteUrl) : null;
            return (currentSiteUrl ? currentSiteUrl.origin : TENANT_ORIGIN) + path;
        } catch (e) {
            return path;
        }
    }

    function getLibraryRootUrlByPath(serverRelativePath) {
        const siteConfig = getSiteConfigByServerRelativePath(serverRelativePath);
        const site = siteConfig?.site || getSiteSlugFromServerRelativePath(serverRelativePath);
        return site ? buildLibraryUrl(site) : "";
    }

    function highlightText(text, query) {
        if (!query) return text;
        var regex = new RegExp("(" + escapeRegex(query) + ")", "gi");
        return text.replace(regex, '<span class="sp-highlight">$1</span>');
    }

})();
