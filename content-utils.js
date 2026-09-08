// ============================================
// SharePoint Utils - Shared, testable helpers
// Loaded before content.js by the service worker.
// ============================================
(function (root) {
    const FILE_ICONS_SVG = {
        word: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32"><rect width="32" height="32" rx="4" fill="#2B579A"/><path d="M7 8h12l6 6v10a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V10a2 2 0 0 1 2-2z" fill="#fff" opacity=".15"/><text x="16" y="22" text-anchor="middle" font-family="Arial,sans-serif" font-weight="700" font-size="13" fill="#fff">W</text></svg>`,
        excel: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32"><rect width="32" height="32" rx="4" fill="#217346"/><path d="M7 8h12l6 6v10a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V10a2 2 0 0 1 2-2z" fill="#fff" opacity=".15"/><text x="16" y="22" text-anchor="middle" font-family="Arial,sans-serif" font-weight="700" font-size="13" fill="#fff">X</text></svg>`,
        ppt: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32"><rect width="32" height="32" rx="4" fill="#D24726"/><path d="M7 8h12l6 6v10a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V10a2 2 0 0 1 2-2z" fill="#fff" opacity=".15"/><text x="16" y="22" text-anchor="middle" font-family="Arial,sans-serif" font-weight="700" font-size="13" fill="#fff">P</text></svg>`,
        pdf: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32"><rect width="32" height="32" rx="4" fill="#E33B2E"/><path d="M7 8h12l6 6v10a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V10a2 2 0 0 1 2-2z" fill="#fff" opacity=".15"/><text x="16" y="22" text-anchor="middle" font-family="Arial,sans-serif" font-weight="700" font-size="11" fill="#fff">PDF</text></svg>`
    };

    function normalizeSearchInput(value) {
        return String(value || "")
            .replace(/[()（）{}[\]【】「」『』.·。_\-]/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    function sanitizeSearchTerm(term) {
        return String(term || "").replace(/["'*]/g, "").trim();
    }

    function matchesAllSearchTerms(text, searchTerms) {
        const value = String(text || "").normalize("NFKC").toLowerCase();
        return searchTerms.every(function (term) {
            return value.includes(String(term || "").normalize("NFKC").toLowerCase());
        });
    }

    function encodeSharePointPath(path) {
        return root.encodeURIComponent(String(path || "").replace(/'/g, "''"));
    }

    function encodeODataString(value) {
        return root.encodeURIComponent(String(value || "").replace(/'/g, "''"));
    }

    function isFileCheckedOut(checkOutType) {
        const value = parseInt(checkOutType, 10);
        if (Number.isNaN(value)) return false;
        return value !== 2;
    }

    function getOfficeDesktopProtocol(fileName) {
        const ext = String(fileName || "").split(".").pop().toLowerCase();
        if (ext === "doc" || ext === "docx") return "ms-word:ofe|u|";
        if (["xls", "xlsx", "xlsm", "xlsb"].includes(ext)) return "ms-excel:ofe|u|";
        if (["ppt", "pptx", "pptm"].includes(ext)) return "ms-powerpoint:ofe|u|";
        return "";
    }

    function getFileTypeCategory(fileName) {
        const ext = String(fileName || "").split(".").pop().toLowerCase();
        if (["xls", "xlsx", "xlsm", "xlsb", "csv"].includes(ext)) return "excel";
        if (["doc", "docx", "dot", "dotx"].includes(ext)) return "word";
        if (["ppt", "pptx", "pptm", "pps", "ppsx"].includes(ext)) return "powerpoint";
        if (ext === "pdf") return "pdf";
        return "other";
    }

    function getFileIcon(filename) {
        const ext = String(filename || "").split(".").pop().toLowerCase();
        if (["doc", "docx", "dot", "dotx"].includes(ext)) return FILE_ICONS_SVG.word;
        if (["xls", "xlsx", "xlsm", "xlsb", "csv"].includes(ext)) return FILE_ICONS_SVG.excel;
        if (["ppt", "pptx", "pptm", "pps", "ppsx"].includes(ext)) return FILE_ICONS_SVG.ppt;
        if (ext === "pdf") return FILE_ICONS_SVG.pdf;
        const emojiIcons = {
            txt: "📄", md: "📄", json: "📄", xml: "📄",
            jpg: "🖼️", jpeg: "🖼️", png: "🖼️", gif: "🖼️", svg: "🖼️",
            mp4: "🎬", avi: "🎬", mov: "🎬",
            mp3: "🎵", wav: "🎵",
            zip: "📦", rar: "📦",
            html: "🌐", css: "🎨", js: "⚡",
            py: "🐍", java: "☕", sql: "🗃️"
        };
        return emojiIcons[ext] || "📄";
    }

    function formatFileSize(bytes) {
        if (!bytes || bytes === 0) return "";
        const units = ["B", "KB", "MB", "GB"];
        let index = 0;
        let size = Number(bytes);
        while (size >= 1024 && index < units.length - 1) { size /= 1024; index++; }
        return size.toFixed(1) + " " + units[index];
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    function escapeRegex(value) {
        return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    function formatEditorName(raw) {
        if (!raw) return "";
        const parts = String(raw).split("|");
        for (const part of parts) {
            const trimmed = part.trim();
            if (trimmed && !trimmed.includes("@") && !trimmed.includes("i:0#.f")) return trimmed;
        }
        return parts.length > 1 ? parts[1].trim() : parts[0].trim();
    }

    function formatDate(dateString) {
        if (!dateString) return "";
        try {
            const date = new Date(dateString);
            if (Number.isNaN(date.getTime())) return dateString;
            return date.toLocaleDateString() + " " + date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        } catch (e) {
            return dateString;
        }
    }

    root.SharePointUtils = Object.freeze({
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
        matchesAllSearchTerms,
        normalizeSearchInput,
        sanitizeSearchTerm
    });
})(globalThis);
