#!/usr/bin/env python3
"""
Multi-File HTML Templates for Microsoft Teams Chat Export

This module contains templates for the multi-file (SPA) export structure:
- index.html: Main navigation hub with content area
- stats.html: Statistics page loaded in content area
- chats/*.html: Individual chat files loaded in content area
- assets/: Shared CSS, JS, and search index

Features:
- Responsive sidebar navigation (persistent)
- Dynamic content loading via iframes
- Shared search functionality across all files
- Unified styling and theme support
- Statistics dashboard

Author: Alexander Wegner
Version: v0.1.8
Last Updated: 2026-01-15
"""

# =============================================================================
# SHARED CSS (assets/style.css)
# =============================================================================

SHARED_CSS = r"""
/* === THEME VARIABLES === */
:root {
    --bg: #f4f6f8;
    --text: #0f172a;
    --sidebar-bg: #152238;
    --sidebar-border: #2a3a56;
    --sidebar-hover: #22324d;
    --sidebar-active: #2b3f5f;
    --sidebar-link-bg: #1a2a43;
    --sidebar-header-bg: #1e2f4a;
    --sidebar-team-bg: #22334f;
    --sidebar-channel-bg: #243754;
    --sidebar-text: #edf3fb;
    --sidebar-muted: #b7c6db;
    --content-bg: #f4f6f8;
    --message-mine: #d9f5fb;
    --message-theirs: #e5e7ef;
    --meta: #4b5563;
    --nav-btn-bg: #0ea5e9;
    --nav-btn-hover: #22d3ee;
    --date-separator-bg: #e8ecf3;
    --date-separator-text: #4b5563;
    --input-bg: #ffffff;
    --input-border: #d1d5db;
    --input-text: #0f172a;
    --highlight-bg: #bff4ff;
    --highlight-text: #0f172a;
    /* Category accents */
    --cat-one: #06b6d4;      /* cyan */
    --cat-group: #8b5cf6;    /* violet */
    --cat-meeting: #f59e0b;  /* amber */
    --cat-channel: #22c55e;  /* green */
}

body.dark {
    --bg: #0b1017;
    --text: #e8edf5;
    --sidebar-bg: #122034;
    --sidebar-border: #243650;
    --sidebar-hover: #1c2f49;
    --sidebar-active: #294161;
    --sidebar-link-bg: #16283f;
    --sidebar-header-bg: #1b2c45;
    --sidebar-team-bg: #223650;
    --sidebar-channel-bg: #27405e;
    --sidebar-text: #e8edf5;
    --sidebar-muted: #9ca3b5;
    --content-bg: #0b1017;
    --message-mine: #0f2c36;
    --message-theirs: #122032;
    --meta: #9ca3b5;
    --nav-btn-bg: #12495c;
    --nav-btn-hover: #22d3ee;
    --date-separator-bg: #111826;
    --date-separator-text: #9ca3b5;
    --input-bg: #0f1724;
    --input-border: #1b2433;
    --input-text: #e8edf5;
    --highlight-bg: #22d3ee;
    --highlight-text: #0b1017;
    /* Category accents */
    --cat-one: #22d3ee;
    --cat-group: #a78bfa;
    --cat-meeting: #fbbf24;
    --cat-channel: #34d399;
}

/* === BASE STYLES === */
html, body {
    font-family: Arial, sans-serif;
    margin: 0;
    padding: 0;
    width: 100%;
    height: 100%;
    overflow: hidden;
    background: var(--bg);
    color: var(--text);
}

/* Allow scrolling inside stats/chat iframes */
body.stats-page,
body.chat-page {
    overflow-y: auto;
}

/* === SIDEBAR STYLES === */
.sidebar {
    width: 340px;
    color: var(--sidebar-text);
    position: fixed;
    top: 0;
    bottom: 0;
    left: 0;
    overflow-y: scroll;
    overflow-x: hidden;
    scrollbar-gutter: stable;
    background: var(--sidebar-bg);
    z-index: 100;
    box-shadow: 2px 0 8px rgba(0, 0, 0, 0.1);
}

.sidebar a {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px;
    color: var(--sidebar-text);
    text-decoration: none;
    border-bottom: 1px solid var(--sidebar-border);
    cursor: pointer;
    transition: all 0.2s ease;
}

.sidebar a:hover {
    background: var(--sidebar-hover);
}

.sidebar a.active {
    background: var(--sidebar-active);
    font-weight: bold;
}

/* === CONTENT AREA STYLES === */
.content {
    position: absolute;
    left: 340px;
    right: 0;
    top: 0;
    bottom: 0;
    padding: 20px;
    background-color: var(--content-bg);
    box-sizing: border-box;
    overflow-y: auto;
    overflow-x: hidden;
}

.content iframe {
    width: 100%;
    height: calc(100vh - 40px);
    border: none;
    background: var(--content-bg);
    display: block;
}

/* === SEARCH BAR STYLES === */
.search-container {
    padding: 15px;
    background: var(--sidebar-bg);
    border-bottom: 1px solid var(--sidebar-border);
    position: sticky;
    top: 0;
    z-index: 50;
}

.sidebar-meta {
    margin: 0 0 12px 0;
    font-size: 0.9em;
    color: var(--sidebar-muted);
}

#search-input {
    width: 100%;
    padding: 10px 34px 10px 12px;
    box-sizing: border-box;
    border: 1px solid var(--sidebar-border);
    border-radius: 6px;
    background: var(--input-bg);
    color: var(--input-text);
    font-size: 0.9em;
}

.search-input-wrap {
    position: relative;
}

.search-clear-btn {
    position: absolute;
    right: 8px;
    top: 50%;
    transform: translateY(-50%);
    width: 22px;
    height: 22px;
    border: none;
    border-radius: 50%;
    background: color-mix(in srgb, var(--sidebar-hover) 75%, #ffffff 25%);
    color: var(--sidebar-text);
    font-size: 14px;
    line-height: 1;
    cursor: pointer;
    display: none;
}

.search-clear-btn.show {
    display: inline-flex;
    align-items: center;
    justify-content: center;
}

.search-clear-btn:hover {
    background: var(--sidebar-hover);
}

.filter-buttons {
    display: flex;
    gap: 6px;
    margin-top: 10px;
    flex-wrap: nowrap;
}

.filter-btn {
    padding: 6px 12px;
    border: 1px solid var(--input-border);
    border-radius: 20px;
    background: var(--sidebar-hover);
    color: var(--sidebar-text);
    cursor: pointer;
    font-size: 0.85em;
    transition: all 0.2s ease;
}

.filter-btn:hover {
    background: var(--sidebar-hover);
}

.filter-btn.active {
    background: var(--nav-btn-bg);
    color: white;
    border-color: var(--nav-btn-bg);
}

#clear-filters-btn {
    margin-left: 0;
    background: #ef4444;
    color: white;
    border-color: #ef4444;
    border-radius: 6px;
}

.advanced-toggle {
    width: 100%;
    padding: 8px;
    margin-top: 10px;
    background: var(--sidebar-hover);
    border: 1px solid var(--sidebar-border);
    color: var(--sidebar-text);
    border-radius: 6px;
    cursor: pointer;
    font-size: 0.9em;
}

.advanced-panel {
    display: none;
    margin-top: 10px;
    padding: 10px;
    background: var(--sidebar-hover);
    border: 1px solid var(--sidebar-border);
    border-radius: 6px;
}

.advanced-panel.show {
    display: block;
}

.advanced-label {
    display: block;
    font-size: 0.8em;
    color: var(--sidebar-muted);
    margin-bottom: 4px;
    font-weight: 600;
}

.advanced-input {
    width: 100%;
    padding: 8px;
    margin-bottom: 8px;
    box-sizing: border-box;
    background: var(--input-bg);
    border: 1px solid var(--sidebar-border);
    border-radius: 4px;
    color: var(--input-text);
    font-size: 0.85em;
}

.date-row {
    display: flex;
    gap: 6px;
    margin-bottom: 8px;
}

.toolbar-divider {
    border-top: 1px solid var(--sidebar-border);
    margin-top: 10px;
    padding-top: 10px;
}

.switch {
    position: relative;
    display: inline-block;
    width: 44px;
    height: 24px;
    margin-right: 10px;
}

.switch input {
    opacity: 0;
    width: 0;
    height: 0;
}

.slider {
    position: absolute;
    cursor: pointer;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: #334155;
    transition: 0.2s;
    border-radius: 24px;
}

.slider:before {
    position: absolute;
    content: "";
    height: 18px;
    width: 18px;
    left: 3px;
    bottom: 3px;
    background-color: white;
    transition: 0.2s;
    border-radius: 50%;
}

.switch input:checked + .slider {
    background-color: var(--nav-btn-bg);
}

.switch input:checked + .slider:before {
    transform: translateX(20px);
}

.home-btn {
    display: block;
    padding: 10px;
    margin-top: 10px;
    background: var(--sidebar-active);
    border-radius: 6px;
    color: var(--sidebar-text);
    text-decoration: none;
    text-align: center;
    font-weight: 500;
    font-size: 0.9em;
    border: 1px solid var(--sidebar-border);
}

.home-btn:hover {
    background: var(--sidebar-hover);
}

/* === SIDEBAR SECTION STYLES === */
.sidebar-section-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 12px 15px;
    background: var(--sidebar-header-bg);
    color: var(--sidebar-text);
    cursor: pointer;
    border-bottom: 1px solid var(--sidebar-border);
    font-weight: bold;
    font-size: 0.95em;
    transition: all 0.2s ease;
    box-shadow: inset 0 0 0 1px color-mix(in srgb, var(--sidebar-border) 80%, transparent);
}

.sidebar-section-header:hover {
    background: var(--sidebar-hover);
}

.sidebar-section-header.top-header {
    background: var(--sidebar-header-bg);
    margin-top: 8px;
    margin-bottom: 6px;
    border-top: 1px solid var(--sidebar-border);
    border-bottom: 1px solid var(--sidebar-border);
    border-left: 4px solid transparent;
}

.sidebar-section-header.top-header[data-cat="oneonone"] { border-left-color: var(--cat-one); }
.sidebar-section-header.top-header[data-cat="group"] { border-left-color: var(--cat-group); }
.sidebar-section-header.top-header[data-cat="meeting"] { border-left-color: var(--cat-meeting); }
.sidebar-section-header.top-header[data-cat="channel"] { border-left-color: var(--cat-channel); }

.sidebar-section-header.top-header::before {
    display: inline-block;
    margin-right: 8px;
    opacity: 0.95;
}

.sidebar-section-header.top-header[data-cat="oneonone"]::before { content: "💬"; }
.sidebar-section-header.top-header[data-cat="group"]::before { content: "👪"; }
.sidebar-section-header.top-header[data-cat="meeting"]::before { content: "📅"; }
.sidebar-section-header.top-header[data-cat="channel"]::before { content: "📢"; }

.sidebar-section-header .toggle-icon {
    font-size: 1.2em;
    transition: transform 0.2s ease;
    min-width: 16px;
    text-align: center;
}

.sidebar-section-header.expanded .toggle-icon {
    transform: none;
}

.sidebar-section-header.expanded {
    background: color-mix(in srgb, var(--sidebar-active) 80%, var(--sidebar-bg) 20%);
    box-shadow: inset 0 0 0 1px color-mix(in srgb, var(--nav-btn-bg) 25%, var(--sidebar-border));
}

.sidebar-section-content {
    background: var(--sidebar-link-bg);
    max-height: 500px;
    overflow-y: auto;
    border-left: 2px solid transparent;
}

.sidebar-section-content[data-cat="oneonone"] {
    border-left-color: var(--cat-one);
}

.sidebar-section-content[data-cat="group"] {
    border-left-color: var(--cat-group);
}

.sidebar-section-content[data-cat="meeting"] {
    border-left-color: var(--cat-meeting);
}

.sidebar-section-content[data-cat="channel"] {
    border-left-color: var(--cat-channel);
}

.sidebar-section-content a {
    display: block;
    padding: 10px 16px;
    margin: 0;
    border-bottom: 1px solid var(--sidebar-border);
    transition: all 0.2s ease;
}

.sidebar-section-content a:hover {
    background: var(--sidebar-hover);
    padding-left: 20px;
}

.sidebar-section-content a.active {
    background: var(--sidebar-active);
    font-weight: bold;
}

/* Channel team rows should look nested and easier to read when expanded */
.sidebar-section-content[data-cat="channel"] > .sidebar-section-header {
    background: color-mix(in srgb, var(--sidebar-team-bg) 88%, var(--sidebar-bg) 12%);
    border-left: 3px solid color-mix(in srgb, var(--cat-channel) 60%, transparent);
    margin-left: 6px;
}

.sidebar-section-content[data-cat="channel"] > .sidebar-section-content {
    margin-left: 6px;
    border-left: 2px dashed color-mix(in srgb, var(--cat-channel) 55%, transparent);
}

/* === CHAT STYLES === */
.chat-section {
    display: block;
}

.chat-header {
    margin-top: 0;
    padding-bottom: 15px;
    border-bottom: 2px solid var(--input-border);
    margin-bottom: 20px;
}

.chat-members {
    color: var(--meta);
    font-size: 0.95em;
    margin-bottom: 20px;
    padding: 12px;
    background: color-mix(in srgb, var(--input-bg) 50%, var(--content-bg));
    border-radius: 4px;
    border-left: 3px solid var(--nav-btn-bg);
}

/* === MESSAGE STYLES === */
.message {
    padding: 12px 14px;
    border-radius: 16px;
    margin: 8px 0;
    max-width: 900px;
    clear: both;
    word-wrap: break-word;
    white-space: pre-wrap;
    word-break: break-word;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    transition: all 0.2s ease;
    display: flex;
    gap: 10px;
    align-items: flex-start;
}

.message:hover {
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
    transform: translateY(-1px);
}

.mine {
    background-color: var(--message-mine);
    float: right;
    text-align: left;
    margin-right: 24px;
}

.theirs {
    background-color: var(--message-theirs);
    float: left;
    text-align: left;
}

.message.sender-filtered {
    background-color: var(--highlight-bg) !important;
    border: 2px solid rgba(34, 211, 238, 0.4);
    box-shadow: 0 0 0 4px rgba(34, 211, 238, 0.1);
}

.meta {
    font-size: 0.8em;
    color: var(--meta);
    margin-bottom: 0;
    margin-left: 8px;
    display: flex;
    align-items: center;
    gap: 8px;
}

.copy-btn {
    background: none;
    border: none;
    cursor: pointer;
    font-size: 0.9em;
    opacity: 0;
    transition: opacity 0.2s ease;
    padding: 2px 4px;
    margin-left: auto;
}

.message:hover .copy-btn {
    opacity: 1;
}

.copy-btn:hover {
    transform: scale(1.2);
}

.text {
    margin: 0;
    padding: 0;
    margin-top: 2px;
}

.message-content {
    display: flex;
    flex-direction: column;
    flex: 1;
}

.message-header {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 8px;
}

.mine .avatar {
    display: none;
}

/* Clearfix for float layout */
.clearfix::after {
    content: '';
    display: block;
    clear: both;
}

.clearfix::after {
    content: "";
    clear: both;
    display: table;
}

/* === AVATAR STYLES === */
.avatar {
    width: 36px;
    height: 36px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.85em;
    font-weight: bold;
    color: white;
    flex-shrink: 0;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.15);
    user-select: none;
}

/* === IMAGE STYLES === */
img {
    max-width: 100%;
    height: auto;
    margin-top: 5px;
    cursor: pointer;
    display: block;
    object-fit: contain;
    border-radius: 4px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}

a.lightbox {
    text-decoration: none;
}

/* === LIGHTBOX STYLES === */
#lightbox-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0,0,0,0.8);
    display: none;
    justify-content: center;
    align-items: center;
    z-index: 9999;
    padding: 24px;
    box-sizing: border-box;
}

#lightbox-overlay.active {
    display: flex;
}

#lightbox-image {
    max-width: 96vw;
    max-height: 92vh;
    object-fit: contain;
    border-radius: 8px;
    box-shadow: 0 12px 30px rgba(0, 0, 0, 0.45);
}

#lightbox-close-btn {
    position: absolute;
    top: 12px;
    right: 16px;
    width: 40px;
    height: 40px;
    border: none;
    border-radius: 50%;
    background: rgba(0, 0, 0, 0.6);
    color: #ffffff;
    font-size: 28px;
    line-height: 1;
    cursor: pointer;
}

#lightbox-close-btn:hover {
    background: rgba(0, 0, 0, 0.8);
}

/* === DATE SEPARATORS === */
.date-separator {
    clear: both;
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 24px 0 16px;
    color: var(--date-separator-text);
    font-size: 0.9em;
}

.date-separator::before,
.date-separator::after {
    content: "";
    flex: 1;
    height: 1px;
    background: linear-gradient(to right, transparent, var(--input-border), transparent);
}

.date-separator::before {
    background: linear-gradient(to right, transparent, var(--input-border));
}

.date-separator::after {
    background: linear-gradient(to left, transparent, var(--input-border));
}

.date-separator span {
    background: var(--content-bg);
    padding: 6px 14px;
    border-radius: 16px;
    border: 1px solid var(--input-border);
    font-weight: 500;
    white-space: nowrap;
    flex-shrink: 0;
}

/* === NAVIGATION BUTTONS === */
.nav-button {
    position: fixed;
    right: 20px;
    background: var(--nav-btn-bg);
    color: white;
    border: none;
    padding: 12px 16px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 1.1em;
    z-index: 99;
    transition: all 0.2s ease;
}

.nav-button:hover {
    background: var(--nav-btn-hover);
    transform: scale(1.1);
}

#scroll-up-btn {
    bottom: 80px;
}

#scroll-down-btn {
    bottom: 20px;
}

/* === THEME TOGGLE === */
.theme-toggle {
    position: fixed;
    bottom: 20px;
    right: 80px;
    width: 40px;
    height: 40px;
    border-radius: 50%;
    background: var(--nav-btn-bg);
    border: none;
    cursor: pointer;
    color: white;
    font-size: 1.2em;
    z-index: 99;
    transition: all 0.2s ease;
}

.theme-toggle:hover {
    background: var(--nav-btn-hover);
    transform: scale(1.1);
}

/* === STATISTICS PANEL === */
.stats-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 15px;
    margin-bottom: 25px;
}

.stats-item {
    padding: 15px;
    background: var(--input-bg);
    border-radius: 8px;
    border-left: 4px solid;
    border: 1px solid color-mix(in srgb, var(--highlight-bg) 20%, var(--input-border));
}

.stats-label {
    font-size: 0.75em;
    color: var(--meta);
    margin-bottom: 5px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    font-weight: 600;
}

.stats-value {
    font-size: 2.2em;
    font-weight: bold;
    color: var(--text);
}

/* === HIDDEN ELEMENTS === */
.hidden {
    display: none !important;
}

.search-hidden {
    display: none !important;
}

mark.search-highlight {
    background: #ffea00;
    color: #111111;
    padding: 0 3px;
    border-radius: 3px;
    font-weight: 700;
}

/* === RESPONSIVE DESIGN === */
@media (max-width: 768px) {
    .sidebar {
        width: 280px;
    }
    
    .content {
        left: 280px;
        padding: 10px;
    }
    
    .stats-grid {
        grid-template-columns: 1fr;
    }
    
    .message {
        max-width: 100%;
    }
    
    .nav-button {
        right: 10px;
        padding: 10px 12px;
        font-size: 0.9em;
    }
    
    #scroll-up-btn {
        bottom: 70px;
    }
}

@media (max-width: 480px) {
    .sidebar {
        width: 230px;
    }
    
    .content {
        left: 230px;
        padding: 8px;
    }
    
    .message {
        padding: 8px 10px;
        font-size: 0.9em;
    }
}
"""

# =============================================================================
# INDEX.HTML TEMPLATE
# =============================================================================

INDEX_HTML = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Teams Chat Export</title>
<link rel="stylesheet" href="assets/style.css">
</head>
<body class="index-page">

<!-- SIDEBAR -->
<div class="sidebar">
    <!-- Search Bar -->
    <div class="search-container">
        <p class="sidebar-meta">Exported on <!--EXPORT_DATE--></p>
        <div class="search-input-wrap">
            <input type="text" id="search-input" placeholder="Search all messages...">
            <button id="search-clear-btn" class="search-clear-btn" type="button" aria-label="Clear search">&times;</button>
        </div>
        <div class="filter-buttons">
            <button id="filter-image-btn" class="filter-btn" title="Filter messages with images">🖼️ Images</button>
            <button id="filter-code-btn" class="filter-btn" title="Filter messages with code">💻 Code</button>
            <button id="filter-url-btn" class="filter-btn" title="Filter messages with URLs">🔗 URLs</button>
        </div>

        <button id="advanced-filters-toggle" class="advanced-toggle">⚙️ Advanced Filters <span id="advFilterToggleIcon">+</span></button>

        <div id="advanced-filters-panel" class="advanced-panel">
            <label class="advanced-label" for="sender-filter-input">👤 Sender</label>
            <input type="text" id="sender-filter-input" class="advanced-input" placeholder="Filter by sender">

            <label class="advanced-label">📅 Date Range</label>
            <div class="date-row">
                <input type="date" id="date-from-input" class="advanced-input" style="margin-bottom:0;">
                <input type="date" id="date-to-input" class="advanced-input" style="margin-bottom:0;">
            </div>
            <label style="display:flex;align-items:center;gap:6px;font-size:0.85em;cursor:pointer;">
                <input type="checkbox" id="date-filter-enabled" checked>
                Apply date filter
            </label>
        </div>

        <button id="clear-filters-btn" class="filter-btn" style="width:100%; margin-top:10px;">🗑️ Clear All Filters</button>

        <div class="toolbar-divider" style="display:flex;align-items:center;">
            <label class="switch" aria-label="Toggle dark mode">
                <input type="checkbox" id="dark-mode-toggle">
                <span class="slider"></span>
            </label>
            <span style="font-size:0.9em;">Dark mode</span>
        </div>

        <a id="home-menu-btn" class="home-btn" href="#">🏠 Home</a>
    </div>
    
    <!-- Navigation Sections -->
    <div id="sidebar-content">
        <!--SIDEBAR_SECTIONS_PLACEHOLDER-->
    </div>
</div>

<!-- CONTENT AREA -->
<div class="content" id="content-area">
    <!--CONTENT_PLACEHOLDER-->
</div>

<!-- LIGHTBOX OVERLAY -->
<div id="lightbox-overlay">
    <button id="lightbox-close-btn" type="button" aria-label="Close image preview">&times;</button>
    <img id="lightbox-image" src="" alt="">
</div>

<script>
window.__INLINE_SEARCH_INDEX__ = <!--INLINE_SEARCH_INDEX-->;
</script>

<script src="assets/script.js"></script>
</body>
</html>
"""

# =============================================================================
# STATS PAGE TEMPLATE (stats.html)
# =============================================================================

STATS_HTML = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Export Statistics</title>
<link rel="stylesheet" href="assets/style.css">
<style>
body {
    padding: 0;
}
#cover-page {
    display: block;
    text-align: center;
    padding: 60px 20px 40px 20px;
    color: var(--text);
}
#cover-page h1 {
    font-size: 2.5em;
    margin-bottom: 0.2em;
}
#cover-page p {
    font-size: 1.2em;
    max-width: 600px;
    margin: 0 auto 1.5em auto;
}
.stats-panel {
    background: var(--input-bg);
    border-radius: 12px;
    padding: 30px;
    margin: 2em auto;
    max-width: 900px;
    border: 1px solid var(--input-border);
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}
.stats-panel h2 {
    font-size: 1.3em;
    margin-top: 0;
    color: var(--text);
    margin-bottom: 1.5em;
}
.stats-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 15px;
    margin-bottom: 25px;
}
.stats-item {
    padding: 15px;
    background: var(--input-bg);
    border-radius: 8px;
    border-left: 4px solid;
    border: 1px solid var(--input-border);
}
.breakdown-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 10px;
}
.breakdown-item {
    padding: 12px;
    border-radius: 6px;
    border: 1px solid var(--input-border);
    text-align: center;
}
.breakdown-emoji {
    font-size: 1.5em;
    margin-bottom: 5px;
}
.breakdown-label {
    font-size: 0.7em;
    font-weight: 600;
    margin-bottom: 5px;
}
.breakdown-value {
    font-size: 1.8em;
    font-weight: bold;
    color: var(--text);
}
.instructions-list {
    text-align: left;
    display: inline-block;
    margin-bottom: 1.5em;
    color: var(--text);
}
.instructions-list li {
    margin-bottom: 8px;
}
.date-range-section {
    background: color-mix(in srgb, #ecfdf5 10%, var(--input-bg));
    border-radius: 8px;
    padding: 15px;
    margin-bottom: 20px;
    border: 1px solid color-mix(in srgb, #059669 20%, var(--input-border));
}
.date-range-label {
    font-size: 0.75em;
    color: #047857;
    margin-bottom: 8px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    font-weight: 600;
}
.participants-section {
    background: color-mix(in srgb, #f3e8ff 10%, var(--input-bg));
    border-radius: 8px;
    padding: 15px;
    border: 1px solid color-mix(in srgb, #8b5cf6 20%, var(--input-border));
}
.participants-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 8px;
}
.participant-item {
    padding: 12px;
    border-radius: 8px;
    border: 1px solid var(--input-border);
}
.participant-rank {
    font-size: 0.85em;
    color: var(--text);
    font-weight: 600;
    margin-bottom: 4px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.participant-count {
    font-size: 0.9em;
    color: var(--meta);
}
.footer {
    margin-top: 2em;
    font-size: 0.9em;
    color: var(--meta);
}
</style>
</head>
<body class="stats-page">

<div id="cover-page">
    <h1>Teams Chat Export</h1>
    <p>
        Welcome!<br>
        This file contains all your exported Microsoft Teams chats, meetings, and channel messages.<br>
        <span style="color:var(--meta);">Exported on <span id="cover-export-date"><!--EXPORT_DATE--></span></span>
    </p>
    
    <!-- Statistics Panel -->
    <div class="stats-panel">
        <h2>Export Statistics</h2>
        
        <!-- Main Stats Grid -->
        <div class="stats-grid">
            <div class="stats-item" style="border-left-color: #3b82f6;">
                <div class="stats-label">📊 Total Messages</div>
                <div class="stats-value"><!--TOTAL_MESSAGES--></div>
            </div>
            <div class="stats-item" style="border-left-color: #8b5cf6;">
                <div class="stats-label">💬 Total Chats</div>
                <div class="stats-value"><!--TOTAL_CHATS--></div>
            </div>
            <div class="stats-item" style="border-left-color: #10b981;">
                <div class="stats-label">🖼️ Total Images</div>
                <div class="stats-value"><!--TOTAL_IMAGES--></div>
            </div>
            <div class="stats-item" style="border-left-color: #f59e0b;">
                <div class="stats-label">💾 Storage Used</div>
                <div class="stats-value"><!--STORAGE_SIZE--></div>
            </div>
        </div>
        
        <!-- Date Range Section -->
        <div class="date-range-section" id="date-range-section">
            <div class="date-range-label">📅 Message Date Range</div>
            <div style="font-size: 1em; color: var(--text); margin: 0;">
                <strong><!--EARLIEST_DATE--></strong> → <strong><!--LATEST_DATE--></strong>
            </div>
        </div>
        
        <!-- Breakdown Section -->
        <div style="border-top: 1px solid var(--input-border); padding-top: 20px; margin-bottom: 20px;">
            <div style="font-size: 0.85em; color: var(--meta); margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600;">Chat Breakdown</div>
            <div class="breakdown-grid">
                <div class="breakdown-item" style="background: color-mix(in srgb, #dbeafe 10%, var(--input-bg));">
                    <div class="breakdown-emoji">💬</div>
                    <div class="breakdown-label" style="color: #1e40af;">1:1 CHATS</div>
                    <div class="breakdown-value"><!--ONE_ON_ONE_COUNT--></div>
                </div>
                <div class="breakdown-item" style="background: color-mix(in srgb, #ede9fe 10%, var(--input-bg));">
                    <div class="breakdown-emoji">👥</div>
                    <div class="breakdown-label" style="color: #5b21b6;">GROUP</div>
                    <div class="breakdown-value"><!--GROUP_COUNT--></div>
                </div>
                <div class="breakdown-item" style="background: color-mix(in srgb, #fef3c7 10%, var(--input-bg));">
                    <div class="breakdown-emoji">📅</div>
                    <div class="breakdown-label" style="color: #92400e;">MEETING</div>
                    <div class="breakdown-value"><!--MEETING_COUNT--></div>
                </div>
                <div class="breakdown-item" style="background: color-mix(in srgb, #dcfce7 10%, var(--input-bg));">
                    <div class="breakdown-emoji">📢</div>
                    <div class="breakdown-label" style="color: #15803d;">CHANNEL</div>
                    <div class="breakdown-value"><!--CHANNEL_COUNT--></div>
                </div>
            </div>
        </div>
        
        <!-- Top Participants Section -->
        <div class="participants-section" id="participants-section">
            <div style="font-size: 0.85em; color: var(--meta); margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600;">🏆 Top 10 Most Active Participants</div>
            <div class="participants-grid" id="participants-grid">
                <!--TOP_PARTICIPANTS_HTML-->
            </div>
        </div>
    </div>
    
    <ul class="instructions-list">
        <li>Browse all your chats and channels using the sidebar.</li>
        <li>Click on a chat or channel to view its messages.</li>
        <li>Use the search box to filter messages across all chats.</li>
        <li>Click images to view them in a lightbox.</li>
    </ul>
    <div class="footer">Powered by Teams Chat Export Script <!--SCRIPT_VERSION--></div>
</div>

<script src="assets/script.js"></script>
</body>
</html>
"""

# =============================================================================
# INDIVIDUAL CHAT PAGE TEMPLATE (chats/*.html)
# =============================================================================

CHAT_HTML = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title><!--CHAT_NAME--></title>
<link rel="stylesheet" href="../assets/style.css">
<style>
body {
    padding: 20px;
    background: var(--content-bg);
    color: var(--text);
}
.chat-header {
    margin-top: 0;
    padding-bottom: 15px;
    border-bottom: 2px solid var(--input-border);
    margin-bottom: 20px;
}
.chat-members {
    color: var(--meta);
    font-size: 0.95em;
    margin-bottom: 20px;
    padding: 12px;
    background: color-mix(in srgb, var(--input-bg) 50%, var(--content-bg));
    border-radius: 4px;
    border-left: 3px solid var(--nav-btn-bg);
}
#breadcrumb {
    margin-bottom: 20px;
    font-size: 0.95em;
    color: var(--meta);
}
.breadcrumb-item {
    margin: 0 5px;
}
.breadcrumb-item a {
    color: var(--nav-btn-bg);
    text-decoration: none;
}
.breadcrumb-item a:hover {
    text-decoration: underline;
}
</style>
</head>
<body class="chat-page">

<div id="breadcrumb"></div>

<div id="chat-content">
    <!--CHAT_CONTENT_PLACEHOLDER-->
</div>

<script src="../assets/script.js"></script>
</body>
</html>
"""

# =============================================================================
# SHARED JAVASCRIPT (assets/script.js)
# =============================================================================

SHARED_JAVASCRIPT = r"""
/**
 * Shared JavaScript for Teams Chat Export (Multi-File SPA)
 * Handles navigation, search, theme, and dynamic content loading
 */

let searchIndexData = null;
let currentPage = 'stats';
let chatDictionary = {};
let currentSearchState = {
    searchText: '',
    senderText: '',
    filterCode: false,
    filterUrl: false,
    filterImage: false,
    applyDateFilter: false,
    dateFrom: '',
    dateTo: ''
};

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    initializeLightbox();
    initializeThemeMessageSync();

    // Chat page runs inside iframe and only needs local message filtering.
    if (document.body.classList.contains('chat-page')) {
        initializeChatPageFiltering();
        initializeDarkMode();
        return;
    }

    initializeDarkMode();
    loadSearchIndex();
    initializeSearch();
    initializeNavigation();
    initializeToolbarControls();
    // Always start on stats for predictability
    loadStats();
});

/**
 * Load search index from JSON file
 */
function loadSearchIndex() {
    // Prefer inline data so search also works when opening index.html via file://
    if (window.__INLINE_SEARCH_INDEX__ && window.__INLINE_SEARCH_INDEX__.messages) {
        searchIndexData = window.__INLINE_SEARCH_INDEX__;
        if (searchIndexData.chatDictionary) {
            chatDictionary = searchIndexData.chatDictionary;
        }
        return;
    }

    // Skip loading when served from file:// to avoid CORS errors; search will be disabled
    if (location.protocol === 'file:') {
        console.warn('Search index not loaded (file:// protocol)');
        return;
    }
    // Only load on index page (not in iframe)
    if (window.location.pathname.includes('/chats/')) {
        return;
    }
    
    const assetPath = 'assets/search-index.json';
    fetch(assetPath)
        .then(response => response.ok ? response.json() : null)
        .then(data => {
            searchIndexData = data;
            if (data && data.chatDictionary) {
                chatDictionary = data.chatDictionary;
            }
        })
        .catch(err => console.log('Search index not available'));
}

/**
 * Initialize dark mode with localStorage
 */
function initializeDarkMode() {
    const isDark = localStorage.getItem('teams-export-dark-mode') === 'true';
    if (isDark) {
        document.body.classList.add('dark');
    }
    const darkToggle = document.getElementById('dark-mode-toggle');
    if (darkToggle) {
        darkToggle.checked = isDark;
    }
}

/**
 * Initialize theme toggle button
 */
function initializeToolbarControls() {
    const darkToggle = document.getElementById('dark-mode-toggle');
    const homeBtn = document.getElementById('home-menu-btn');
    const advToggle = document.getElementById('advanced-filters-toggle');
    const advPanel = document.getElementById('advanced-filters-panel');
    const advIcon = document.getElementById('advFilterToggleIcon');

    if (darkToggle) {
        darkToggle.addEventListener('change', function() {
            document.body.classList.toggle('dark', this.checked);
            localStorage.setItem('teams-export-dark-mode', this.checked ? 'true' : 'false');
            broadcastThemeToIframe(this.checked);
        });
    }

    if (homeBtn) {
        homeBtn.addEventListener('click', function(e) {
            e.preventDefault();
            showHome();
        });
    }

    if (advToggle && advPanel && advIcon) {
        advToggle.addEventListener('click', function() {
            const open = advPanel.classList.toggle('show');
            advIcon.textContent = open ? '-' : '+';
        });
    }
}

function initializeThemeMessageSync() {
    window.addEventListener('message', function(event) {
        if (!event.data || event.data.type !== 'teams-export-theme') return;
        document.body.classList.toggle('dark', Boolean(event.data.isDark));
    });
}

function broadcastThemeToIframe(isDark) {
    const iframe = document.getElementById('content-iframe');
    if (!iframe || !iframe.contentWindow) return;
    iframe.contentWindow.postMessage({ type: 'teams-export-theme', isDark: Boolean(isDark) }, '*');
}

/**
 * Initialize search functionality
 */
function initializeSearch() {
    const searchInput = document.getElementById('search-input');
    const searchClearBtn = document.getElementById('search-clear-btn');
    const senderInput = document.getElementById('sender-filter-input');
    const dateFromInput = document.getElementById('date-from-input');
    const dateToInput = document.getElementById('date-to-input');
    const dateFilterEnabled = document.getElementById('date-filter-enabled');
    const filterCodeBtn = document.getElementById('filter-code-btn');
    const filterUrlBtn = document.getElementById('filter-url-btn');
    const filterImageBtn = document.getElementById('filter-image-btn');
    const clearFiltersBtn = document.getElementById('clear-filters-btn');
    
    if (searchInput) {
        searchInput.addEventListener('input', function() {
            if (searchClearBtn) {
                searchClearBtn.classList.toggle('show', Boolean(searchInput.value));
            }
            performSearch();
        });
    }

    if (searchClearBtn) {
        searchClearBtn.addEventListener('click', function() {
            if (searchInput) {
                searchInput.value = '';
                searchInput.focus();
            }
            searchClearBtn.classList.remove('show');
            performSearch();
        });
    }
    if (senderInput) {
        senderInput.addEventListener('input', performSearch);
    }
    if (dateFromInput) {
        dateFromInput.addEventListener('change', performSearch);
    }
    if (dateToInput) {
        dateToInput.addEventListener('change', performSearch);
    }
    if (dateFilterEnabled) {
        dateFilterEnabled.addEventListener('change', performSearch);
    }
    
    if (filterCodeBtn) filterCodeBtn.addEventListener('click', function() {
        this.classList.toggle('active');
        performSearch();
    });
    
    if (filterUrlBtn) filterUrlBtn.addEventListener('click', function() {
        this.classList.toggle('active');
        performSearch();
    });
    
    if (filterImageBtn) filterImageBtn.addEventListener('click', function() {
        this.classList.toggle('active');
        performSearch();
    });
    
    if (clearFiltersBtn) clearFiltersBtn.addEventListener('click', function() {
        if (searchInput) searchInput.value = '';
        if (searchClearBtn) searchClearBtn.classList.remove('show');
        if (senderInput) senderInput.value = '';
        if (dateFromInput) dateFromInput.value = '';
        if (dateToInput) dateToInput.value = '';
        if (dateFilterEnabled) dateFilterEnabled.checked = true;
        if (filterCodeBtn) filterCodeBtn.classList.remove('active');
        if (filterUrlBtn) filterUrlBtn.classList.remove('active');
        if (filterImageBtn) filterImageBtn.classList.remove('active');
        performSearch();
    });
}

/**
 * Perform search across all indexed messages
 */
function performSearch() {
    const searchInput = document.getElementById('search-input');
    const senderInput = document.getElementById('sender-filter-input');
    const dateFromInput = document.getElementById('date-from-input');
    const dateToInput = document.getElementById('date-to-input');
    const dateFilterEnabled = document.getElementById('date-filter-enabled');
    const searchText = searchInput ? searchInput.value.trim().toLowerCase() : '';
    const senderText = senderInput ? senderInput.value.trim().toLowerCase() : '';
    const dateFrom = dateFromInput ? dateFromInput.value : '';
    const dateTo = dateToInput ? dateToInput.value : '';
    const applyDateFilter = dateFilterEnabled ? dateFilterEnabled.checked : false;
    const filterCode = document.getElementById('filter-code-btn')?.classList.contains('active') || false;
    const filterUrl = document.getElementById('filter-url-btn')?.classList.contains('active') || false;
    const filterImage = document.getElementById('filter-image-btn')?.classList.contains('active') || false;

    currentSearchState = {
        searchText,
        senderText,
        filterCode,
        filterUrl,
        filterImage,
        applyDateFilter,
        dateFrom,
        dateTo
    };
    
    if (!searchIndexData) return;

    const searchRegex = buildSearchRegex(searchText, false);
    const normalizedSenderText = normalizeSearchKey(senderText);
    
    // Get results from search index
    const results = searchIndexData.messages.filter(msg => {
        let matches = true;
        
        // Text search
        if (searchText) {
            const msgContent = String(msg.content || '');
            const msgSender = String(msg.sender || '');
            matches = Boolean(
                (searchRegex && searchRegex.test(msgContent)) ||
                (searchRegex && searchRegex.test(msgSender))
            );
        }

        const senderValue = normalizeSearchKey(msg.sender || '');
        if (normalizedSenderText && !senderValue.includes(normalizedSenderText)) {
            matches = false;
        }

        if (applyDateFilter && (dateFrom || dateTo)) {
            const msgDate = extractDatePart(msg.timestamp || '');
            if (dateFrom && msgDate < dateFrom) matches = false;
            if (dateTo && msgDate > dateTo) matches = false;
        }
        
        // Code filter
        if (filterCode && !msg.hasCode) matches = false;
        
        // URL filter
        if (filterUrl && !msg.hasUrl) matches = false;
        
        // Image filter
        if (filterImage && !msg.hasImage) matches = false;
        
        return matches;
    });

    applySidebarSearch(results, currentSearchState);
    broadcastSearchToIframe(currentSearchState);
    
    // Update results display (broadcast to current page)
    if (window.updateSearchResults) {
        window.updateSearchResults(results, searchText, filterCode, filterUrl, filterImage);
    }
}

function applySidebarSearch(results, state) {
    const hasActiveSearch = Boolean(
        state.searchText || state.senderText || state.filterCode || state.filterUrl ||
        state.filterImage || (state.applyDateFilter && (state.dateFrom || state.dateTo))
    );

    const normalizeKey = (value) => String(value || '').trim().toLowerCase();

    const linkCountsById = new Map();
    const linkCountsByName = new Map();
    results.forEach(item => {
        if (item.chatId) {
            linkCountsById.set(item.chatId, (linkCountsById.get(item.chatId) || 0) + 1);
        }
        if (item.chat) {
            const key = normalizeKey(item.chat);
            linkCountsByName.set(key, (linkCountsByName.get(key) || 0) + 1);
        }
    });

    const links = document.querySelectorAll('.sidebar a[data-chat-id]');
    links.forEach(link => {
        const chatName = link.getAttribute('data-chat-name') || '';
        const chatId = link.getAttribute('data-chat-id') || '';
        const baseLabel = link.dataset.baseLabel || chatName;
        const baseCount = parseInt(link.dataset.baseCount || '0', 10) || 0;
        const idCount = linkCountsById.get(chatId);
        const nameCount = linkCountsByName.get(normalizeKey(chatName)) || 0;
        const dynamicCount = hasActiveSearch ? ((idCount !== undefined) ? idCount : nameCount) : baseCount;

        link.dataset.currentCount = String(dynamicCount);
        link.textContent = `${baseLabel} (${dynamicCount})`;
        link.classList.toggle('search-hidden', hasActiveSearch && dynamicCount === 0);
    });

    document.querySelectorAll('.sidebar-section-content > .sidebar-section-header:not(.top-header)').forEach(teamHeader => {
        const teamLinksContainer = teamHeader.nextElementSibling;
        const headerContent = teamHeader.querySelector('.header-content');
        if (!teamLinksContainer || !headerContent) return;

        const visibleLinks = Array.from(teamLinksContainer.querySelectorAll('a[data-chat-id]'))
            .filter(a => !a.classList.contains('search-hidden'));
        const teamTotal = visibleLinks.reduce((sum, a) => sum + (parseInt(a.dataset.currentCount || '0', 10) || 0), 0);

        const baseLabel = headerContent.dataset.baseLabel || parseLabelAndCount(headerContent.textContent).label;
        const baseCount = parseInt(headerContent.dataset.baseCount || '0', 10) || teamTotal;
        headerContent.textContent = `${baseLabel} (${hasActiveSearch ? teamTotal : baseCount})`;

        const hideTeam = hasActiveSearch && teamTotal === 0;
        teamHeader.classList.toggle('search-hidden', hideTeam);
        teamLinksContainer.classList.toggle('search-hidden', hideTeam);
    });

    document.querySelectorAll('.sidebar-section-header.top-header').forEach(topHeader => {
        const section = topHeader.nextElementSibling;
        const headerContent = topHeader.querySelector('.header-content');
        if (!section || !headerContent) return;

        const visibleLinks = Array.from(section.querySelectorAll('a[data-chat-id]'))
            .filter(a => !a.classList.contains('search-hidden'));
        const sectionTotal = visibleLinks.reduce((sum, a) => sum + (parseInt(a.dataset.currentCount || '0', 10) || 0), 0);

        const baseLabel = headerContent.dataset.baseLabel || parseLabelAndCount(headerContent.textContent).label;
        const baseCount = parseInt(headerContent.dataset.baseCount || '0', 10) || sectionTotal;
        headerContent.textContent = `${baseLabel} (${hasActiveSearch ? sectionTotal : baseCount})`;

        const hideSection = hasActiveSearch && sectionTotal === 0;
        topHeader.classList.toggle('search-hidden', hideSection);
        section.classList.toggle('search-hidden', hideSection);
    });
}

function parseLabelAndCount(text) {
    const trimmed = (text || '').trim();
    const match = trimmed.match(/^(.*)\s\((\d+)\)$/);
    if (!match) {
        return { label: trimmed, count: 0 };
    }
    return { label: match[1], count: parseInt(match[2], 10) || 0 };
}

function initializeSidebarMetadata() {
    document.querySelectorAll('.sidebar a[data-chat-id]').forEach(link => {
        const parsed = parseLabelAndCount(link.textContent);
        link.dataset.baseLabel = parsed.label;
        link.dataset.baseCount = String(parsed.count);
        link.dataset.currentCount = String(parsed.count);
    });

    document.querySelectorAll('.sidebar .sidebar-section-header .header-content').forEach(content => {
        const parsed = parseLabelAndCount(content.textContent);
        content.dataset.baseLabel = parsed.label;
        content.dataset.baseCount = String(parsed.count);
    });
}

function broadcastSearchToIframe(state) {
    const iframe = document.getElementById('content-iframe');
    if (!iframe || !iframe.contentWindow) return;
    iframe.contentWindow.postMessage({ type: 'teams-export-search', state }, '*');
}

/**
 * Initialize sidebar navigation
 */
function initializeNavigation() {
    // Add event listeners to all chat items
    const sidebarItems = document.querySelectorAll('.sidebar a[data-chat-id]');
    
    sidebarItems.forEach(item => {
        item.addEventListener('click', function(e) {
            e.preventDefault();
            const chatId = this.getAttribute('data-chat-id');
            const chatName = this.getAttribute('data-chat-name');
            const chatType = this.getAttribute('data-chat-type');
            
            loadChat(chatId, chatName, chatType);
            
            // Update active state
            sidebarItems.forEach(i => i.classList.remove('active'));
            this.classList.add('active');
        });
    });
    initializeSidebarMetadata();
    // Section headers already have inline onclick toggles in the template
}

/**
 * Load a specific chat
 */
function loadChat(chatId, chatName = '', chatType = 'chat') {
    const contentArea = document.getElementById('content-area');
    if (!contentArea) return;
    
    // Create iframe to load chat file using chat ID
    const chatFile = `chats/${chatId}.html`;
    
    // Use iframe for isolation
    contentArea.innerHTML = `<iframe id="content-iframe" src="${chatFile}" style="width:100%; height: calc(100vh - 40px); border:none; background:var(--content-bg);"></iframe>`;

    const iframe = document.getElementById('content-iframe');
    if (iframe) {
        iframe.addEventListener('load', function() {
            broadcastThemeToIframe(document.body.classList.contains('dark'));
            broadcastSearchToIframe(currentSearchState);
        });
    }
    
    currentPage = 'chat';
    localStorage.setItem('last-page', 'chat');
}

/**
 * Load stats page
 */
function loadStats() {
    const contentArea = document.getElementById('content-area');
    if (!contentArea) return;
    
    contentArea.innerHTML = '<iframe id="content-iframe" src="stats.html" style="width:100%; height: calc(100vh - 40px); border:none; background:var(--content-bg);"></iframe>';

    const iframe = document.getElementById('content-iframe');
    if (iframe) {
        iframe.addEventListener('load', function() {
            broadcastThemeToIframe(document.body.classList.contains('dark'));
        });
    }
    
    currentPage = 'stats';
    localStorage.setItem('last-page', 'stats');
}

function initializeChatPageFiltering() {
    window.addEventListener('message', function(event) {
        if (!event.data || event.data.type !== 'teams-export-search') return;
        applyChatFilters(event.data.state || {});
    });
}

function escapeRegExp(str) {
    return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeSearchKey(value) {
    return String(value || '').trim().toLowerCase();
}

function extractDatePart(value) {
    const text = String(value || '');
    const match = text.match(/\d{4}-\d{2}-\d{2}/);
    return match ? match[0] : '';
}

function buildSearchRegex(term, isGlobal) {
    const t = String(term || '').trim();
    if (!t) return null;

    const escaped = escapeRegExp(t);
    const flags = isGlobal ? 'giu' : 'iu';

    // Single-token searches should match whole words only.
    if (!/\s/.test(t)) {
        return new RegExp(`\\b${escaped}\\b`, flags);
    }

    // Multi-word searches are treated as phrase searches.
    return new RegExp(escaped, flags);
}

function clearHighlights(root) {
    root.querySelectorAll('mark.search-highlight').forEach(mark => {
        const textNode = document.createTextNode(mark.textContent || '');
        mark.replaceWith(textNode);
    });
    root.normalize();
}

function highlightInElement(element, term) {
    if (!term) return;
    const regex = buildSearchRegex(term, true);
    if (!regex) return;

    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, {
        acceptNode(node) {
            if (!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
            const parent = node.parentElement;
            if (!parent) return NodeFilter.FILTER_REJECT;
            if (parent.closest('script, style, mark.search-highlight')) return NodeFilter.FILTER_REJECT;
            return NodeFilter.FILTER_ACCEPT;
        }
    });

    const textNodes = [];
    while (walker.nextNode()) {
        textNodes.push(walker.currentNode);
    }

    textNodes.forEach(node => {
        const text = node.nodeValue || '';
        regex.lastIndex = 0;
        if (!regex.test(text)) return;

        const frag = document.createDocumentFragment();
        let lastIndex = 0;
        text.replace(regex, (match, offset) => {
            if (offset > lastIndex) {
                frag.appendChild(document.createTextNode(text.slice(lastIndex, offset)));
            }
            const mark = document.createElement('mark');
            mark.className = 'search-highlight';
            mark.textContent = match;
            frag.appendChild(mark);
            lastIndex = offset + match.length;
            return match;
        });
        if (lastIndex < text.length) {
            frag.appendChild(document.createTextNode(text.slice(lastIndex)));
        }
        node.parentNode.replaceChild(frag, node);
    });
}

function applyChatFilters(state) {
    const searchText = (state.searchText || '').toLowerCase();
    const senderText = (state.senderText || '').toLowerCase();
    const filterCode = Boolean(state.filterCode);
    const filterUrl = Boolean(state.filterUrl);
    const filterImage = Boolean(state.filterImage);
    const applyDateFilter = Boolean(state.applyDateFilter);
    const dateFrom = state.dateFrom || '';
    const dateTo = state.dateTo || '';
    const searchRegex = buildSearchRegex(searchText, false);

    const messages = document.querySelectorAll('.message');
    messages.forEach(msg => {
        const wrapper = msg.closest('.clearfix') || msg;
        const textContainer = msg.querySelector('.text');
        if (textContainer) {
            clearHighlights(textContainer);
        }
        const text = (msg.textContent || '').toLowerCase();
        const sender = (msg.getAttribute('data-sender') || '').toLowerCase();
        const hasCode = Boolean(msg.querySelector('pre, code'));
        const hasImage = Boolean(msg.querySelector('img'));
        const hasUrl = Boolean(msg.querySelector('a[href^="http"], a[href^="https"], a[href^="www."]')) || /(https?:\/\/|www\.)/i.test(msg.textContent || '');

        let matches = true;
        if (searchText && !(searchRegex && (searchRegex.test(text) || searchRegex.test(sender)))) matches = false;
        if (senderText && !sender.includes(senderText)) matches = false;
        if (filterCode && !hasCode) matches = false;
        if (filterUrl && !hasUrl) matches = false;
        if (filterImage && !hasImage) matches = false;

        if (applyDateFilter && (dateFrom || dateTo)) {
            const meta = msg.querySelector('.meta');
            const metaText = meta ? (meta.textContent || '') : '';
            const dateMatch = metaText.match(/\d{4}-\d{2}-\d{2}/);
            const msgDate = dateMatch ? dateMatch[0] : '';
            if (dateFrom && msgDate && msgDate < dateFrom) matches = false;
            if (dateTo && msgDate && msgDate > dateTo) matches = false;
        }

        wrapper.style.display = matches ? '' : 'none';

        if (matches && searchText && textContainer) {
            highlightInElement(textContainer, searchText);
        }
    });

    // Show date separators only when following message block has visible items.
    const separators = document.querySelectorAll('.date-separator');
    separators.forEach(separator => {
        let hasVisibleMessages = false;
        let node = separator.nextElementSibling;
        while (node && !node.classList.contains('date-separator')) {
            if (node.classList.contains('clearfix') && node.style.display !== 'none') {
                hasVisibleMessages = true;
                break;
            }
            node = node.nextElementSibling;
        }
        separator.style.display = hasVisibleMessages ? 'flex' : 'none';
    });
}

/**
 * Load default page (stats or last visited)
 */
function loadDefaultPage() {
    const lastPage = localStorage.getItem('last-page');
    if (lastPage === 'stats' || !lastPage) {
        loadStats();
    }
}

/**
 * Initialize scroll buttons
 */
function initializeScrollButtons() {
    const upBtn = document.getElementById('scroll-up-btn');
    const downBtn = document.getElementById('scroll-down-btn');
    
    if (upBtn) {
        upBtn.addEventListener('click', function() {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });
    }
    
    if (downBtn) {
        downBtn.addEventListener('click', function() {
            window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
        });
    }
}

/**
 * Copy message to clipboard
 */
function copyMessage(messageId, messageText) {
    navigator.clipboard.writeText(messageText).then(() => {
        alert('Message copied to clipboard!');
    }).catch(err => {
        console.error('Failed to copy:', err);
    });
}

function ensureLightboxElements() {
    let overlay = document.getElementById('lightbox-overlay');
    if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'lightbox-overlay';
        overlay.innerHTML = '<button id="lightbox-close-btn" type="button" aria-label="Close image preview">&times;</button><img id="lightbox-image" src="" alt="">';
        document.body.appendChild(overlay);
    }

    if (!overlay.querySelector('#lightbox-close-btn')) {
        const closeBtn = document.createElement('button');
        closeBtn.id = 'lightbox-close-btn';
        closeBtn.type = 'button';
        closeBtn.setAttribute('aria-label', 'Close image preview');
        closeBtn.innerHTML = '&times;';
        overlay.insertBefore(closeBtn, overlay.firstChild);
    }

    return overlay;
}

function initializeLightbox() {
    const overlay = ensureLightboxElements();
    const closeBtn = overlay.querySelector('#lightbox-close-btn');

    document.addEventListener('click', function(e) {
        const link = e.target.closest('a.lightbox');
        const img = e.target.closest('img');

        if (link) {
            e.preventDefault();
            openLightbox(link.getAttribute('href') || '');
            return;
        }

        // Fallback: open standalone images inside message/chat content.
        if (img && img.closest('.message, .chat-section, #chat-content') && !overlay.contains(img)) {
            const src = img.getAttribute('src') || '';
            if (src) {
                e.preventDefault();
                openLightbox(src);
            }
        }
    });

    if (closeBtn) {
        closeBtn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            closeLightbox();
        });
    }

    overlay.addEventListener('click', function(e) {
        if (e.target === overlay) {
            closeLightbox();
        }
    });
}

/**
 * Open image in lightbox
 */
function openLightbox(imageSrc) {
    const overlay = document.getElementById('lightbox-overlay');
    const image = document.getElementById('lightbox-image');
    
    if (overlay && image) {
        image.src = imageSrc;
        overlay.classList.add('active');
    }
}

/**
 * Close lightbox
 */
function closeLightbox() {
    const overlay = document.getElementById('lightbox-overlay');
    if (overlay) {
        overlay.classList.remove('active');
    }
}

/**
 * Lightbox keyboard controls
 */
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        closeLightbox();
    }
});

/**
 * Debounce utility function
 */
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Legacy helper not used (sidebar uses data attributes). Keep as no-op for safety.
function showChat(encodedChatName) {
    // no-op
}

/**
 * Show home/stats page
 */
function showHome() {
    loadStats();
    
    // Reset sidebar active states
    document.querySelectorAll('[data-chat-name]').forEach(item => {
        item.classList.remove('active');
    });
}

/**
 * Get chat type from name
 */
function getChatType(chatName) {
    if (chatDictionary[chatName]) {
        return chatDictionary[chatName];
    }
    return 'chat';
}
"""

# =============================================================================
# SIDEBAR SECTIONS TEMPLATE
# =============================================================================

SIDEBAR_SECTION_HTML = '''
    <div class="sidebar-section">
        <div class="sidebar-section-header top-header oneonone-section-header" data-cat="oneonone" onclick="const open=this.nextElementSibling.style.display==='none'; this.nextElementSibling.style.display=open?'block':'none'; this.classList.toggle('expanded', open); const icon=this.querySelector('.toggle-icon'); icon&&(icon.textContent=open?'-':'+');">
            <div class="header-content">{one_title}</div>
            <span class="toggle-icon">+</span>
        </div>
        <div class="sidebar-section-content" data-cat="oneonone" style="display:none;">
            {one_items}
        </div>
        
        <div class="sidebar-section-header top-header group-section-header" data-cat="group" onclick="const open=this.nextElementSibling.style.display==='none'; this.nextElementSibling.style.display=open?'block':'none'; this.classList.toggle('expanded', open); const icon=this.querySelector('.toggle-icon'); icon&&(icon.textContent=open?'-':'+');">
            <div class="header-content">{group_title}</div>
            <span class="toggle-icon">+</span>
        </div>
        <div class="sidebar-section-content" data-cat="group" style="display:none;">
            {group_items}
        </div>
        
        <div class="sidebar-section-header top-header meeting-section-header" data-cat="meeting" onclick="const open=this.nextElementSibling.style.display==='none'; this.nextElementSibling.style.display=open?'block':'none'; this.classList.toggle('expanded', open); const icon=this.querySelector('.toggle-icon'); icon&&(icon.textContent=open?'-':'+');">
            <div class="header-content">{meeting_title}</div>
            <span class="toggle-icon">+</span>
        </div>
        <div class="sidebar-section-content" data-cat="meeting" style="display:none;">
            {meeting_items}
        </div>
        
        {channel_section}
    </div>
'''

CHANNEL_SECTION_HTML = '''
    <div class="sidebar-section-header top-header channel-section-header" data-cat="channel" onclick="const open=this.nextElementSibling.style.display==='none'; this.nextElementSibling.style.display=open?'block':'none'; this.classList.toggle('expanded', open); const icon=this.querySelector('.toggle-icon'); icon&&(icon.textContent=open?'-':'+');">
            <div class="header-content">{channel_title}</div>
            <span class="toggle-icon">+</span>
        </div>
        <div class="sidebar-section-content" data-cat="channel" style="display:none;">
            {channel_items}
        </div>
'''

CHAT_ITEM_HTML = '''<a href="#" data-chat-id="{chat_id}" data-chat-name="{name}" data-chat-type="{chat_type}" data-cat="{chat_type}">{name} ({count})</a>
'''

CHANNEL_TEAM_SECTION_HTML = '''
            <div class="sidebar-section-header" onclick="const open=this.nextElementSibling.style.display==='none'; this.nextElementSibling.style.display=open?'block':'none'; this.classList.toggle('expanded', open); const icon=this.querySelector('.toggle-icon'); icon&&(icon.textContent=open?'-':'+');">
                <div class="header-content">{team_name} ({team_count})</div>
                <span class="toggle-icon">+</span>
            </div>
            <div class="sidebar-section-content" style="display:none;">
                {channel_items}
            </div>
'''

SEARCH_INDEX_TEMPLATE = {
    'version': '2.0',
    'chatDictionary': {},
    'messages': []
}
