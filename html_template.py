#!/usr/bin/env python3
"""
HTML Template for Microsoft Teams Chat Export

This module contains the HTML template with embedded CSS and JavaScript for the Teams chat export functionality.
The template provides a responsive web interface with sidebar navigation, message display, search functionality,
and navigation buttons for scrolling through chat content.

Features:
- Responsive sidebar navigation for chats, groups, meetings, and channels
- Message display with proper formatting and styling
- Search functionality with message filtering
- Lightbox image viewer for embedded images
- Navigation buttons for smooth scrolling
- Collapsible sections for better organization

Structure:
- CSS: Styling for sidebar, content area, messages, and navigation buttons
- JavaScript: Interactive functionality for navigation, search, and UI controls
- HTML: Base structure with placeholders for dynamic content

Author: Alexander Wegner
Version: v0.1.5
Last Updated: 2026-01-09
"""

# =============================================================================
# NAVIGATION BUTTONS HTML
# =============================================================================

NAVIGATION_BUTTONS_HTML = '''
<!-- Navigation buttons -->
<button id="scroll-up-btn" class="nav-button" title="Scroll to top">↑</button>
<button id="scroll-down-btn" class="nav-button" title="Scroll to bottom">↓</button>
'''

# HTML template with embedded CSS and JavaScript
# This template provides the complete web interface for the Teams chat export
html_content = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Teams Chat Export</title>
<style>
/* === THEME VARIABLES === */
:root {
    --bg: #f4f6f8;
    --text: #0f172a;
    --sidebar-bg: #111827;
    --sidebar-border: #1f2937;
    --sidebar-hover: #1f2937;
    --sidebar-active: #243147;
    --sidebar-link-bg: #111827;
    --sidebar-header-bg: #0f172a;
    --sidebar-team-bg: #162133;
    --sidebar-channel-bg: #1a2538;
    --sidebar-text: #e5e7eb;
    --sidebar-muted: #9ca3af;
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
    --sidebar-bg: #0a0f19;
    --sidebar-border: #131c2b;
    --sidebar-hover: #121a28;
    --sidebar-active: #1b2433;
    --sidebar-link-bg: #0f172a;
    --sidebar-header-bg: #0d1421;
    --sidebar-team-bg: #111b2a;
    --sidebar-channel-bg: #142033;
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
body { font-family: Arial, sans-serif; margin: 0; display: flex; background: var(--bg); color: var(--text);} 

/* === SIDEBAR STYLES === */
.sidebar { width: 300px; color: var(--sidebar-text); position: fixed; top: 0; bottom: 0; overflow-y: auto; background: var(--sidebar-bg); }
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
.sidebar a:hover { background: var(--sidebar-hover); }
.sidebar a.active { background: var(--sidebar-active); font-weight: bold; }

/* === CONTENT AREA STYLES === */
.content {
    flex: 1;
    padding: 20px;
    background-color: var(--content-bg);
    margin-left: 300px;
    max-width: 1500px;   /* Limit content width for better readability */
    box-sizing: border-box;
    min-height: 100vh;
    overflow-x: auto; 
}

/* === CHAT SECTION STYLES === */
.chat-section { display: none; }

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
.mine { background-color: var(--message-mine); float: right; text-align: left; }
.theirs { background-color: var(--message-theirs); float: left; text-align: left; }
.message.sender-filtered {
    background-color: var(--highlight-bg) !important;
    border: 2px solid rgba(34, 211, 238, 0.4);
    box-shadow: 0 0 0 4px rgba(34, 211, 238, 0.1);
}
.meta { font-size: 0.8em; color: var(--meta); margin-bottom: 0; margin-left: 8px; display: flex; align-items: center; gap: 8px; }
.copy-btn { background: none; border: none; cursor: pointer; font-size: 0.9em; opacity: 0; transition: opacity 0.2s ease; padding: 2px 4px; margin-left: auto; }
.message:hover .copy-btn { opacity: 1; }
.copy-btn:hover { transform: scale(1.2); }
.text { margin: 0; padding: 0; margin-top: 2px; }
.message-content { display: flex; flex-direction: column; flex: 1; }
.message-header { display: flex; flex-direction: row; align-items: center; gap: 8px; }
/* Avatar always left, hide for sent messages (mine) */
.mine .avatar { display: none; }
.clearfix::after { content: ""; clear: both; display: table; }

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
    max-width: 100%;     /* Keep images inside their container */
    height: auto;
    margin-top: 5px;
    cursor: pointer;
    display: block;
    object-fit: contain;
    border-radius: 4px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}
a.lightbox { text-decoration: none; }

/* === LIGHTBOX STYLES === */
#lightbox-overlay {
    position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0,0,0,0.8); display: none; justify-content: center; align-items: center;
}
#lightbox-overlay img {
    max-width: 90%; max-height: 90%;
}

/* === SIDEBAR SECTION STYLES === */
.sidebar-section-header {
    cursor: pointer;
    background: var(--sidebar-header-bg);
    padding: 8px 10px;
    font-weight: bold;
    border-bottom: 1px solid var(--sidebar-border);
    user-select: none;
}
.top-header {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 12px 12px;
    margin-top: 8px;
    margin-bottom: 6px;
    border-top: 1px solid var(--sidebar-border);
    border-bottom: 1px solid var(--sidebar-border);
    border-radius: 4px;
}
.top-header[data-cat="oneonone"] { border-left: 4px solid var(--cat-one); }
.top-header[data-cat="group"] { border-left: 4px solid var(--cat-group); }
.top-header[data-cat="meeting"] { border-left: 4px solid var(--cat-meeting); }
.top-header[data-cat="channel"] { border-left: 4px solid var(--cat-channel); }

.top-header::before {
    display: inline-block;
    opacity: 0.95;
}
.top-header[data-cat="oneonone"]::before { content: "💬"; }
.top-header[data-cat="group"]::before { content: "👪"; }
.top-header[data-cat="meeting"]::before { content: "📅"; }
.top-header[data-cat="channel"]::before { content: "📢"; }
.sidebar-section-header:hover {
    background: var(--sidebar-hover);
}
.sidebar-section-content {
    padding-left: 0;
}

/* === CHAT TYPE SPECIFIC STYLES === */
/* One on One Chats */
#oneonone-section.sidebar-section-content {
    background: var(--sidebar-link-bg);
}
#oneonone-section a {
    background: var(--sidebar-link-bg);
    color: var(--sidebar-text);
}

/* Group Chats */
#group-section.sidebar-section-content {
    background: var(--sidebar-link-bg);
}
#group-section a {
    background: var(--sidebar-link-bg);
    color: var(--sidebar-text);
}

/* Meeting Chats */
#meeting-section.sidebar-section-content {
    background: var(--sidebar-link-bg);
}
#meeting-section a {
    background: var(--sidebar-link-bg);
    color: var(--sidebar-text);
}

/* Channel Chats main section */
#channel-section.sidebar-section-content {
    background: var(--sidebar-channel-bg);
}

/* Team headers under Channel Chats */
#channel-section .sidebar-section-header {
    background: var(--sidebar-team-bg);
    color: var(--sidebar-text);
    padding-left: 18px;
}

/* Channel links under teams */
#channel-section .sidebar-section-content a {
    background: var(--sidebar-channel-bg);
    color: var(--sidebar-text);
    padding-left: 32px;
}

/* === SIDEBAR LINKS === */
.sidebar a {
    display: block;
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
    background: #7c83a0 !important;
    color: #fff !important;
}

/* === HOVER AND ACTIVE STATES === */
.sidebar-section-header:hover {
    background: var(--sidebar-hover);
}

/* === BODY OVERFLOW CONTROL === */
/* Prevent horizontal scroll on body */
body {
    overflow-x: hidden;
}

/* === NAVIGATION BUTTONS === */
.nav-button {
    position: fixed;
    right: 20px;
    width: 50px;
    height: 50px;
    background-color: var(--nav-btn-bg);
    color: white;
    border: none;
    border-radius: 50%;
    cursor: pointer;
    box-shadow: 0 2px 10px rgba(0,0,0,0.3);
    z-index: 1000;
    font-size: 18px;
    font-weight: bold;
    display: none;
    transition: all 0.3s ease;
}

.nav-button:hover {
    background-color: var(--nav-btn-hover);
    transform: scale(1.1);
}

.nav-button:active {
    transform: scale(0.95);
}

/* Disabled state for navigation buttons */
.nav-button.disabled {
    opacity: 0.3;
    cursor: default;
}

.nav-button.disabled:hover {
    background-color: var(--nav-btn-bg);
    transform: none;
}

/* Button positioning */
#scroll-down-btn {
    bottom: 20px;
}

#scroll-up-btn {
    top: 20px;
}

/* === BREADCRUMB NAVIGATION === */
#breadcrumb {
    padding: 12px 20px;
    background: var(--input-bg);
    border-bottom: 1px solid var(--input-border);
    font-size: 0.95em;
    display: none;
}

#breadcrumb.show {
    display: block;
}

#breadcrumb a {
    color: var(--nav-btn-bg);
    text-decoration: none;
    cursor: pointer;
    transition: color 0.2s ease;
}

#breadcrumb a:hover {
    color: var(--nav-btn-hover);
    text-decoration: underline;
}

.breadcrumb-item {
    display: inline-flex;
    align-items: center;
    gap: 8px;
}

.breadcrumb-item::after {
    content: ">";
    margin-left: 8px;
    color: var(--meta);
    font-weight: bold;
}

.breadcrumb-item:last-child::after {
    content: "";
    margin-left: 0;
}

.breadcrumb-item:last-child {
    color: var(--text);
    font-weight: 600;
}

.breadcrumb-emoji {
    opacity: 0.8;
    font-size: 0.95em;
}

/* === INPUTS & TOGGLES === */
.sidebar input[type="text"] {
    background: var(--input-bg);
    border: 1px solid var(--input-border);
    color: var(--input-text);
    border-radius: 6px;
}
.sidebar input[type="text"]::placeholder {
    color: var(--sidebar-muted);
}
.sidebar label {
    color: var(--sidebar-muted);
}
.toggle-row {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 10px;
    padding: 6px 0 10px 0;
    color: var(--sidebar-muted);
    font-size: 0.95em;
}
.switch {
    position: relative;
    display: inline-block;
    width: 42px;
    height: 22px;
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
    background-color: var(--sidebar-border);
    transition: 0.2s ease;
    border-radius: 14px;
    box-shadow: inset 0 0 0 1px rgba(255,255,255,0.08);
}
.slider:before {
    position: absolute;
    content: "";
    height: 16px;
    width: 16px;
    left: 3px;
    bottom: 3px;
    background-color: var(--sidebar-text);
    transition: 0.2s ease;
    border-radius: 50%;
    box-shadow: 0 1px 2px rgba(0,0,0,0.25);
}
.switch input:checked + .slider {
    background-color: var(--nav-btn-hover);
}
.switch input:checked + .slider:before {
    transform: translateX(20px);
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

/* === SEARCH HIGHLIGHTING === */
mark.search-hit {
    background: #ffeb3b;
    color: #000;
    padding: 1px 4px;
    border-radius: 3px;
    font-weight: 600;
}

/* === RICH TEXT (CODE, QUOTES, LISTS) === */
.text pre {
    background: color-mix(in srgb, var(--input-bg) 85%, #000 15%);
    color: var(--text);
    border: 1px solid var(--input-border);
    border-radius: 8px;
    padding: 12px 14px;
    overflow: auto;
    line-height: 1.45;
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
    font-size: 0.9em;
    tab-size: 4;
    white-space: pre; /* keep exact formatting inside code blocks */
}

.text code {
    background: color-mix(in srgb, var(--input-bg) 90%, #000 10%);
    border: 1px solid color-mix(in srgb, var(--input-border) 80%, #000 20%);
    border-radius: 5px;
    padding: 0.15em 0.35em;
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
    font-size: 0.95em;
}

.text pre code {
    background: transparent;
    border: 0;
    padding: 0;
    font-size: 1em;
}

.text blockquote {
    margin: 10px 0;
    padding: 10px 14px;
    border-left: 4px solid color-mix(in srgb, var(--nav-btn-bg) 70%, #000 30%);
    background: color-mix(in srgb, var(--input-bg) 90%, #000 10%);
    border-radius: 6px;
    color: var(--text);
}

.text ul,
.text ol {
    margin: 8px 0 8px 1.25em;
    padding-left: 1.25em;
}

.text li {
    margin: 4px 0;
    line-height: 1.5;
}

.text li > ul,
.text li > ol {
    margin-top: 4px;
}

/* === DATE INPUT ICONS === */
.date-input {
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='18' height='18' viewBox='0 0 24 24' fill='none' stroke='%2322d3ee' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Crect x='3' y='4' width='18' height='18' rx='2' ry='2'/%3E%3Cline x1='16' y1='2' x2='16' y2='6'/%3E%3Cline x1='8' y1='2' x2='8' y2='6'/%3E%3Cline x1='3' y1='10' x2='21' y2='10'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: 6px center;
    background-size: 20px 20px;
    padding-left: 32px !important;
    cursor: pointer;
}
</style>

<script>
/* === AVATAR AND COLOR UTILITIES === */
/**
 * Generate a color based on sender name (consistent colors for same sender)
 * @param {string} senderName - The name of the sender
 * @returns {string} - A hex color code
 */
function getAvatarColor(senderName) {
    if (!senderName) return '#9ca3af';
    let hash = 0;
    for (let i = 0; i < senderName.length; i++) {
        hash = senderName.charCodeAt(i) + ((hash << 5) - hash);
    }
    const colors = [
        '#3b82f6', '#8b5cf6', '#ec4899', '#f59e0b', '#10b981',
        '#06b6d4', '#6366f1', '#d946ef', '#ea580c', '#16a34a'
    ];
    return colors[Math.abs(hash) % colors.length];
}

/**
 * Get initials from sender name
 * @param {string} senderName - The full name of the sender
 * @returns {string} - The initials (up to 2 characters)
 */
function getInitials(senderName) {
    if (!senderName) return '?';
    const names = senderName.trim().split(/\s+/);
    if (names.length >= 2) {
        return (names[0][0] + names[names.length - 1][0]).toUpperCase();
    }
    return senderName.substring(0, 2).toUpperCase();
}
</script>

<script>
/* === CHAT NAVIGATION FUNCTIONALITY === */
/**
 * Shows the specified chat section and updates active states
 * @param {string} id - The ID of the chat section to show
 */
function showChat(id) {
    // Hide cover page when showing a chat
    const coverPage = document.getElementById('cover-page');
    if (coverPage) {
        coverPage.style.display = 'none';
    }
    
    document.querySelectorAll('.sidebar a').forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('onclick')?.includes("'" + id + "'")) {
            link.classList.add('active');
        }
    });
    document.querySelectorAll('.chat-section').forEach(div => div.style.display = 'none');
    document.getElementById(id).style.display = 'block';
    
    // Update breadcrumb navigation
    updateBreadcrumb(id);
}

/* === LIGHTBOX FUNCTIONALITY === */
/**
 * Initialize lightbox functionality for images
 */
document.addEventListener("DOMContentLoaded", function() {
    const overlay = document.createElement("div");
    overlay.id = "lightbox-overlay";
    overlay.innerHTML = "<img>";
    document.body.appendChild(overlay);
    document.querySelectorAll("a.lightbox").forEach(link => {
        link.addEventListener("click", function(e) {
            e.preventDefault();
            const img = overlay.querySelector("img");
            img.src = this.href;
            overlay.style.display = "flex";
        });
    });
    overlay.addEventListener("click", function() {
        overlay.style.display = "none";
    });
});
</script>

<script>
/* === SEARCH FUNCTIONALITY === */
/**
 * Toggle advanced filters panel
 */
function toggleAdvancedFilters() {
    var panel = document.getElementById('advancedFiltersPanel');
    var icon = document.getElementById('advFilterToggleIcon');
    if (panel.style.display === 'none') {
        panel.style.display = 'block';
        icon.textContent = '▼';
    } else {
        panel.style.display = 'none';
        icon.textContent = '▶';
    }
}

/**
 * Toggle image filter and update active filters
 */
function toggleImageFilter() {
    var imageOnlyCheckbox = document.getElementById('imageOnlyCheckbox');
    var imageFilterBtn = document.getElementById('imageFilterBtn');
    
    imageOnlyCheckbox.checked = !imageOnlyCheckbox.checked;
    
    // Update button style
    if (imageOnlyCheckbox.checked) {
        imageFilterBtn.style.background = '#3b82f6';
        imageFilterBtn.style.color = 'white';
        imageFilterBtn.style.borderColor = '#3b82f6';
    } else {
        imageFilterBtn.style.background = 'var(--sidebar-hover)';
        imageFilterBtn.style.color = 'var(--sidebar-text)';
        imageFilterBtn.style.borderColor = 'var(--sidebar-border)';
    }
    
    updateActiveFilters();
    updateSearch();
}

/**
 * Toggle code block filter
 */
function toggleCodeFilter() {
    var codeOnlyCheckbox = document.getElementById('codeOnlyCheckbox');
    var codeFilterBtn = document.getElementById('codeFilterBtn');
    
    codeOnlyCheckbox.checked = !codeOnlyCheckbox.checked;
    
    // Update button style
    if (codeOnlyCheckbox.checked) {
        codeFilterBtn.style.background = '#3b82f6';
        codeFilterBtn.style.color = 'white';
        codeFilterBtn.style.borderColor = '#3b82f6';
    } else {
        codeFilterBtn.style.background = 'var(--sidebar-hover)';
        codeFilterBtn.style.color = 'var(--sidebar-text)';
        codeFilterBtn.style.borderColor = 'var(--sidebar-border)';
    }
    
    updateActiveFilters();
    updateSearch();
}

/**
 * Update active filters display
 */
function updateActiveFilters() {
    var activeFiltersContainer = document.getElementById('activeFiltersContainer');
    var activeFiltersList = document.getElementById('activeFiltersList');
    var clearAllBtn = document.getElementById('clearAllBtn');
    var filters = [];
    
    // Check all active filters
    var searchInput = document.getElementById('searchInput');
    var senderInput = document.getElementById('senderFilterInput');
    var imageOnlyCheckbox = document.getElementById('imageOnlyCheckbox');
    var codeOnlyCheckbox = document.getElementById('codeOnlyCheckbox');
    var dateFromInput = document.getElementById('dateFromInput');
    var dateToInput = document.getElementById('dateToInput');
    
    if (searchInput && searchInput.value) {
        filters.push({ text: '🔍 ' + searchInput.value, type: 'search' });
    }
    if (senderInput && senderInput.value) {
        filters.push({ text: '👤 ' + senderInput.value, type: 'sender' });
    }
    if (imageOnlyCheckbox && imageOnlyCheckbox.checked) {
        filters.push({ text: '🖼️ Images only', type: 'images' });
    }
    if (codeOnlyCheckbox && codeOnlyCheckbox.checked) {
        filters.push({ text: '💻 Code only', type: 'code' });
    }
    if (dateFromInput && dateFromInput.value) {
        filters.push({ text: '📅 From: ' + dateFromInput.value, type: 'dateFrom' });
    }
    if (dateToInput && dateToInput.value) {
        filters.push({ text: '📅 To: ' + dateToInput.value, type: 'dateTo' });
    }
    
    if (filters.length > 0) {
        activeFiltersContainer.style.display = 'block';
        clearAllBtn.style.display = 'block';
        activeFiltersList.innerHTML = filters.map(f => 
            '<span style="background: #3b82f6; color: white; padding: 4px 10px; border-radius: 12px; font-size: 0.8em; display: inline-flex; align-items: center; gap: 4px;">' +
            f.text +
            '</span>'
        ).join('');
    } else {
        activeFiltersContainer.style.display = 'none';
        clearAllBtn.style.display = 'none';
        activeFiltersList.innerHTML = '';
    }
}

/**
 * Clear all filters
 */
function clearAllFilters() {
    document.getElementById('searchInput').value = '';
    document.getElementById('senderFilterInput').value = '';
    document.getElementById('dateFromInput').value = '';
    document.getElementById('dateToInput').value = '';
    document.getElementById('imageOnlyCheckbox').checked = false;
    var codeOnlyCheckbox = document.getElementById('codeOnlyCheckbox');
    if (codeOnlyCheckbox) codeOnlyCheckbox.checked = false;
    
    // Reset button styles for quick filter tags
    var imageFilterBtn = document.getElementById('imageFilterBtn');
    var codeFilterBtn = document.getElementById('codeFilterBtn');
    if (imageFilterBtn) {
        imageFilterBtn.style.background = 'var(--sidebar-hover)';
        imageFilterBtn.style.color = 'var(--sidebar-text)';
        imageFilterBtn.style.borderColor = 'var(--sidebar-border)';
    }
    if (codeFilterBtn) {
        codeFilterBtn.style.background = 'var(--sidebar-hover)';
        codeFilterBtn.style.color = 'var(--sidebar-text)';
        codeFilterBtn.style.borderColor = 'var(--sidebar-border)';
    }
    
    updateActiveFilters();
    updateSearch();
}

/**
 * Advanced search functionality with debouncing and image filtering
 * Provides real-time search across all chat messages with result counts
 */
document.addEventListener("DOMContentLoaded", function() {
    const input = document.getElementById("searchInput");
    const imageOnlyCheckbox = document.getElementById("imageOnlyCheckbox");
    const codeOnlyCheckbox = document.getElementById("codeOnlyCheckbox");
    const senderInput = document.getElementById("senderFilterInput");
    const sidebarLinks = document.querySelectorAll(".sidebar a");
    const chatSections = document.querySelectorAll(".chat-section");

    let debounceTimer = null;
    
    /**
     * Debounce function to limit search frequency
     * @param {Function} fn - Function to debounce
     * @param {number} delay - Delay in milliseconds
     * @returns {Function} Debounced function
     */
    function debounce(fn, delay) {
        return function(...args) {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(() => fn.apply(this, args), delay);
        };
    }

    /**
     * Check if a message contains an image
     * @param {HTMLElement} msg - Message element to check
     * @returns {boolean} True if message contains an image
     */
    function messageHasImage(msg) {
        // Checks if the message contains an <img ...> tag
        return msg.querySelector("img") !== null;
    }

    /**
     * Remove existing highlights from a container
     * @param {HTMLElement} container
     */
    function clearHighlights(container) {
        if (!container) return;
        container.querySelectorAll('mark.search-hit').forEach(mark => {
            const textNode = document.createTextNode(mark.textContent);
            mark.parentNode.replaceChild(textNode, mark);
        });
        container.normalize();
    }

    /**
     * Apply highlight markup to matching text nodes
     * @param {HTMLElement} container
     * @param {string} query
     */
    function applyHighlights(container, query) {
        if (!container || !query) return;
        const lowered = query.toLowerCase();
        const walker = document.createTreeWalker(
            container,
            NodeFilter.SHOW_TEXT,
            {
                acceptNode(node) {
                    if (!node.parentNode) return NodeFilter.FILTER_REJECT;
                    const parentName = node.parentNode.nodeName;
                    if (parentName === 'SCRIPT' || parentName === 'STYLE') return NodeFilter.FILTER_REJECT;
                    if (!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
                    return NodeFilter.FILTER_ACCEPT;
                }
            }
        );

        const nodes = [];
        while (walker.nextNode()) {
            nodes.push(walker.currentNode);
        }

        nodes.forEach(node => {
            const text = node.nodeValue;
            let idx = text.toLowerCase().indexOf(lowered);
            if (idx === -1) return;

            const frag = document.createDocumentFragment();
            let lastIndex = 0;

            while (idx !== -1) {
                if (idx > lastIndex) {
                    frag.appendChild(document.createTextNode(text.slice(lastIndex, idx)));
                }
                const mark = document.createElement('mark');
                mark.className = 'search-hit';
                mark.textContent = text.slice(idx, idx + query.length);
                frag.appendChild(mark);
                lastIndex = idx + query.length;
                idx = text.toLowerCase().indexOf(lowered, lastIndex);
            }

            if (lastIndex < text.length) {
                frag.appendChild(document.createTextNode(text.slice(lastIndex)));
            }

            node.parentNode.replaceChild(frag, node);
        });
    }

    /**
     * Determine if a message should be shown based on query and image filter
     * @param {HTMLElement} msg
     * @param {string} query
     * @param {boolean} imageOnly
     * @param {string} senderQuery
     * @param {Date|null} dateFrom
     * @param {Date|null} dateTo
     * @returns {boolean}
     */
    function shouldShowMessage(msg, query, imageOnly, codeOnly, senderQuery, dateFrom, dateTo) {
        if (!msg.dataset.lowerText) {
            msg.dataset.lowerText = msg.innerText.toLowerCase();
        }
        if (!msg.dataset.senderLower) {
            const s = (msg.getAttribute('data-sender') || '').toLowerCase();
            msg.dataset.senderLower = s;
        }
        const hasImg = messageHasImage(msg);
        const hasCode = msg.querySelector('code, pre') !== null;
        const matchText = msg.dataset.lowerText.includes(query);
        const senderOk = !senderQuery || msg.dataset.senderLower.includes(senderQuery);
        const textOk = (query === "") || matchText;
        const imageOk = !imageOnly || hasImg;
        const codeOk = !codeOnly || hasCode;
        
        // Date range filter
        let dateOk = true;
        if (dateFrom || dateTo) {
            const metaEl = msg.querySelector('.meta');
            if (metaEl) {
                // Extract timestamp from meta text (format: "Sender • YYYY-MM-DD HH:MM:SS")
                const metaText = metaEl.textContent;
                const dateMatch = metaText.match(/(\d{4}-\d{2}-\d{2})/);
                if (dateMatch) {
                    const msgDate = new Date(dateMatch[1]);
                    if (dateFrom && msgDate < dateFrom) dateOk = false;
                    if (dateTo && msgDate > dateTo) dateOk = false;
                }
            }
        }
        
        return textOk && imageOk && codeOk && senderOk && dateOk;
    }

    /**
     * Update search results based on current query and filters
     */
    function updateSearch() {
        const query = input.value.toLowerCase();
        const senderQuery = (senderInput ? senderInput.value.toLowerCase() : "");
        const imageOnly = imageOnlyCheckbox.checked;
        const codeOnly = codeOnlyCheckbox.checked;
        
        // Get date range values
        const dateFromInput = document.getElementById('dateFromInput');
        const dateToInput = document.getElementById('dateToInput');
        const dateRangeOnlyCheckbox = document.getElementById('dateRangeOnlyCheckbox');
        const rawDateFrom = dateFromInput && dateFromInput.value ? new Date(dateFromInput.value) : null;
        const rawDateTo = dateToInput && dateToInput.value ? new Date(dateToInput.value + 'T23:59:59') : null;
        const useDateFilter = (dateRangeOnlyCheckbox ? dateRangeOnlyCheckbox.checked : false) && (rawDateFrom || rawDateTo);
        const dateFrom = useDateFilter ? rawDateFrom : null;
        const dateTo = useDateFilter ? rawDateTo : null;
        
        const isFiltering = query !== "" || imageOnly || codeOnly || senderQuery !== "" || useDateFilter;

        /**
         * Process a specific chat section and update match counts
         * @param {string} sectionId - ID of the section to process
         * @param {string} sectionName - Display name of the section
         */
        function processSection(sectionId, sectionName) {
            const section = document.getElementById(sectionId);
            if (!section) return;
            
            const links = section.querySelectorAll('a');
            let totalMatches = 0;

            links.forEach(link => {
                // Store original text on first run
                if (!link.dataset.originalText) {
                    link.dataset.originalText = link.innerText;
                }
                
                try {
                    const chatId = link.getAttribute("onclick").match(/'(.*?)'/)[1];
                    const sectionDiv = document.getElementById(chatId);
                    if (!sectionDiv) return;
                    
                    let matchCount = 0;

                    sectionDiv.querySelectorAll(".message").forEach(msg => {
                        const show = shouldShowMessage(msg, query, imageOnly, codeOnly, senderQuery, dateFrom, dateTo);
                        // Hide/show the message and its wrapper to fully remove from flow
                        msg.style.display = show ? "block" : "none";
                        const wrapper = msg.closest('.clearfix');
                        if (wrapper) wrapper.style.display = show ? "" : "none";
                        
                        // Add sender-filtered class for highlighting when sender filter is active
                        if (show && senderQuery && msg.dataset.senderLower.includes(senderQuery)) {
                            msg.classList.add('sender-filtered');
                        } else {
                            msg.classList.remove('sender-filtered');
                        }
                        
                        const target = msg.querySelector('.text') || msg;
                        clearHighlights(target);
                        if (show && query) {
                            applyHighlights(target, query);
                        }
                        const meta = msg.querySelector('.meta');
                        if (meta) {
                            clearHighlights(meta);
                            if (show && senderQuery) {
                                applyHighlights(meta, senderQuery);
                            }
                        }
                        if (show) matchCount++;
                    });

                    // Hide date separators that have no visible messages beneath them
                    sectionDiv.querySelectorAll('.date-separator').forEach(sep => {
                        let hasVisibleAfter = false;
                        let node = sep.nextElementSibling;
                        while (node && !hasVisibleAfter) {
                            if (node.classList && node.classList.contains('date-separator')) break;
                            if (node.classList && node.classList.contains('clearfix')) {
                                if (node.style.display !== 'none') hasVisibleAfter = true;
                            }
                            node = node.nextElementSibling;
                        }
                        sep.style.display = hasVisibleAfter ? "" : "none";
                    });

                    if (isFiltering) {
                        // When filtering is active, show filtered count or hide
                        if (matchCount > 0) {
                            link.style.display = "";
                            link.innerText = link.getAttribute("data-chat-name") + ` (${matchCount})`;
                        } else {
                            link.style.display = "none";
                        }
                    } else {
                        // When no filtering, restore original text with permanent counts
                        link.style.display = "";
                        link.innerText = link.dataset.originalText;
                    }

                    totalMatches += matchCount;
                } catch (e) {
                    console.error("Error processing link:", e);
                }
            });

            // Store original header text
            const header = section.previousElementSibling;
            if (!header) return;
            
            if (!header.dataset.originalText) {
                header.dataset.originalText = header.innerText;
            }
            
            let arrow = header.innerText.trim().charAt(0);
            if (arrow !== '▶' && arrow !== '▼') arrow = '▶';
            
            if (isFiltering) {
                // Keep header visible and show filtered count (even when zero)
                header.style.display = "";
                header.innerText = `${arrow} ${sectionName} (${totalMatches})`;
            } else {
                // When no filtering, restore original text and show section
                header.style.display = "";
                const origText = header.dataset.originalText;
                header.innerText = origText.replace(/^[▶▼]\s*/, `${arrow} `);
            }
        }

        processSection('oneonone-section', 'One on One Chats');
        processSection('group-section', 'Group Chats');
        processSection('meeting-section', 'Meeting Chats');

        // === Channel Chats ===
        const channelSection = document.getElementById('channel-section');
        const teamHeaders = channelSection ? channelSection.querySelectorAll('.sidebar-section-header') : [];
        let totalChannelMatches = 0;

        teamHeaders.forEach(header => {
            // Store original header text
            if (!header.dataset.originalText) {
                header.dataset.originalText = header.innerText;
            }
            
            const teamDiv = header.nextElementSibling;
            const channelLinks = teamDiv ? teamDiv.querySelectorAll('a') : [];
            let teamMatchCount = 0;

            channelLinks.forEach(link => {
                // Store original link text
                if (!link.dataset.originalText) {
                    link.dataset.originalText = link.innerText;
                }
                
                const chatId = link.getAttribute("onclick").match(/'(.*?)'/)[1];
                const sectionDiv = document.getElementById(chatId);
                let matchCount = 0;

                if (sectionDiv) {
                    sectionDiv.querySelectorAll(".message").forEach(msg => {
                        const show = shouldShowMessage(msg, query, imageOnly, codeOnly, senderQuery, dateFrom, dateTo);
                        // Hide/show the message and its wrapper to fully remove from flow
                        msg.style.display = show ? "block" : "none";
                        const wrapper = msg.closest('.clearfix');
                        if (wrapper) wrapper.style.display = show ? "" : "none";
                        const target = msg.querySelector('.text') || msg;
                        clearHighlights(target);
                        if (show && query) {
                            applyHighlights(target, query);
                        }
                        const meta = msg.querySelector('.meta');
                        if (meta) {
                            clearHighlights(meta);
                            if (show && senderQuery) {
                                applyHighlights(meta, senderQuery);
                            }
                        }
                        if (show) matchCount++;
                    });
                    // Hide date separators that have no visible messages beneath them
                    sectionDiv.querySelectorAll('.date-separator').forEach(sep => {
                        let hasVisibleAfter = false;
                        let node = sep.nextElementSibling;
                        while (node && !hasVisibleAfter) {
                            if (node.classList && node.classList.contains('date-separator')) break;
                            if (node.classList && node.classList.contains('clearfix')) {
                                if (node.style.display !== 'none') hasVisibleAfter = true;
                            }
                            node = node.nextElementSibling;
                        }
                        sep.style.display = hasVisibleAfter ? "" : "none";
                    });
                }

                if (isFiltering) {
                    // When filtering is active, show filtered count
                    if (matchCount > 0) {
                        link.style.display = "";
                        link.innerText = link.getAttribute("data-chat-name") + ` (${matchCount})`;
                    } else {
                        link.style.display = "none";
                    }
                } else {
                    // When no filtering, restore original text with permanent counts
                    link.style.display = "";
                    link.innerText = link.dataset.originalText;
                }

                teamMatchCount += matchCount;
            });

            // Ensure we know the current expanded state (defaults to current display)
            if (!header.dataset.expanded) {
                header.dataset.expanded = teamDiv && teamDiv.style.display !== "none" ? "true" : "false";
            }

            // Update team header and visibility
            let arrow = header.innerText.trim().charAt(0);
            if (arrow !== '▶' && arrow !== '▼') arrow = '▶';
            const origText = header.dataset.originalText.replace(/^[▶▼]\s*/, '').replace(/\s*\(\d+\)$/, '');
            
            if (isFiltering) {
                // When filtering, show filtered count and hide team if no matches
                if (teamMatchCount > 0) {
                    header.style.display = "";
                    if (teamDiv) {
                        const expanded = header.dataset.expanded === "true";
                        teamDiv.style.display = expanded ? "" : "none";
                    }
                    header.innerText = `${arrow} ${origText} (${teamMatchCount})`;
                } else {
                    header.style.display = "none";
                    if (teamDiv) teamDiv.style.display = "none";
                }
            } else {
                // When no filtering, restore to the user's expanded/collapsed state
                header.style.display = "";
                if (teamDiv) {
                    const expanded = header.dataset.expanded === "true";
                    teamDiv.style.display = expanded ? "" : "none";
                }
                const desiredArrow = header.dataset.expanded === "true" ? '▼' : '▶';
                header.innerText = header.dataset.originalText.replace(/^[▶▼]\s*/, `${desiredArrow} `);
            }

            totalChannelMatches += teamMatchCount;
        });

        const channelHeader = channelSection ? channelSection.previousElementSibling : null;
        if (channelHeader) {
            // Store original channel header text
            if (!channelHeader.dataset.originalText) {
                channelHeader.dataset.originalText = channelHeader.innerText;
            }
            if (!channelHeader.dataset.expanded) {
                channelHeader.dataset.expanded = channelSection && channelSection.style.display !== "none" ? "true" : "false";
            }
            
            let arrow = channelHeader.innerText.trim().charAt(0);
            if (arrow !== '▶' && arrow !== '▼') arrow = '▶';
            const name = "Channel Chats";
            
            if (isFiltering) {
                // Keep header visible and show filtered count (even when zero)
                channelHeader.style.display = "";
                channelHeader.innerText = `${arrow} ${name} (${totalChannelMatches})`;
                if (channelSection) {
                    const expanded = channelHeader.dataset.expanded === "true";
                    // Hide section body when zero matches to avoid empty space
                    channelSection.style.display = totalChannelMatches > 0 && expanded ? "" : "none";
                }
            } else {
                // When no filtering, restore to user's expanded/collapsed state
                channelHeader.style.display = "";
                if (channelSection) {
                    const expanded = channelHeader.dataset.expanded === "true";
                    channelSection.style.display = expanded ? "" : "none";
                }
                const desiredArrow = channelHeader.dataset.expanded === "true" ? '▼' : '▶';
                const origText = channelHeader.dataset.originalText;
                channelHeader.innerText = origText.replace(/^[▶▼]\s*/, `${desiredArrow} `);
                // Make updateSearch available globally so toggle functions can call it
                window.updateSearch = updateSearch;
            }
        }
    }

    /**
     * Enhanced showChat function with search integration
     * @param {string} id - Chat ID to display
     */
    window.showChat = function(id) {
        sidebarLinks.forEach(link => {
            link.classList.remove('active');
            if (link.getAttribute('onclick')?.includes("'" + id + "'")) {
                link.classList.add('active');
            }
        });
        chatSections.forEach(div => div.style.display = "none");
        const section = document.getElementById(id);
        section.style.display = "block";
        // Hide cover page when showing a chat
        document.getElementById("cover-page").style.display = "none";
        
        // Apply current search filters to the displayed chat
        const query = input.value.toLowerCase();
        const senderQuery = (senderInput ? senderInput.value.toLowerCase() : "");
        const imageOnly = imageOnlyCheckbox.checked;
        const codeOnly = codeOnlyCheckbox.checked;
        
        // Get date range values
        const dateFromInput = document.getElementById('dateFromInput');
        const dateToInput = document.getElementById('dateToInput');
        const dateRangeOnlyCheckbox = document.getElementById('dateRangeOnlyCheckbox');
        const rawDateFrom = dateFromInput && dateFromInput.value ? new Date(dateFromInput.value) : null;
        const rawDateTo = dateToInput && dateToInput.value ? new Date(dateToInput.value + 'T23:59:59') : null;
        const useDateFilter = (dateRangeOnlyCheckbox ? dateRangeOnlyCheckbox.checked : false) && (rawDateFrom || rawDateTo);
        const dateFrom = useDateFilter ? rawDateFrom : null;
        const dateTo = useDateFilter ? rawDateTo : null;
        
        section.querySelectorAll(".message").forEach(msg => {
            const show = shouldShowMessage(msg, query, imageOnly, codeOnly, senderQuery, dateFrom, dateTo);
            msg.style.display = show ? "block" : "none";
            const target = msg.querySelector('.text') || msg;
            clearHighlights(target);
            if (show && query) {
                applyHighlights(target, query);
            }
            const meta = msg.querySelector('.meta');
            if (meta) {
                clearHighlights(meta);
                if (show && senderQuery) {
                    applyHighlights(meta, senderQuery);
                }
            }
        });
    };

    // Initialize search event listeners
    if (input) {
        input.addEventListener("input", debounce(function() { updateSearch(); updateActiveFilters(); }, 400));
    }
    if (imageOnlyCheckbox) {
        imageOnlyCheckbox.addEventListener("change", function() { updateSearch(); updateActiveFilters(); });
    }
    if (codeOnlyCheckbox) {
        codeOnlyCheckbox.addEventListener("change", function() { updateSearch(); updateActiveFilters(); });
    }
    if (senderInput) {
        senderInput.addEventListener("input", debounce(function() { updateSearch(); updateActiveFilters(); }, 400));
    }
    
    // Date range filter listeners
    const dateFromInput = document.getElementById('dateFromInput');
    const dateToInput = document.getElementById('dateToInput');
    const dateRangeOnlyCheckbox = document.getElementById('dateRangeOnlyCheckbox');
    
    if (dateFromInput) {
        dateFromInput.addEventListener("change", function() { updateSearch(); updateActiveFilters(); });
    }
    if (dateToInput) {
        dateToInput.addEventListener("change", function() { updateSearch(); updateActiveFilters(); });
    }
    if (dateRangeOnlyCheckbox) {
        dateRangeOnlyCheckbox.addEventListener("change", function() { updateSearch(); updateActiveFilters(); });
    }
    
    updateSearch(); // Initial search update
});
</script>

<script>
/* === THEME TOGGLE === */
document.addEventListener("DOMContentLoaded", function() {
    const toggle = document.getElementById('darkModeToggle');
    if (!toggle) return;

    const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    const saved = localStorage.getItem('tce-theme');
    const initialDark = saved ? saved === 'dark' : prefersDark;

    function applyTheme(isDark) {
        document.body.classList.toggle('dark', isDark);
        localStorage.setItem('tce-theme', isDark ? 'dark' : 'light');
    }

    applyTheme(initialDark);
    toggle.checked = initialDark;

    toggle.addEventListener('change', function() {
        applyTheme(this.checked);
    });
});
</script>

<script>
/* === NAVIGATION BUTTONS FUNCTIONALITY === */
/**
 * Smart navigation buttons that appear/disappear based on content scrollability
 * Provides smooth scrolling to top/bottom with disabled states at boundaries
 */
document.addEventListener("DOMContentLoaded", function() {
    const scrollDownBtn = document.getElementById('scroll-down-btn');
    const scrollUpBtn = document.getElementById('scroll-up-btn');
    const content = document.querySelector('.content');

    if (!scrollDownBtn || !scrollUpBtn || !content) {
        console.log('Navigation buttons or content not found');
        return;
    }

    /**
     * Check if content area has a vertical scrollbar
     * @returns {boolean} True if content is scrollable
     */
    function hasVerticalScrollbar() {
        // Check both content div and window/document scrollability
        const contentScrollable = content.scrollHeight > content.clientHeight;
        const windowScrollable = document.body.scrollHeight > window.innerHeight;
        return contentScrollable || windowScrollable;
    }

    /**
     * Update navigation button visibility and disabled states
     * Hides buttons on cover page or when content is not scrollable
     */
    function updateButtonVisibility() {
        const hasScrollbar = hasVerticalScrollbar();
        const coverPage = document.getElementById('cover-page');
        const isCoverPageVisible = coverPage && coverPage.style.display !== 'none';
        
        console.log('Content height:', content.scrollHeight, 'Client height:', content.clientHeight);
        console.log('Body height:', document.body.scrollHeight, 'Window height:', window.innerHeight);
        console.log('Has scrollbar:', hasScrollbar);
        console.log('Cover page visible:', isCoverPageVisible);
        
        // Hide buttons if cover page is visible OR if there's no scrollbar
        const shouldShow = hasScrollbar && !isCoverPageVisible;
        scrollDownBtn.style.display = shouldShow ? 'block' : 'none';
        scrollUpBtn.style.display = shouldShow ? 'block' : 'none';
        
        if (shouldShow) {
            // Check scroll position and update button states
            let scrollTop, scrollHeight, clientHeight;
            
            // Determine which element is scrollable
            if (content.scrollHeight > content.clientHeight) {
                scrollTop = content.scrollTop;
                scrollHeight = content.scrollHeight;
                clientHeight = content.clientHeight;
            } else {
                scrollTop = window.pageYOffset || document.documentElement.scrollTop;
                scrollHeight = document.body.scrollHeight;
                clientHeight = window.innerHeight;
            }
            
            // Check if at top (tolerance of 5px)
            const atTop = scrollTop <= 5;
            // Check if at bottom (tolerance of 5px)
            const atBottom = scrollTop + clientHeight >= scrollHeight - 5;
            
            // Update button states based on scroll position
            if (atTop) {
                scrollUpBtn.classList.add('disabled');
            } else {
                scrollUpBtn.classList.remove('disabled');
            }
            
            if (atBottom) {
                scrollDownBtn.classList.add('disabled');
            } else {
                scrollDownBtn.classList.remove('disabled');
            }
            
            console.log('Scroll position:', scrollTop, 'At top:', atTop, 'At bottom:', atBottom);
        }
    }

    /**
     * Scroll to bottom button click handler
     * Prevents scrolling if button is disabled
     */
    scrollDownBtn.addEventListener('click', function() {
        // Don't scroll if button is disabled
        if (this.classList.contains('disabled')) {
            return;
        }
        
        console.log('Down button clicked');
        // Try both content scrolling and window scrolling
        if (content.scrollHeight > content.clientHeight) {
            content.scrollTo({
                top: content.scrollHeight,
                behavior: 'smooth'
            });
        } else {
            window.scrollTo({
                top: document.body.scrollHeight,
                behavior: 'smooth'
            });
        }
    });

    /**
     * Scroll to top button click handler
     * Prevents scrolling if button is disabled
     */
    scrollUpBtn.addEventListener('click', function() {
        // Don't scroll if button is disabled
        if (this.classList.contains('disabled')) {
            return;
        }
        
        console.log('Up button clicked');
        // Try both content scrolling and window scrolling
        if (content.scrollHeight > content.clientHeight) {
            content.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        } else {
            window.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        }
    });

    // Initial button visibility check
    setTimeout(updateButtonVisibility, 100);

    // Check button visibility when window is resized
    window.addEventListener('resize', updateButtonVisibility);

    // Check button visibility when content changes (e.g., when switching chats)
    const observer = new MutationObserver(function() {
        setTimeout(updateButtonVisibility, 100);
    });

    observer.observe(content, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['style']
    });

    // Update button visibility when scrolling
    content.addEventListener('scroll', updateButtonVisibility);
    window.addEventListener('scroll', updateButtonVisibility);

    // Also check when chats are shown
    const originalShowChat = window.showChat;
    window.showChat = function(id) {
        if (originalShowChat) {
            originalShowChat(id);
        }
        // Update button visibility after showing a chat (cover page is hidden)
        setTimeout(updateButtonVisibility, 200);
    };

    // Monitor cover page visibility changes to update button states
    const coverPage = document.getElementById('cover-page');
    if (coverPage) {
        const coverObserver = new MutationObserver(function(mutations) {
            mutations.forEach(function(mutation) {
                if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
                    setTimeout(updateButtonVisibility, 100);
                }
            });
        });
        coverObserver.observe(coverPage, { attributes: true });
    }

    // === HOME BUTTON FUNCTIONALITY ===
    /**
     * Navigate to home (cover page)
     */
    window.goHome = function() {
        const coverPage = document.getElementById('cover-page');
        if (coverPage) {
            // Show cover page
            coverPage.style.display = 'block';
            
            // Hide all chat sections using the .chat-section class
            document.querySelectorAll('.chat-section').forEach(section => {
                section.style.display = 'none';
            });
            
            // Reset sidebar active states
            document.querySelectorAll('.sidebar a').forEach(link => {
                link.classList.remove('active');
            });

            // Collapse all sidebar sections back to their headers
            document.querySelectorAll('.sidebar-section-content').forEach(sec => {
                sec.style.display = 'none';
                const header = sec.previousElementSibling;
                if (header && header.classList.contains('sidebar-section-header')) {
                    header.innerHTML = header.innerHTML.replace('▼', '▶');
                    header.dataset.expanded = "false";
                }
            });
            
            // Hide breadcrumb
            const breadcrumb = document.getElementById('breadcrumb');
            if (breadcrumb) {
                breadcrumb.classList.remove('show');
            }
            
            // Update button visibility
            setTimeout(updateButtonVisibility, 100);
        }
    };
    
    // Add event listener to home button
    const homeMenuBtn = document.getElementById('home-menu-btn');
    if (homeMenuBtn) {
        homeMenuBtn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            window.goHome();
        });
    }
});
</script>

<script>
/* === BREADCRUMB NAVIGATION FUNCTIONALITY === */
/**
 * Update breadcrumb navigation based on selected chat
 * Extracts team, channel, and topic information from the chat hierarchy
 * @param {string} chatId - The ID of the selected chat
 */
function updateBreadcrumb(chatId) {
    const breadcrumb = document.getElementById('breadcrumb');
    const chatElement = document.getElementById(chatId);
    
    if (!breadcrumb || !chatElement) return;
    
    // Get the chat name from the h2 header or data attribute
    const chatNameElement = chatElement.querySelector('h2');
    const chatName = chatNameElement ? chatNameElement.textContent : chatId;
    
    // Get the active sidebar link to determine hierarchy
    const activeLink = document.querySelector('.sidebar a.active');
    let team = '', channel = '', topic = chatName;
    let emoji = '💬';
    
    if (activeLink) {
        const sectionId = findParentSectionId(activeLink);
        
        // Determine category and emoji
        if (sectionId === 'oneonone-section') {
            emoji = '💬';
            topic = chatName;
        } else if (sectionId === 'group-section') {
            emoji = '👥';
            topic = chatName;
        } else if (sectionId === 'meeting-section') {
            emoji = '📅';
            topic = chatName;
        } else if (sectionId === 'channel-section') {
            emoji = '📢';
            // For channels, try to extract team and channel from hierarchy
            const teamHeader = findParentTeamHeader(activeLink);
            if (teamHeader) {
                team = teamHeader.textContent.trim();
            }
            channel = chatName;
            topic = '';
        }
    }
    
    // Build breadcrumb HTML
    let breadcrumbHtml = '<div class="breadcrumb-item"><a onclick="window.goHome(); return false;"><span class="breadcrumb-emoji">🏠</span>Home</a></div>';
    
    if (team && channel) {
        breadcrumbHtml += `<div class="breadcrumb-item"><span class="breadcrumb-emoji">${emoji}</span>${team}</div>`;
        breadcrumbHtml += `<div class="breadcrumb-item">${channel}</div>`;
    } else if (topic) {
        breadcrumbHtml += `<div class="breadcrumb-item"><span class="breadcrumb-emoji">${emoji}</span>${topic}</div>`;
    }
    
    breadcrumb.innerHTML = breadcrumbHtml;
    breadcrumb.classList.add('show');
}

/**
 * Find parent section ID for a sidebar link
 * @param {HTMLElement} element - The element to search from
 * @returns {string|null} The section ID or null if not found
 */
function findParentSectionId(element) {
    let parent = element.parentElement;
    while (parent) {
        if (parent.id && parent.id.endsWith('-section')) {
            return parent.id;
        }
        parent = parent.parentElement;
    }
    return null;
}

/**
 * Find parent team header for a channel link
 * @param {HTMLElement} element - The element to search from
 * @returns {HTMLElement|null} The team header element or null if not found
 */
function findParentTeamHeader(element) {
    let prev = element.previousElementSibling;
    while (prev) {
        if (prev.classList.contains('sidebar-section-header') && !prev.classList.contains('top-header')) {
            return prev;
        }
        prev = prev.previousElementSibling;
    }
    // Also search in parent
    let parent = element.parentElement;
    while (parent) {
        if (parent.id === 'channel-section') {
            // Find the header before this element
            prev = element.previousElementSibling;
            while (prev) {
                if (prev.classList.contains('sidebar-section-header')) {
                    return prev;
                }
                prev = prev.previousElementSibling;
            }
        }
        parent = parent.parentElement;
    }
    return null;
}
</script>

<script>
/* === CHAT EXPORT FUNCTIONALITY === */
// Create export toolbar and attach to the UI
document.addEventListener('DOMContentLoaded', function() {
    // Build toolbar element
    const toolbar = document.createElement('div');
    toolbar.id = 'export-toolbar';
    toolbar.style.cssText = 'display:none; gap:8px; align-items:center; padding:10px 20px; border-top:1px solid var(--input-border); border-bottom:1px solid var(--input-border); background: var(--content-bg); position: sticky; top: 0; z-index: 500;';
    toolbar.innerHTML = '<div style="font-weight:600; color: var(--text);">Export chat:</div>' +
        '<button id="export-txt" style="padding:6px 10px; border:1px solid var(--input-border); border-radius:6px; background: var(--input-bg); color: var(--text); cursor:pointer;">TXT</button>' +
        '<button id="export-md" style="padding:6px 10px; border:1px solid var(--input-border); border-radius:6px; background: var(--input-bg); color: var(--text); cursor:pointer;">MD</button>' +
        '<button id="export-html" style="padding:6px 10px; border:1px solid var(--input-border); border-radius:6px; background: var(--input-bg); color: var(--text); cursor:pointer;">Print HTML</button>' +
        '<div id="export-chat-name" style="margin-left:auto; color: var(--meta); font-size:0.9em;"></div>';

    // Do not insert yet; we'll attach under the active chat's heading when a chat is opened

    function sanitizeFileName(name) {
        return (name || 'chat').replace(/[^a-zA-Z0-9-_\. ]/g, '_').trim() || 'chat';
    }

    function getActiveSection() {
        let active = null;
        document.querySelectorAll('.chat-section').forEach(sec => {
            if (sec.style.display !== 'none') active = sec;
        });
        return active;
    }

    function extractMessages(section) {
        const msgs = [];
        if (!section) return msgs;
        section.querySelectorAll('.message').forEach(msg => {
            // Skip hidden messages
            if (msg.style.display === 'none') return;
            const metaEl = msg.querySelector('.meta');
            const textEl = msg.querySelector('.text');
            const sender = (msg.getAttribute('data-sender') || '').trim();
            const metaText = metaEl ? metaEl.textContent : '';
            // Extract timestamp
            let timestamp = '';
            const m = metaText.match(/(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})/);
            if (m) timestamp = m[1];
            // Content
            const html = textEl ? textEl.innerHTML : '';
            const plain = textEl ? textEl.textContent || '' : '';
            // Detect code block
            const hasBlockCode = html.includes('<pre');
            // Collect image srcs
            const imgs = [];
            if (textEl) {
                textEl.querySelectorAll('img').forEach(i => imgs.push(i.getAttribute('src')));
            }
            msgs.push({ sender, timestamp, html, plain, hasBlockCode, imgs });
        });
        return msgs;
    }

    function downloadString(filename, mime, content) {
        const blob = new Blob([content], { type: mime + ';charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            URL.revokeObjectURL(url);
            a.remove();
        }, 100);
    }

    function exportTxt() {
        const section = getActiveSection();
        if (!section) return;
        const chatName = section.querySelector('h2')?.textContent || 'Chat';
        const msgs = extractMessages(section);
        const lines = msgs.map(m => {
            const header = `${m.sender} • ${m.timestamp}`.trim();
            const imgLines = (m.imgs || []).map(src => src ? `[image] ${src}` : '').filter(Boolean);
            return header + '\n' + (m.plain || '').trim() + (imgLines.length ? '\n' + imgLines.join('\n') : '');
        }).join('\n\n');
        const fname = sanitizeFileName(chatName) + '.txt';
        downloadString(fname, 'text/plain', lines);
    }

    function exportMd() {
        const section = getActiveSection();
        if (!section) return;
        const chatName = section.querySelector('h2')?.textContent || 'Chat';
        const msgs = extractMessages(section);
        const md = [`# ${chatName}`].concat(msgs.map(m => {
            let body = (m.plain || '').trim();
            if (m.hasBlockCode && body) {
                body = '```\n' + body + '\n```';
            }
            const images = (m.imgs || []).map(src => src ? `![](${src})` : '').join('\n');
            const header = `**${m.sender}** • ${m.timestamp}`.trim();
            return header + '\n\n' + body + (images ? '\n\n' + images : '');
        })).join('\n\n---\n\n');
        const fname = sanitizeFileName(chatName) + '.md';
        downloadString(fname, 'text/markdown', md);
    }

    function exportHtml() {
        const section = getActiveSection();
        if (!section) return;
        const chatName = section.querySelector('h2')?.textContent || 'Chat';
        const msgs = extractMessages(section);
        const items = msgs.map(m => {
            const safeHtml = m.html || '';
            return `<div class="msg"><div class="meta">${m.sender} • ${m.timestamp}</div><div class="text">${safeHtml}</div></div>`;
        }).join('\n');
        const doc = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${chatName} - Export</title>` +
            `<style>body{font-family:system-ui,Segoe UI,Arial;padding:20px;background:#fff;color:#111} .meta{color:#555;margin-bottom:6px;font-size:0.9em} .msg{padding:10px 12px;border:1px solid #ddd;border-radius:8px;margin:8px 0;page-break-inside:avoid} .text img{max-width:100%;height:auto} h1{margin:0 0 12px 0}</style></head><body>` +
            `<h1>${chatName}</h1>` + items + `</body></html>`;
        const fname = sanitizeFileName(chatName) + '.html';
        downloadString(fname, 'text/html', doc);
    }

    // Attach click handlers (bind before DOM attachment)
    toolbar.querySelector('#export-txt')?.addEventListener('click', exportTxt);
    toolbar.querySelector('#export-md')?.addEventListener('click', exportMd);
    toolbar.querySelector('#export-html')?.addEventListener('click', exportHtml);

    // Expose helper to update toolbar visibility
    window.__updateExportToolbar = function(chatId) {
        const section = chatId ? document.getElementById(chatId) : getActiveSection();
        const name = section?.querySelector('h2')?.textContent || '';
        const label = toolbar.querySelector('#export-chat-name');
        if (label) label.textContent = name ? `Selected: ${name}` : '';
        if (section) {
            // Prefer placing under the Members section (if present), otherwise under the h2
            const members = section.querySelector('.chat-members');
            if (members) {
                members.insertAdjacentElement('afterend', toolbar);
            } else {
                const h2 = section.querySelector('h2');
                if (h2) {
                    h2.insertAdjacentElement('afterend', toolbar);
                } else {
                    // Fallback: append at the top of the section
                    section.insertAdjacentElement('afterbegin', toolbar);
                }
            }
            toolbar.style.display = 'flex';
        } else {
            toolbar.style.display = 'none';
        }
    };
});

// Ensure toolbar shows/hides with navigation
document.addEventListener('DOMContentLoaded', function() {
    // Wrap showChat to update export toolbar
    const prev = window.showChat;
    window.showChat = function(id) {
        if (typeof prev === 'function') prev(id);
        setTimeout(function(){ if (window.__updateExportToolbar) window.__updateExportToolbar(id); }, 50);
    };
    // Hide toolbar on home
    const prevHome = window.goHome;
    window.goHome = function() {
        if (typeof prevHome === 'function') prevHome();
        const tb = document.getElementById('export-toolbar');
        if (tb) tb.style.display = 'none';
    };
});
</script>

<script>
/* === SIDEBAR SECTION TOGGLE FUNCTIONALITY === */
/**
 * Toggle visibility of sidebar sections (expand/collapse)
 * @param {string} sectionId - ID of the section to toggle
 */
function toggleSection(sectionId) {
    var section = document.getElementById(sectionId);
    var header = section.previousElementSibling;
    var isOpening = section.style.display === "none";
    section.style.display = isOpening ? "block" : "none";
    header.innerHTML = header.innerHTML.replace(isOpening ? '▶' : '▼', isOpening ? '▼' : '▶');
    header.dataset.expanded = isOpening ? "true" : "false";
}

/**
 * Copy message content to clipboard with visual feedback
 * @param {HTMLElement} button - The copy button element
 */
function copyMessageFromBtn(button) {
    var messageId = button.getAttribute('data-message-id');
    var plainText = button.getAttribute('data-message-text');
    
    // Decode HTML entities
    var textToCopy = plainText
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&');
    
    // Try using the modern Clipboard API
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(textToCopy).then(function() {
            // Show success feedback
            var originalText = button.textContent;
            button.textContent = '✓';
            button.style.color = '#22c55e';
            setTimeout(function() {
                button.textContent = originalText;
                button.style.color = '';
            }, 2000);
        }).catch(function(err) {
            // Fallback if clipboard API fails
            fallbackCopyMessage(textToCopy);
        });
    } else {
        // Fallback for older browsers
        fallbackCopyMessage(textToCopy);
    }
}

/**
 * Fallback copy method for browsers without Clipboard API
 * @param {string} text - Text to copy
 */
function fallbackCopyMessage(text) {
    var textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.left = '-999999px';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    try {
        document.execCommand('copy');
        alert('Message copied to clipboard!');
    } catch (err) {
        console.error('Failed to copy:', err);
    }
    document.body.removeChild(textArea);
}

/**
 * Initialize sidebar sections in collapsed state
 */
document.addEventListener("DOMContentLoaded", function() {
    // Ensure all sections are collapsed by default
    document.querySelectorAll('.sidebar-section-content').forEach(function(sec) {
        sec.style.display = "none";
    });
    document.querySelectorAll('.sidebar-section-header').forEach(function(header) {
        if (!header.innerHTML.trim().startsWith('▶')) {
            header.innerHTML = '▶ ' + header.innerHTML.trim().replace(/^▼|▶/, '');
        }
    });
});
</script>

</head>
<body>
<!-- === SIDEBAR STRUCTURE === -->
<div class="sidebar" style="font-size: 0.8em;">

<!-- === FILTER PANEL === -->
<div style="padding: 12px; border-bottom: 1px solid var(--sidebar-border);">
  <p style="margin: 0 0 12px 0; font-size: 0.9em; color: var(--sidebar-muted);">Exported on <!--EXPORT_DATE--></p>
  
  <!-- Main Search -->
  <input type="text" id="searchInput" placeholder="🔍 Search messages..." style="width: 100%; padding: 10px; margin-bottom: 10px; box-sizing: border-box; background: var(--input-bg); border: 1px solid var(--input-border); border-radius: 6px; color: var(--input-text); font-size: 0.9em;">
  
  <!-- Hidden checkboxes for filter state -->
  <input type="checkbox" id="imageOnlyCheckbox" style="display: none;">
  <input type="checkbox" id="codeOnlyCheckbox" style="display: none;">
  
  <!-- Quick Filters Row -->
  <div style="display: flex; gap: 6px; margin-bottom: 10px; flex-wrap: wrap;">
    <button id="imageFilterBtn" class="quick-filter-tag" onclick="toggleImageFilter()" style="padding: 6px 12px; background: var(--sidebar-hover); border: 1px solid var(--sidebar-border); border-radius: 20px; color: var(--sidebar-text); cursor: pointer; font-size: 0.85em; transition: all 0.2s;" title="Show only messages with images">🖼️ Images</button>
    <button id="codeFilterBtn" class="quick-filter-tag" onclick="toggleCodeFilter()" style="padding: 6px 12px; background: var(--sidebar-hover); border: 1px solid var(--sidebar-border); border-radius: 20px; color: var(--sidebar-text); cursor: pointer; font-size: 0.85em; transition: all 0.2s;" title="Show only messages with code blocks">💻 Code</button>
  </div>
  
  <!-- Active Filters Display -->
  <div id="activeFiltersContainer" style="display: none; margin-bottom: 10px; padding: 8px; background: rgba(59, 130, 246, 0.1); border-radius: 6px; border-left: 3px solid #3b82f6;">
    <div style="font-size: 0.75em; color: var(--sidebar-muted); margin-bottom: 6px; font-weight: 600; text-transform: uppercase;">Active Filters</div>
    <div id="activeFiltersList" style="display: flex; gap: 6px; flex-wrap: wrap;"></div>
  </div>
  
  <!-- Advanced Filters Toggle -->
  <button onclick="toggleAdvancedFilters()" style="width: 100%; padding: 8px; background: var(--sidebar-hover); border: 1px solid var(--sidebar-border); color: var(--sidebar-text); border-radius: 6px; cursor: pointer; font-size: 0.9em; margin-bottom: 10px; transition: all 0.2s;">⚙️ Advanced Filters <span id="advFilterToggleIcon">▶</span></button>
  
  <!-- Advanced Filters Panel -->
  <div id="advancedFiltersPanel" style="display: none; padding: 10px; background: var(--sidebar-hover); border-radius: 6px; border: 1px solid var(--sidebar-border); margin-bottom: 10px;">
    <!-- Sender filter -->
    <label style="display: block; font-size: 0.8em; color: var(--sidebar-muted); margin-bottom: 4px; font-weight: 600;">👤 Sender</label>
    <input type="text" id="senderFilterInput" placeholder="Filter by sender (e.g., 'alex')" style="width: 100%; padding: 8px; margin-bottom: 10px; box-sizing: border-box; background: var(--input-bg); border: 1px solid var(--sidebar-border); border-radius: 4px; color: var(--input-text); font-size: 0.85em;">
    
    <!-- Date range filter -->
    <label style="display: block; font-size: 0.8em; color: var(--sidebar-muted); margin-bottom: 6px; font-weight: 600;">📅 Date Range</label>
    <div style="display: flex; gap: 6px; margin-bottom: 6px; align-items: center;">
      <div style="flex: 1; position: relative; display: flex; align-items: center;">
        <input type="date" id="dateFromInput" placeholder="From" style="flex: 1; padding: 8px 32px 8px 8px; box-sizing: border-box; background: var(--input-bg); border: 1px solid var(--sidebar-border); border-radius: 4px; color: var(--input-text); font-size: 0.85em; color-scheme: dark; z-index: 1;">
        <button type="button" onclick="document.getElementById('dateFromInput').click(); return false;" style="position: absolute; right: 0px; top: 0px; bottom: 0px; width: 32px; background: none; border: none; cursor: pointer; font-size: 1.2em; color: var(--sidebar-text); margin: 0; z-index: 10; pointer-events: auto; display: flex; align-items: center; justify-content: center;" title="Open date picker">📆</button>
      </div>
      <div style="flex: 1; position: relative; display: flex; align-items: center;">
        <input type="date" id="dateToInput" placeholder="To" style="flex: 1; padding: 8px 32px 8px 8px; box-sizing: border-box; background: var(--input-bg); border: 1px solid var(--sidebar-border); border-radius: 4px; color: var(--input-text); font-size: 0.85em; color-scheme: dark; z-index: 1;">
        <button type="button" onclick="document.getElementById('dateToInput').click(); return false;" style="position: absolute; right: 0px; top: 0px; bottom: 0px; width: 32px; background: none; border: none; cursor: pointer; font-size: 1.2em; color: var(--sidebar-text); margin: 0; z-index: 10; pointer-events: auto; display: flex; align-items: center; justify-content: center;" title="Open date picker">📆</button>
      </div>
    </div>
    <label style="display: flex; align-items: center; color: var(--sidebar-text); margin-bottom: 10px; font-size: 0.85em; cursor: pointer;">
      <input type="checkbox" id="dateRangeOnlyCheckbox" checked style="margin-right: 6px; cursor: pointer;">
      Apply date filter
    </label>
  </div>
  
  <!-- Clear All Filters Button -->
  <button onclick="clearAllFilters()" id="clearAllBtn" style="width: 100%; padding: 8px; background: #ef4444; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 0.9em; font-weight: 600; transition: all 0.2s; margin-bottom: 12px; display: none;">🗑️ Clear All Filters</button>
  
  <!-- Dark Mode Toggle -->
  <div style="display: flex; align-items: center; padding: 8px 0; border-top: 1px solid var(--sidebar-border); margin-top: 10px;">
    <label class="switch" aria-label="Toggle dark mode" style="margin-right: 10px;">
      <input type="checkbox" id="darkModeToggle">
      <span class="slider"></span>
    </label>
    <span style="font-size: 0.9em;">Dark mode</span>
  </div>
  
  <!-- Home Button -->
  <a id="home-menu-btn" href="#" style="display: block; padding: 10px; margin-top: 10px; background: var(--sidebar-active); border-radius: 6px; color: var(--sidebar-text); text-decoration: none; text-align: center; font-weight: 500; transition: all 0.2s ease; cursor: pointer; font-size: 0.9em;" onmouseover="this.style.background='var(--sidebar-hover)'" onmouseout="this.style.background='var(--sidebar-active)'">
    🏠 Home
  </a>
</div>

<!-- Chat navigation sections will be dynamically populated here -->
<!-- Navigation buttons will be added after the content div -->
"""