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
Version: v0.1.1
Last Updated: 2025-01-17
"""

# HTML template with embedded CSS and JavaScript
# This template provides the complete web interface for the Teams chat export
html_content = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Teams Chat Export</title>
<style>
/* === BASE STYLES === */
body { font-family: Arial, sans-serif; margin: 0; display: flex; background: #f4f4f4;}

/* === SIDEBAR STYLES === */
.sidebar { width: 300px; color: white; position: fixed; top: 0; bottom: 0; overflow-y: auto; }
/* Sidebar base */
.sidebar {
    background: #23272e;
}
.sidebar a { display: block; padding: 10px; color: white; text-decoration: none; border-bottom: 1px solid #444; }
.sidebar a:hover { background: #444; }
.sidebar a.active { background: #555; font-weight: bold; }

/* === CONTENT AREA STYLES === */
.content {
    flex: 1;
    padding: 20px;
    background-color: #f4f4f4;
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
    padding: 8px 12px;
    border-radius: 10px;
    margin: 2px 0;
    max-width: 60%;
    clear: both;
    word-wrap: break-word;
    white-space: pre-wrap;
    max-width: 900px;
    word-break: break-word;
}
.mine { background-color: #d4f8d4; float: right; text-align: right; }
.theirs { background-color: #d4e6f8; float: left; text-align: left; }
.meta { font-size: 0.8em; color: #555; margin-bottom: 3px; }
.text { margin: 0; padding: 0; }
.clearfix::after { content: ""; clear: both; display: table; }

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
    background: #222;
    padding: 8px 10px;
    font-weight: bold;
    border-bottom: 1px solid #444;
    user-select: none;
}
.sidebar-section-header:hover {
    background: #444;
}
.sidebar-section-content {
    padding-left: 0;
}

/* === CHAT TYPE SPECIFIC STYLES === */
/* One on One Chats */
#oneonone-section.sidebar-section-content {
    background: #50546A;
}
#oneonone-section a {
    background: #50546A;
    color: #fff;
}

/* Group Chats */
#group-section.sidebar-section-content {
    background: #50546A;
}
#group-section a {
    background: #50546A;
    color: #fff;
}

/* Meeting Chats */
#meeting-section.sidebar-section-content {
    background: #50546A;
}
#meeting-section a {
    background: #50546A;
    color: #fff;
}

/* Channel Chats main section */
#channel-section.sidebar-section-content {
    background: #3b3e4a;
}

/* Team headers under Channel Chats */
#channel-section .sidebar-section-header {
    background: #363944;
    color: #e0e6f0;
    padding-left: 18px;
}

/* Channel links under teams */
#channel-section .sidebar-section-content a {
    background: #50546a;
    color: #e0e6f0;
    padding-left: 32px;
}

/* === HOVER AND ACTIVE STATES === */
.sidebar a:hover,
.sidebar-section-header:hover {
    background: #5c6370 !important;
}
.sidebar a.active {
    background: #7c83a0 !important;
    color: #fff !important;
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
    background-color: #4a5568;
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
    background-color: #5a6578;
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
    background-color: #4a5568;
    transform: none;
}

/* Button positioning */
#scroll-down-btn {
    bottom: 20px;
}

#scroll-up-btn {
    top: 20px;
}
</style>

<script>
/* === CHAT NAVIGATION FUNCTIONALITY === */
/**
 * Shows the specified chat section and updates active states
 * @param {string} id - The ID of the chat section to show
 */
function showChat(id) {
    document.querySelectorAll('.sidebar a').forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('onclick')?.includes("'" + id + "'")) {
            link.classList.add('active');
        }
    });
    document.querySelectorAll('.chat-section').forEach(div => div.style.display = 'none');
    document.getElementById(id).style.display = 'block';
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
 * Advanced search functionality with debouncing and image filtering
 * Provides real-time search across all chat messages with result counts
 */
document.addEventListener("DOMContentLoaded", function() {
    const input = document.getElementById("searchInput");
    const imageOnlyCheckbox = document.getElementById("imageOnlyCheckbox");
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
     * Update search results based on current query and filters
     */
    function updateSearch() {
        const query = input.value.toLowerCase();
        const imageOnly = imageOnlyCheckbox.checked;

        /**
         * Process a specific chat section and update match counts
         * @param {string} sectionId - ID of the section to process
         * @param {string} sectionName - Display name of the section
         */
        function processSection(sectionId, sectionName) {
            const section = document.getElementById(sectionId);
            const links = section.querySelectorAll('a');
            let totalMatches = 0;

            links.forEach(link => {
                const chatId = link.getAttribute("onclick").match(/'(.*?)'/)[1];
                const sectionDiv = document.getElementById(chatId);
                let matchCount = 0;

                sectionDiv.querySelectorAll(".message").forEach(msg => {
                    // Cache lowercased text once
                    if (!msg.dataset.lowerText) {
                        msg.dataset.lowerText = msg.innerText.toLowerCase();
                    }
                    const hasImg = messageHasImage(msg);
                    const matchText = msg.dataset.lowerText.includes(query);
                    let show = false;
                    if (imageOnly) {
                        if (query === "") {
                            show = hasImg;
                        } else {
                            show = hasImg && matchText;
                        }
                    } else {
                        show = (query === "" || matchText);
                    }
                    msg.style.display = show ? "block" : "none";
                    if (show) matchCount++;
                });

                if (matchCount > 0) {
                    link.style.display = "";
                    link.innerText = link.getAttribute("data-chat-name") + (matchCount > 0 ? ` (${matchCount})` : "");
                } else {
                    link.style.display = "none";
                }

                totalMatches += matchCount;
            });

            const header = section.previousElementSibling;
            let arrow = header.innerText.trim().charAt(0);
            if (arrow !== '▶' && arrow !== '▼') arrow = '▶';
            if (totalMatches > 0) {
                header.innerText = `${arrow} ${sectionName} (${totalMatches})`;
            } else {
                header.innerText = `${arrow} ${sectionName}`;
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
            const teamDiv = header.nextElementSibling;
            const channelLinks = teamDiv ? teamDiv.querySelectorAll('a') : [];
            let teamMatchCount = 0;

            channelLinks.forEach(link => {
                const chatId = link.getAttribute("onclick").match(/'(.*?)'/)[1];
                const sectionDiv = document.getElementById(chatId);
                let matchCount = 0;

                if (sectionDiv) {
                    sectionDiv.querySelectorAll(".message").forEach(msg => {
                        if (!msg.dataset.lowerText) {
                            msg.dataset.lowerText = msg.innerText.toLowerCase();
                        }
                        const hasImg = messageHasImage(msg);
                        const matchText = msg.dataset.lowerText.includes(query);
                        let show = false;
                        if (imageOnly) {
                            if (query === "") {
                                show = hasImg;
                            } else {
                                show = hasImg && matchText;
                            }
                        } else {
                            show = (query === "" || matchText);
                        }
                        msg.style.display = show ? "block" : "none";
                        if (show) matchCount++;
                    });
                }

                if (matchCount > 0) {
                    link.style.display = "";
                    link.innerText = link.getAttribute("data-chat-name") + (matchCount > 0 ? ` (${matchCount})` : "");
                } else {
                    link.style.display = "none";
                }

                teamMatchCount += matchCount;
            });

            const origName = header.innerText.replace(/^[▶▼]\s*/, '').replace(/\s*\(\d+\)$/, '');
            let arrow = header.innerText.trim().charAt(0);
            if (arrow !== '▶' && arrow !== '▼') arrow = '▶';
            if (teamMatchCount > 0) {
                header.innerText = `${arrow} ${origName} (${teamMatchCount})`;
            } else {
                header.innerText = `${arrow} ${origName}`;
            }

            totalChannelMatches += teamMatchCount;
        });

        const channelHeader = channelSection ? channelSection.previousElementSibling : null;
        if (channelHeader) {
            let arrow = channelHeader.innerText.trim().charAt(0);
            if (arrow !== '▶' && arrow !== '▼') arrow = '▶';
            const name = "Channel Chats";
            if (totalChannelMatches > 0) {
                channelHeader.innerText = `${arrow} ${name} (${totalChannelMatches})`;
            } else {
                channelHeader.innerText = `${arrow} ${name}`;
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
        const imageOnly = imageOnlyCheckbox.checked;
        section.querySelectorAll(".message").forEach(msg => {
            const text = msg.innerText.toLowerCase();
            const hasImg = msg.querySelector("img") !== null;
            const match = text.includes(query);
            let show = false;
            if (imageOnly) {
                if (query === "") {
                    show = hasImg;
                } else {
                    show = hasImg && match;
                }
            } else {
                show = (query === "" || match);
            }
            msg.style.display = show ? "block" : "none";
        });
    };

    // Initialize search event listeners
    if (input) {
        input.addEventListener("input", debounce(updateSearch, 400));
    }
    if (imageOnlyCheckbox) {
        imageOnlyCheckbox.addEventListener("change", updateSearch);
    }
    updateSearch(); // Initial search update
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
    if (section.style.display === "none") {
        section.style.display = "block";
        header.innerHTML = header.innerHTML.replace('▶', '▼');
    } else {
        section.style.display = "none";
        header.innerHTML = header.innerHTML.replace('▼', '▶');
    }
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
<p style="padding: 10px; font-size: 0.9em; color: #ccc;">Exported on <!--EXPORT_DATE--></p>

<!-- Search functionality -->
<input type="text" id="searchInput" placeholder="Search all messages..." style="width: 100%; padding: 8px; margin-bottom: 10px; box-sizing: border-box;">
<!-- Image filter checkbox -->
<label style="display:block; color:#ccc; margin:3px 0 10px 0; font-size: 0.95em;">
  <input type="checkbox" id="imageOnlyCheckbox" style="vertical-align:middle; margin-right:3px;">
  Show only messages with images
</label>

<!-- Chat navigation sections will be dynamically populated here -->
<!-- Navigation buttons will be added after the content div -->
"""