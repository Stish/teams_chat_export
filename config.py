"""
Configuration file for Microsoft Teams Chat Export Script

This file contains all configuration parameters and constants used by the Teams export script.
Modify these values according to your needs.

Author: Alexander Wegner
Version: v0.1.1
"""

# =============================================================================
# AUTHENTICATION CONFIGURATION
# =============================================================================

# Microsoft Graph API access token
# Replace with your actual token from Graph Explorer or Azure AD app registration
ACCESS_TOKEN = 'YOUR_ACCESS_TOKEN_HERE'

# User display name will be automatically fetched from Microsoft Graph API
# No need to manually configure this - it will be determined from the access token
USER_DISPLAY_NAME = None  # Will be set automatically during initialization

# =============================================================================
# PATH CONFIGURATION
# =============================================================================

# Output folder for generated files
OUTPUT_FOLDER = 'C:\\work\\_scripts\\py\\teams_chat_export\\output\\'

# Image folder name (relative to output folder)
IMAGE_FOLDER = 'img'

# =============================================================================
# SCRIPT CONFIGURATION
# =============================================================================

# Script version
SCRIPT_VERSION = 'v0.1.2'

# =============================================================================
# API CONFIGURATION
# =============================================================================

# Microsoft Graph API endpoints
GRAPH_API_BASE_URL = 'https://graph.microsoft.com/v1.0'

# API request headers
def get_api_headers(access_token):
    """Get standard headers for Microsoft Graph API requests."""
    return {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json',
        'ConsistencyLevel': 'eventual'
    }

# =============================================================================
# PAGINATION CONFIGURATION
# =============================================================================

# Number of items to fetch per API call (for pagination)
ITEMS_PER_PAGE = 50

# =============================================================================
# HTML GENERATION CONFIGURATION
# =============================================================================

# Navigation buttons HTML
NAVIGATION_BUTTONS_HTML = '''
<!-- Navigation buttons -->
<button id="scroll-up-btn" class="nav-button" title="Scroll to top">↑</button>
<button id="scroll-down-btn" class="nav-button" title="Scroll to bottom">↓</button>
'''

# Output HTML file name
OUTPUT_HTML_FILE = 'index.html'

# =============================================================================
# CHAT TYPES
# =============================================================================

# Teams chat types
CHAT_TYPE_ONE_ON_ONE = 'oneOnOne'
CHAT_TYPE_GROUP = 'group'
CHAT_TYPE_MEETING = 'meeting'

# Message types
MESSAGE_TYPE_USER = 'message'

# =============================================================================
# DISPLAY LIMITS
# =============================================================================

# Maximum number of member names to display in group chat names
MAX_MEMBERS_IN_CHAT_NAME = 3

# =============================================================================
# IGNORE LIST CONFIGURATION
# =============================================================================

# List of team/channel combinations to ignore during message fetching
# Format: [("Team Name", "Channel Name"), ...]
# These channels will still appear in the sidebar but will show an ignore message
# instead of fetching actual messages
# 
# Examples:
# - To ignore the "General" channel in "Employee Platform" team:
#   ("Employee Platform", "General")
# - To ignore the "Random" channel in "Development Team" team:
#   ("Development Team", "Random")
#
IGNORED_CHANNELS = [
    ("Employee Platform", "General"),
    # Add more team/channel combinations here as needed
    # ("Another Team", "Another Channel"),
    # ("Development Team", "Random"),
]

# List of chat names to ignore during message fetching
# Format: ["Chat Name", ...]
# These chats will still appear in the sidebar but will show an ignore message
# instead of fetching actual messages
IGNORED_CHATS = [
    # Add chat names here as needed
    # "Chat with Someone",
    # "Group: Name1, Name2, Name3",
]
