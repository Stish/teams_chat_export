"""
Configuration file for Microsoft Teams Chat Export Script

This file contains all configuration parameters and constants used by the Teams export script.
Modify these values according to your needs.

Author: Alexander Wegner
Version: v0.1.7
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

# Download images from messages (set to False to use placeholder images instead)
# When True: Images are downloaded and embedded in the export
# When False: Images are replaced with a placeholder/dummy image with a message indicating they were not downloaded
DOWNLOAD_IMAGES = True

# =============================================================================
# SCRIPT CONFIGURATION
# =============================================================================

# Script version
SCRIPT_VERSION = 'v0.1.7'

# =============================================================================
# API CONFIGURATION
# =============================================================================

# Microsoft Graph API endpoints
GRAPH_API_BASE_URL = 'https://graph.microsoft.com/v1.0'

# =============================================================================
# PAGINATION CONFIGURATION
# =============================================================================

# Number of items to fetch per API call (for pagination)
ITEMS_PER_PAGE = 50

# =============================================================================
# HTML GENERATION CONFIGURATION
# =============================================================================

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
# MESSAGE LIMIT CONFIGURATION (FOR TESTING/DEVELOPMENT)
# =============================================================================

# Limit the number of messages fetched per chat (None = no limit, fetch all)
# Set to a low number (e.g., 50, 100) for faster testing
# Set to None for production use to fetch all messages
MESSAGES_LIMIT_PER_CHAT = None  # Examples: 50, 100, 500, or None for unlimited

# Limit the total number of chats to process (None = no limit, process all)
# Set to a low number (e.g., 5, 10) for faster testing
# Set to None for production use to process all chats
CHATS_LIMIT = None  # Examples: 5, 10, 20, or None for unlimited

# Limit the number of channels to process per team (None = no limit)
# For quick testing you can set a small number (e.g., 3 or 5)
# If left as None, the script will fall back to CHATS_LIMIT when set
CHANNELS_LIMIT_PER_TEAM = None  # Examples: 3, 5, 10, or None for unlimited

# =============================================================================
# MESSAGE DATE RANGE CONFIGURATION
# =============================================================================

# Filter messages by date range (optional)
# Format: "YYYY-MM-DD" or None for no date limit
# Only messages within the specified date range will be included in the export
# 
# Examples:
# - Both dates set: Only messages between these dates (inclusive) will be exported
# - Only MESSAGE_DATE_FROM set: All messages from this date onwards will be exported
# - Only MESSAGE_DATE_TO set: All messages up to and including this date will be exported
# - Both None: No date filtering, all messages exported (default)
#
MESSAGE_DATE_FROM = None # e.g., "2024-01-15" or None for no lower limit
MESSAGE_DATE_TO = None # e.g., "2024-12-31" or None for no upper limit
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
