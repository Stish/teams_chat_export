# Microsoft Teams Chat Export

A powerful Python script that exports Microsoft Teams chats, group conversations, meetings, and channel messages to a single, searchable HTML file with embedded images and intuitive navigation.

## 📋 Table of Contents

- [Features](#-features)
- [Screenshots](#-screenshots)
- [Requirements](#-requirements)
- [Installation](#-installation)
- [Configuration](#-configuration)
- [Usage](#-usage)
- [Output](#-output)
- [Advanced Features](#-advanced-features)
- [Troubleshooting](#-troubleshooting)
- [Contributing](#-contributing)
- [License](#-license)

## ✨ Features

### Core Functionality
- **Complete Export**: Exports all one-on-one chats, group chats, meeting chats, and team channel messages
- **Image Support**: Downloads and embeds images locally for offline viewing
- **HTML Output**: Generates a single, self-contained HTML file with all your Teams data
- **Search Functionality**: Built-in search across all messages and conversations
- **Responsive Design**: Works on desktop and mobile devices

### User Experience
- **Sidebar Navigation**: Organized sidebar with collapsible sections for easy browsing
- **Message Filtering**: Automatically filters out system messages and empty content
- **Lightbox Images**: Click images to view them in a fullscreen lightbox
- **Scroll Navigation**: Quick scroll-to-top and scroll-to-bottom buttons
- **Member Lists**: Shows chat participants with tooltips for group conversations

### Advanced Features
- **Ignore Lists**: Configure which channels and chats to skip during export
- **Automatic User Detection**: Automatically identifies your display name from Microsoft Graph API
- **Progress Tracking**: Real-time progress updates during export process
- **Error Handling**: Robust error handling with graceful fallbacks
- **Compact Display**: Optimized HTML output for minimal file size and clean appearance

## 📸 Screenshots

*The exported HTML file provides a clean, Teams-like interface with:*
- **Sidebar Navigation**: Browse all your chats and channels
- **Message Display**: Clean message layout with timestamps and sender information
- **Search Bar**: Filter messages across all conversations
- **Image Gallery**: Embedded images with lightbox viewing

## 🔧 Requirements

### System Requirements
- **Python**: 3.6 or higher
- **Operating System**: Windows, macOS, or Linux
- **Internet Connection**: Required for Microsoft Graph API access

### Python Dependencies
```
requests>=2.25.0
```

### Microsoft Graph API Access
- Valid Microsoft Graph API access token with the following permissions:
  - `Chat.ReadWrite`
  - `ChannelMessage.Read.All`
  - `Team.ReadBasic.All`
  - `User.Read`
  - `Directory.Read.All`

## 🚀 Installation

### 1. Clone or Download
```bash
git clone <repository-url>
# OR download the ZIP file and extract it
```

### 2. Install Dependencies
```bash
pip install requests
```

### 3. File Structure
Ensure your directory contains these files:
```
restTeams/
├── teams_chat_export.py    # Main script
├── config.py              # Configuration file
├── html_template.py       # HTML template
├── teams_utils.py         # Utility functions
└── README.md             # This file
```

## ⚙️ Configuration

### 1. Get Microsoft Graph API Access Token

**Option A: Using Graph Explorer (Recommended for testing)**
1. Go to [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in with your Microsoft account
3. Run any query to generate a token
4. Copy the token from the request headers

**Option B: Azure App Registration (Recommended for production)**
1. Create an Azure AD app registration
2. Configure required permissions (see Requirements section)
3. Generate an access token using your preferred authentication flow

### 2. Configure the Script

Edit `config.py` and update the following:

```python
# Replace with your actual access token
ACCESS_TOKEN = 'your_access_token_here'

# Update output folder path (optional)
OUTPUT_FOLDER = 'C:\\your\\desired\\output\\path\\'

# Configure ignore lists (optional)
IGNORED_CHANNELS = [
    ("Team Name", "Channel Name"),  # Channels to skip
]

IGNORED_CHATS = [
    "Chat Name to Skip",  # Chats to skip
]
```

### 3. Important Configuration Notes

- **Access Token**: The script will automatically detect your user display name from the token
- **Paths**: Use absolute paths for `OUTPUT_FOLDER` to avoid issues
- **Ignore Lists**: Use these to skip large or unimportant channels/chats
- **Case Sensitivity**: Team/channel names in ignore lists are case-sensitive

## 📖 Usage

### Basic Usage

1. **Configure your access token** in `config.py`
2. **Run the script**:
   ```bash
   python teams_chat_export.py
   ```
3. **Wait for completion** - the script will show progress updates
4. **Open the generated HTML file** in your browser

### Command Line Output Example

```
### Microsoft Teams Chat Export Started
Script Version: v0.1.2
✅ Fetched user display name: Your Name
✅ Access token valid - authenticated as: Your Name
##  Fetching all chats
Found 45 total chats
#   Processing: Chat with John Doe
#   Processing: Group: Team Meeting
Processed 12 one-on-one chats
Processed 8 group chats
Processed 3 meeting chats
##  Fetching teams and channels
#   Processing team: Development Team
#     Processing channel: General...(156 total, 142 with content)
##  Generating HTML export
##  Processing chat messages
#   Fetching messages for chat: Chat with John Doe...(89 messages)
✅ Successfully exported to 'C:\path\to\output\index.html'
### Export completed successfully!
```

### Advanced Usage

**Selective Export with Ignore Lists:**
```python
# In config.py
IGNORED_CHANNELS = [
    ("Large Team", "General"),      # Skip busy general channel
    ("Announcements", "Company"),   # Skip announcement channel
]

IGNORED_CHATS = [
    "Bot Notifications",            # Skip automated chats
]
```

## 📂 Output

### Generated Files

The script creates the following structure in your output folder:

```
output/
├── index.html           # Main export file (open this in browser)
├── img/                # Downloaded images folder
│   ├── image1.jpg
│   ├── image2.png
│   └── ...
```

### HTML File Features

- **Self-contained**: All CSS and JavaScript embedded
- **Offline viewing**: Images stored locally
- **Responsive**: Works on all screen sizes
- **Searchable**: Real-time message filtering
- **Navigable**: Sidebar with expandable sections

## 🔧 Advanced Features

### Ignore Lists

Configure channels and chats to skip during export:

```python
# Skip specific team channels
IGNORED_CHANNELS = [
    ("Marketing Team", "Random"),
    ("Development", "Build-Notifications"),
]

# Skip specific chats
IGNORED_CHATS = [
    "Automated Reports",
    "System Notifications",
]
```

### Automatic User Detection

The script automatically:
- Fetches your display name from Microsoft Graph API
- Identifies your messages in conversations
- Applies appropriate styling (your messages vs. others)

### Progress Tracking

Real-time updates show:
- Number of chats/channels being processed
- Message counts for each conversation
- Export progress and completion status

### Error Handling

The script handles:
- Invalid or expired access tokens
- Network connectivity issues
- Malformed message data
- Missing images or attachments

## 🔍 Troubleshooting

### Common Issues

**1. "Access token validation failed"**
```
❌ Access token validation failed: Token is invalid or expired
```
**Solution**: Generate a new access token and update `config.py`

**2. "Permission denied" errors**
```
❌ Error: Insufficient privileges to complete the operation
```
**Solution**: Ensure your access token has the required Microsoft Graph permissions

**3. "No messages found"**
```
Found 0 total chats
```
**Solution**: 
- Verify you have Teams conversations
- Check if your token has `Chat.ReadWrite` permission
- Ensure you're signed in to the correct Microsoft account

**4. Images not displaying**
```
Images show as broken links in the HTML file
```
**Solution**:
- Check internet connectivity during export
- Verify the `img/` folder exists in the output directory
- Ensure sufficient disk space for image downloads

### Debug Mode

To enable detailed error messages, uncomment this line in `teams_chat_export.py`:
```python
# Uncomment for debugging
print(f"Error processing message: {e}")
```

### Token Permissions

Required Microsoft Graph permissions:
- `Chat.ReadWrite` - Read chat messages
- `ChannelMessage.Read.All` - Read channel messages  
- `Team.ReadBasic.All` - List teams and channels
- `User.Read` - Get user profile information
- `Directory.Read.All` - Read user directory (for member names)

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

### Development Setup

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

### Code Style

- Follow PEP 8 Python style guidelines
- Add docstrings for new functions
- Include type hints where appropriate
- Test with different Teams configurations

## 📄 License

This project is provided as-is for educational and personal use. Please ensure you comply with your organization's data export policies and Microsoft's terms of service when using this script.

## 🔗 Links

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
- [Azure App Registration Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)

---

**Version**: v0.1.2  
**Last Updated**: July 2025  
**Tested with**: Microsoft Teams Web, Desktop App  
**Python Compatibility**: 3.6+
