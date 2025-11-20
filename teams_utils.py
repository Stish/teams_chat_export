"""
Microsoft Teams Chat Export Utility Functions

This module contains utility functions for processing Microsoft Teams chat exports,
including message processing, image handling, and API interactions.

Author: Alexander Wegner
Version: v0.1.1
"""

import requests
import html
import re
import os
from datetime import datetime


def replace_emoji_tags(text):
    """
    Replace <emoji> tags in the given text with their alt attribute value.

    This function searches for all <emoji> tags in the input string and replaces each tag
    with the value of its alt attribute, effectively converting embedded emoji HTML tags
    to their textual representation.

    Args:
        text (str): The input string containing <emoji> tags.

    Returns:
        str: The modified string with <emoji> tags replaced by their alt text.
    """
    return re.sub(r'<emoji[^>]*alt="([^"]+)"[^>]*>.*?</emoji>', r'\1', text)


def download_image_from_src(src_url, token, folder, message_id, index):
    """
    Download an image from Microsoft Graph API and save it locally.

    Downloads an image from the specified source URL using a bearer token for authorization,
    saves it to the given folder with a filename based on the message ID and index, and returns
    the relative path to the saved image.

    Args:
        src_url (str): The URL of the image to download.
        token (str): The bearer token used for authorization in the request header.
        folder (str): The local directory where the image will be saved.
        message_id (str): The identifier used to construct the image filename.
        index (int): The index used to differentiate multiple images per message.

    Returns:
        str: The relative path (from the location of 'index.html') to the saved image file.
    """
    filename = f"{message_id}_img{index}.jpg"
    local_path = os.path.join(folder, filename)
    
    # Only download if file doesn't exist
    if not os.path.exists(local_path):
        img_headers = {'Authorization': f'Bearer {token}'}
        try:
            img_response = requests.get(src_url, headers=img_headers)
            if img_response.status_code == 200:
                with open(local_path, 'wb') as f:
                    f.write(img_response.content)
        except Exception as e:
            print(f"Error downloading image {src_url}: {e}")
            return None
    
    return os.path.relpath(local_path, start=os.path.dirname('index.html')).replace('\\', '/')


def clean_img_tags(content):
    """
    Clean <img> tags in HTML content by preserving only the src attribute.

    This function searches for all <img> tags in the input string and rewrites them to retain 
    only the src attribute, removing any other attributes that may be present.

    Args:
        content (str): The HTML content containing <img> tags to be cleaned.

    Returns:
        str: The modified HTML content with cleaned <img> tags.
    """
    return re.sub(r'<img[^>]*src="([^"]+)"[^>]*>', r'<img src="\1">', content)


def wrap_images_with_lightbox(content):
    """
    Wrap all <img> tags in HTML content with <a> tags for lightbox functionality.

    This function searches for all <img> tags in the provided HTML string and wraps each image
    with an <a> tag that links to the image source and adds a "lightbox" CSS class. This enables
    lightbox effects when the images are clicked.

    Args:
        content (str): The HTML content containing <img> tags to be wrapped.

    Returns:
        str: The modified HTML content with each <img> tag wrapped in an <a> tag for lightbox support.
    """
    return re.sub(r'<img src="([^"]+)"\s*/?>', r'<a href="\1" class="lightbox"><img src="\1"></a>', content)


def chat_has_messages(chat_id, headers):
    """
    Check if a Microsoft Teams chat has at least one user message.

    This function queries the Microsoft Graph API for the first 10 messages in the specified chat.
    It returns True if there is at least one message of type 'message' (i.e., a user message),
    otherwise returns False.

    Args:
        chat_id (str): The ID of the chat to check.
        headers (dict): The HTTP headers for the API request, including authorization.

    Returns:
        bool: True if the chat contains at least one user message, False otherwise.
    """
    url = f'https://graph.microsoft.com/v1.0/chats/{chat_id}/messages?$top=10'
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            messages = data.get('value', [])
            for msg in messages:
                msg_type = msg.get('messageType', '')
                if msg_type == 'message':
                    return True
    except Exception as e:
        print(f"Error checking messages for chat {chat_id}: {e}")
    
    return False


def sort_chats_by_name(chat_list):
    """
    Sort chats by name (case-insensitive).

    Args:
        chat_list (list of tuple): List of (chat_name, chat_id) tuples.

    Returns:
        list of tuple: Sorted list of (chat_name, chat_id) tuples, ordered alphabetically by chat_name.
    """
    return sorted(chat_list, key=lambda x: x[0].lower())


def process_message_content(raw_content, message_id, access_token, image_folder):
    """
    Process message content by cleaning HTML, replacing emojis, and handling images.

    Args:
        raw_content (str): The raw message content from Teams API.
        message_id (str): The unique identifier for the message.
        access_token (str): The bearer token for API authentication.
        image_folder (str): The folder path for saving images.

    Returns:
        str: The processed and cleaned message content.
    """
    # Unescape HTML entities
    clean_content = html.unescape(raw_content)
    
    # Replace emoji tags with their alt text
    clean_content = replace_emoji_tags(clean_content)

    # Process images: download and replace URLs
    img_matches = re.findall(r'<img[^>]+src="([^"]+)"', clean_content)
    for idx, img_url in enumerate(img_matches):
        if img_url.startswith("https://graph.microsoft.com"):
            local_img_path = download_image_from_src(img_url, access_token, image_folder, message_id, idx)
            if local_img_path:
                clean_content = clean_content.replace(img_url, local_img_path)

    # Clean img tags and wrap with lightbox
    clean_content = clean_img_tags(clean_content)
    clean_content = wrap_images_with_lightbox(clean_content)

    return clean_content


def format_timestamp(timestamp_str):
    """
    Format a timestamp string to a readable format.

    Args:
        timestamp_str (str): The timestamp string from Teams API.

    Returns:
        str: Formatted timestamp string or the original if formatting fails.
    """
    try:
        return datetime.fromisoformat(timestamp_str.replace('Z', '+00:00')).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return timestamp_str


def create_member_list_display(member_names, max_display=3):
    """
    Create a display-friendly member list with truncation for long lists.

    Args:
        member_names (list): List of member names.
        max_display (int): Maximum number of names to display before truncation.

    Returns:
        tuple: (limited_names, full_member_list) for display and tooltip.
    """
    if not member_names:
        return "Unknown", "Unknown"
    
    # Create limited display string
    limited_names = ', '.join(member_names[:max_display])
    full_member_list = ', '.join(member_names)
    
    # Add ellipsis if truncated
    if len(member_names) > max_display:
        limited_names += "..."
    
    return limited_names, full_member_list


def format_member_list_for_display(full_members):
    """
    Format member list for display by replacing every second comma with semicolon.

    Args:
        full_members (str): Comma-separated list of member names.

    Returns:
        str: Formatted member list with semicolons for better readability.
    """
    if not full_members:
        return ""
    
    parts = full_members.split(', ')
    new_members = []
    for idx, part in enumerate(parts):
        if idx > 0:
            if idx % 2 == 0:
                new_members.append('; ' + part)
            else:
                new_members.append(', ' + part)
        else:
            new_members.append(part)
    
    return ''.join(new_members)


def validate_message_content(msg):
    """
    Validate if a message has text content or images and is not a system message.

    Args:
        msg (dict): Message object from Teams API.

    Returns:
        bool: True if message has text or images and is from a real user, False otherwise.
    """
    # Skip non-message types (system events, etc.)
    if msg.get('messageType') != 'message':
        return False
    
    # Skip messages from Unknown users (system messages)
    from_data = msg.get('from')
    if from_data and isinstance(from_data, dict):
        user_data = from_data.get('user')
        if user_data and isinstance(user_data, dict):
            sender = user_data.get('displayName', 'Unknown')
        else:
            sender = 'Unknown'
    else:
        sender = 'Unknown'
    
    # Skip messages from Unknown users
    if sender == 'Unknown':
        return False
    
    # Check for actual content
    raw_content = msg.get('body', {}).get('content') or ''
    
    # Skip empty messages or messages with only system event tags
    if not raw_content or not raw_content.strip():
        return False
    
    # Skip messages that only contain system event tags
    if '<systemEventMessage/>' in raw_content and len(raw_content.strip()) <= 50:
        return False
    
    # Check for meaningful text content (not just whitespace/HTML tags)
    # Remove HTML tags and check if there's actual text
    import re
    text_only = re.sub(r'<[^>]+>', '', raw_content).strip()
    has_text = bool(text_only)
    
    # Check for images
    img_matches = re.findall(r'<img[^>]+src="([^"]+)"', raw_content)
    has_images = bool(img_matches)
    
    return has_text or has_images
