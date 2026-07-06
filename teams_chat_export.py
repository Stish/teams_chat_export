#!/usr/bin/env python3
"""
Microsoft Teams Chat Export Script

This script exports Microsoft Teams chats, group chats, meetings, and channel messages
to a single HTML file with navigation, search functionality, and image support.

Features:
- Exports one-on-one chats, group chats, and meeting chats
- Exports team channel messages
- Downloads and embeds images locally
- Generates HTML with sidebar navigation
- Includes search functionality and lightbox for images
- Responsive design with navigation buttons

Requirements:
- Microsoft Graph API access token
- Python 3.6+
- requests library

Usage:
    python teams_chat_export.py

Author: Alexander Wegner
Version: v0.1.8
"""

import os
import re
import sys
import urllib.parse
from datetime import datetime
from typing import Dict, List, Tuple, Optional

import requests

# Import local modules
from config import (
    ACCESS_TOKEN, OUTPUT_FOLDER, IMAGE_FOLDER, SCRIPT_VERSION,
    GRAPH_API_BASE_URL, ITEMS_PER_PAGE, OUTPUT_HTML_FILE,
    CHAT_TYPE_ONE_ON_ONE, CHAT_TYPE_GROUP, CHAT_TYPE_MEETING, MESSAGE_TYPE_USER,
    MAX_MEMBERS_IN_CHAT_NAME, IGNORED_CHANNELS, IGNORED_CHATS,
    MESSAGES_LIMIT_PER_CHAT, CHATS_LIMIT, CHANNELS_LIMIT_PER_TEAM, DOWNLOAD_IMAGES,
    MESSAGE_DATE_FROM, MESSAGE_DATE_TO
)
from html_template import html_content, NAVIGATION_BUTTONS_HTML
from html_template_multi import SHARED_CSS, INDEX_HTML, STATS_HTML, CHAT_HTML
from fragmented_export import FragmentedExporter
from teams_utils import (
    get_api_headers, chat_has_messages, sort_chats_by_name, process_message_content,
    format_timestamp, create_member_list_display, format_member_list_for_display,
    validate_message_content, is_message_in_date_range
)


class TeamsExporter:
    """
    Main class for exporting Microsoft Teams chats and channels to HTML.
    """
    
    def __init__(self):
        """Initialize the Teams exporter with configuration."""
        self.access_token = ACCESS_TOKEN
        self.output_folder = OUTPUT_FOLDER
        self.image_folder = IMAGE_FOLDER
        self.script_version = SCRIPT_VERSION
        self.headers = get_api_headers(self.access_token)
        
        # Fetch user display name from Microsoft Graph API
        self.user_display_name = self._fetch_user_display_name()
        
        # Data storage
        self.chats_one_on_one: Dict[str, str] = {}
        self.chats_group: Dict[str, str] = {}
        self.chats_meeting: Dict[str, str] = {}
        self.group_full_member_lists: Dict[str, str] = {}
        self.meeting_full_member_lists: Dict[str, str] = {}
        self.channels_by_team: Dict[str, List[Tuple[str, str, str]]] = {}
        self.channel_messages: Dict[Tuple[str, str], List[Dict]] = {}
        self.chat_message_counts: Dict[str, Dict[str, int]] = {
            'oneonone': {},
            'group': {},
            'meeting': {}
        }
        self.channel_message_counts: Dict[Tuple[str, str], int] = {}
        
        # Setup output directory
        self._setup_output_directory()
        
        # Statistics tracking
        self.message_dates = []  # Track all message dates
        self.participant_message_counts = {}  # Track message count per participant
    
    def _fetch_user_display_name(self) -> str:
        """
        Fetch the current user's display name from Microsoft Graph API.
        
        Returns:
            str: The user's display name, or 'Unknown User' if fetch fails
        """
        try:
            user_url = f'{GRAPH_API_BASE_URL}/me'
            response = requests.get(user_url, headers=self.headers)
            
            if response.status_code == 200:
                user_info = response.json()
                display_name = user_info.get('displayName', 'Unknown User')
                print(f"✅ Fetched user display name: {display_name}")
                return display_name
            else:
                print(f"⚠️  Failed to fetch user display name (HTTP {response.status_code}), using fallback")
                return 'Unknown User'
                
        except Exception as e:
            print(f"⚠️  Error fetching user display name: {e}, using fallback")
            return 'Unknown User'
    
    def _setup_output_directory(self):
        """Create output directory structure."""
        os.makedirs(os.path.join(self.output_folder, self.image_folder), exist_ok=True)
        os.chdir(self.output_folder)
        print(f"Output directory: {self.output_folder}")
    
    def _validate_access_token(self):
        """
        Validate the access token by making a test API call.
        
        Returns:
            bool: True if token is valid, False otherwise
        """
        try:
            # Test the token with a simple API call
            test_url = f'{GRAPH_API_BASE_URL}/me'
            response = requests.get(test_url, headers=self.headers)

            if response.status_code == 200:
                # Token is valid - user display name was already fetched during initialization
                print(f"✅ Access token valid - authenticated as: {self.user_display_name}")
                return True
            elif response.status_code == 401:
                print("❌ Access token validation failed: Token is invalid or expired")
                print("   Please check your access token in config.py")
                return False
            else:
                print(f"❌ Access token validation failed: HTTP {response.status_code}")
                print("   Please check your access token and network connection")
                return False
                
        except Exception as e:
            print(f"❌ Error validating access token: {e}")
            print("   Please check your access token and network connection")
            return False
    
    def _make_paginated_request(self, url: str, progress_callback=None, context="", limit=None) -> List[Dict]:
        """
        Make paginated requests to Microsoft Graph API.
        
        Args:
            url (str): The initial API URL
            progress_callback (callable): Optional callback for progress updates
            context (str): Context string for progress display
            limit (int): Optional limit on number of items to fetch (None = no limit)
            
        Returns:
            List[Dict]: Combined results from all pages
        """
        all_items = []
        
        while url:
            try:
                response = requests.get(url, headers=self.headers)
                if response.status_code == 200:
                    data = response.json()
                    batch = data.get('value', [])
                    all_items.extend(batch)
                    
                    # Call progress callback if provided
                    if progress_callback:
                        progress_callback(len(all_items), context)
                    
                    # Check if limit reached
                    if limit and len(all_items) >= limit:
                        all_items = all_items[:limit]
                        break
                    
                    url = data.get('@odata.nextLink')
                else:
                    print(f"Error in API request: {response.status_code}")
                    break
            except Exception as e:
                print(f"Exception in API request: {e}")
                break
        
        return all_items
    
    def _get_chat_members(self, chat_id: str) -> List[Dict]:
        """
        Get members of a specific chat.
        
        Args:
            chat_id (str): The chat ID
            
        Returns:
            List[Dict]: List of chat members
        """
        url = f'{GRAPH_API_BASE_URL}/chats/{chat_id}/members'
        try:
            response = requests.get(url, headers=self.headers)
            if response.status_code == 200:
                members = response.json().get('value', [])
                print(f'  ✅ Member API returned {len(members)} members')
                return members
            else:
                print(f'  ⚠️  Member API returned HTTP {response.status_code}')
        except Exception as e:
            print(f"  ⚠️  Error fetching members for chat {chat_id}: {e}")
        return []
    
    def _extract_other_user_from_messages(self, chat_id: str) -> str:
        """
        Extract the other user's display name from messages in a 1:1 chat.
        Used as fallback when member lookup fails.
        
        Args:
            chat_id (str): The chat ID
            
        Returns:
            str: The other user's display name, or 'Unknown' if not found
        """
        url = f'{GRAPH_API_BASE_URL}/chats/{chat_id}/messages?$top=50'
        try:
            response = requests.get(url, headers=self.headers)
            if response.status_code == 200:
                messages = response.json().get('value', [])
                print(f'  📧 Fetched {len(messages)} messages from chat {chat_id}')
                for msg in messages:
                    from_data = msg.get('from')
                    if from_data and isinstance(from_data, dict):
                        user_data = from_data.get('user')
                        if user_data and isinstance(user_data, dict):
                            sender = user_data.get('displayName', '').strip()
                            print(f'    Checking sender: "{sender}" (current user: "{self.user_display_name}")')
                            # Return the first sender that isn't us (case-sensitive check)
                            if sender and sender != self.user_display_name and sender != 'Unknown':
                                print(f'  ✅ Extracted user "{sender}" from messages (member lookup failed)')
                                return sender
                print(f'  ⚠️  No valid sender found in {len(messages)} messages')
            else:
                print(f'  ⚠️  Failed to fetch messages: HTTP {response.status_code}')
        except Exception as e:
            print(f"⚠️  Error extracting user from messages for chat {chat_id}: {e}")
        
        return 'Unknown'
    
    def _process_one_on_one_chat(self, chat: Dict):
        """Process a one-on-one chat."""
        chat_id = chat.get('id')
        chat_name = str(chat.get('topic') or '')
        
        if not chat_has_messages(chat_id, self.headers):
            return
        
        # If chat name is empty, use the other member's display name
        if not chat_name:
            members = self._get_chat_members(chat_id)
            print(f'  👥 Found {len(members)} members in chat')
            other_member = next(
                (m for m in members if m.get('displayName') != self.user_display_name), 
                None
            )
            if other_member:
                print(f'    Member found: {other_member.get("displayName")}')
            else:
                print(f'    No members found (all matched current user or list empty)')
            
            display_name = (other_member.get('displayName', 'Unknown') if other_member else 'Unknown') or 'Unknown'
            
            # If member lookup failed, try to extract from messages
            if not display_name or display_name == 'Unknown':
                print(f'  🔍 Member lookup failed (got "{display_name}"), attempting to extract from messages...')
                display_name = self._extract_other_user_from_messages(chat_id)
            
            chat_name = f"Chat with {display_name}"
        
        print(f'#   Processing: {chat_name}')
        self.chats_one_on_one[chat_name] = chat_id
    
    def _process_group_chat(self, chat: Dict):
        """Process a group chat."""
        chat_id = chat.get('id')
        chat_name = str(chat.get('topic') or '')
        
        if not chat_has_messages(chat_id, self.headers):
            return
        
        # Get all members for the chat
        members = self._get_chat_members(chat_id)
        other_members = [
            str(m.get('displayName', 'Unknown') or 'Unknown') 
            for m in members 
            if m.get('displayName') != self.user_display_name
        ]
        
        limited_names, full_member_list = create_member_list_display(
            other_members, MAX_MEMBERS_IN_CHAT_NAME
        )
        
        if not chat_name:
            chat_name = f"Group: {limited_names}" if limited_names else "Group: Unknown"
        
        print(f'#   Processing: {chat_name}')
        self.chats_group[chat_name] = chat_id
        self.group_full_member_lists[chat_name] = full_member_list
    
    def _process_meeting_chat(self, chat: Dict):
        """Process a meeting chat."""
        chat_id = chat.get('id')
        chat_name = str(chat.get('topic') or '').strip()

        # Some meeting chats have no topic at all; keep names stable and non-null.
        if not chat_name:
            chat_name = f"Meeting Chat {chat_id}"
        
        if not chat_has_messages(chat_id, self.headers):
            return
        
        print(f'#   Processing: {chat_name}')
        self.chats_meeting[chat_name] = chat_id
        
        # Get all members for the meeting
        members = self._get_chat_members(chat_id)
        member_names = [
            str(m.get('displayName', 'Unknown') or 'Unknown') 
            for m in members
        ]
        full_member_list = ', '.join(member_names) if member_names else "Unknown"
        self.meeting_full_member_lists[chat_name] = full_member_list
    
    def fetch_all_chats(self):
        """Fetch all chats from Microsoft Teams."""
        print('##  Fetching all chats')
        
        url = f'{GRAPH_API_BASE_URL}/chats?$top={ITEMS_PER_PAGE}'
        all_chats = self._make_paginated_request(url)
        
        print(f'Found {len(all_chats)} total chats')
        
        # Apply CHATS_LIMIT if configured
        if CHATS_LIMIT:
            all_chats = all_chats[:CHATS_LIMIT]
            print(f'⚠️  CHATS_LIMIT is set to {CHATS_LIMIT}. Processing only {len(all_chats)} chats.')
        
        # Process chats by type
        for chat in all_chats:
            chat_type = chat.get('chatType')
            
            if chat_type == CHAT_TYPE_ONE_ON_ONE:
                self._process_one_on_one_chat(chat)
            elif chat_type == CHAT_TYPE_GROUP:
                self._process_group_chat(chat)
            elif chat_type == CHAT_TYPE_MEETING:
                self._process_meeting_chat(chat)
        
        # Sort chats by name
        self.chats_one_on_one = dict(sort_chats_by_name(self.chats_one_on_one.items()))
        self.chats_group = dict(sort_chats_by_name(self.chats_group.items()))
        self.chats_meeting = dict(sort_chats_by_name(self.chats_meeting.items()))
        
        print(f'Processed {len(self.chats_one_on_one)} one-on-one chats')
        print(f'Processed {len(self.chats_group)} group chats')
        print(f'Processed {len(self.chats_meeting)} meeting chats')
    
    def fetch_teams_and_channels(self):
        """Fetch teams and their channels."""
        print('##  Fetching teams and channels')
        
        teams_url = f'{GRAPH_API_BASE_URL}/me/joinedTeams'
        teams = self._make_paginated_request(teams_url)
        
        # Apply CHATS_LIMIT to teams if configured (to skip channel processing when testing)
        if CHATS_LIMIT:
            teams = teams[:CHATS_LIMIT]
            print(f'⚠️  CHATS_LIMIT is set to {CHATS_LIMIT}. Processing only {len(teams)} teams.')

        channel_limit = CHANNELS_LIMIT_PER_TEAM if CHANNELS_LIMIT_PER_TEAM is not None else CHATS_LIMIT
        
        for team in teams:
            team_id = team.get('id')
            team_name = team.get('displayName', f"Team {team_id}")
            print(f'#   Processing team: {team_name}')
            
            # Get channels for this team
            channels_url = f'{GRAPH_API_BASE_URL}/teams/{team_id}/channels'
            channels = self._make_paginated_request(channels_url)

            # Apply channel limit per team (falls back to CHATS_LIMIT when set)
            if channel_limit:
                channels = channels[:channel_limit]
                print(f'⚠️  Channel limit is set to {channel_limit}. Processing only {len(channels)} channels in {team_name}.')
            
            self.channels_by_team[team_name] = []
            
            for channel in channels:
                channel_id = channel.get('id')
                channel_name = channel.get('displayName', f"Channel {channel_id}")
                
                # Progress tracking for channel messages
                all_messages = []
                
                def channel_progress_callback(current_total, context):
                    if current_total % 20 == 0 or current_total < 20:
                        # Count filtered messages in real-time from all_messages
                        filtered_count = sum(1 for msg in all_messages[:current_total] if validate_message_content(msg))
                        print(f'\r#     Processing channel: {channel_name}...({current_total} total, {filtered_count} with content)', end='', flush=True)
                
                print(f'#     Processing channel: {channel_name}...', end='', flush=True)
                
                self.channels_by_team[team_name].append((channel_name, team_id, channel_id))
                
                # Check if this channel should be ignored
                if (team_name, channel_name) in IGNORED_CHANNELS:
                    print(f'\r#     Processing channel: {channel_name}...(IGNORED - skipping message fetch)')
                    # Add ignore placeholder message
                    ignored_message = [{
                        'id': 'ignored_channel',
                        'body': {'content': f'This channel has been ignored due to configuration settings.<br><br>Team: {team_name}<br>Channel: {channel_name}<br><br>To enable message fetching, remove this channel from the IGNORED_CHANNELS list in config.py.'},
                        'createdDateTime': datetime.now().isoformat(),
                        'from': {'user': {'displayName': 'System'}}
                    }]
                    self.channel_messages[(team_name, channel_name)] = ignored_message
                    continue
                
                # Fetch messages for this channel with progress updates
                messages_url = f'{GRAPH_API_BASE_URL}/teams/{team_id}/channels/{channel_id}/messages'
                
                # Custom paginated request with progress for channels
                url = messages_url
                while url:
                    try:
                        response = requests.get(url, headers=self.headers)
                        if response.status_code == 200:
                            data = response.json()
                            batch = data.get('value', [])
                            all_messages.extend(batch)
                            
                            # Update progress every 20 messages
                            if len(all_messages) % 20 == 0 or len(all_messages) < 20:
                                filtered_count = sum(1 for msg in all_messages if validate_message_content(msg))
                                print(f'\r#     Processing channel: {channel_name}...({len(all_messages)} total, {filtered_count} with content)', end='', flush=True)
                            
                            url = data.get('@odata.nextLink')
                        else:
                            print(f"Error in API request: {response.status_code}")
                            break
                    except Exception as e:
                        print(f"Exception in API request: {e}")
                        break
                
                # Filter for user messages with content
                filtered_messages = [msg for msg in all_messages if validate_message_content(msg)]
                
                print(f'\r#     Processing channel: {channel_name}...({len(all_messages)} total, {len(filtered_messages)} with content)')
                
                # Add placeholder if no messages
                if not filtered_messages:
                    filtered_messages = [{
                        'id': 'no_messages',
                        'body': {'content': 'No messages found in this channel.'},
                        'createdDateTime': datetime.now().isoformat(),
                        'from': {'user': {'displayName': 'System'}}
                    }]
                
                self.channel_messages[(team_name, channel_name)] = filtered_messages
    
    def _fetch_chat_messages(self, chat_name: str, chat_id: str) -> List[Dict]:
        """
        Fetch all messages for a specific chat.
        
        Args:
            chat_name (str): The display name of the chat
            chat_id (str): The unique identifier of the chat
            
        Returns:
            List[Dict]: List of message objects
        """
        print(f'#   Fetching messages for chat: {chat_name}...', end='', flush=True)
        
        # Progress tracking for chat messages
        def chat_progress_callback(current_count, context):
            if current_count % 20 == 0 or current_count < 20:
                print(f'\r#   Fetching messages for chat: {chat_name}...({current_count} messages)', end='', flush=True)
        
        url = f'{GRAPH_API_BASE_URL}/chats/{chat_id}/messages'
        messages = self._make_paginated_request(url, chat_progress_callback, f"chat_{chat_name}", limit=MESSAGES_LIMIT_PER_CHAT)
        
        print(f'\r#   Fetching messages for chat: {chat_name}...({len(messages)} messages)')
        return messages
    
    def _generate_sidebar_section(
        self,
        section_id: str,
        title: str,
        chats: Dict[str, str],
        member_lists: Optional[Dict[str, str]] = None,
        message_counts: Optional[Dict[str, int]] = None
    ) -> str:
        """
        Generate HTML for a sidebar section.
        
        Args:
            section_id (str): The HTML ID for the section
            title (str): The display title for the section
            chats (Dict[str, str]): Dictionary of chat names to IDs
            member_lists (Optional[Dict[str, str]]): Member lists for tooltips
            
        Returns:
            str: HTML content for the sidebar section
        """
        # Filter out chats with 0 messages if message counts are provided
        filtered_chats = {}
        for chat_name, chat_id in chats.items():
            if message_counts is None or message_counts.get(chat_name, 0) > 0:
                filtered_chats[chat_name] = chat_id
        
        # If no chats with messages, return empty string (skip this section)
        if not filtered_chats:
            return ''
        
        header_title = title
        if message_counts is not None:
            total_count = sum(message_counts.get(name, 0) for name in filtered_chats.keys())
            header_title = f"{title} ({total_count})"

        # Add category hint classes/attributes for styling and icons
        category = section_id.split('-')[0] if '-' in section_id else section_id
        html = f'''
    <div class="sidebar-section-header top-header {section_id}-header" data-cat="{category}" onclick="toggleSection('{section_id}')"><div class="header-content">{header_title}</div><span class="toggle-icon">+</span></div>
    <div id="{section_id}" class="sidebar-section-content" style="display:none;">
'''
        
        for chat_name, chat_id in filtered_chats.items():
            safe_chat_name = urllib.parse.quote(str(chat_name or ''), safe='')
            tooltip = ""
            
            if member_lists and chat_name in member_lists:
                tooltip = f'title="{member_lists[chat_name]}"'
            
            count_suffix = ''
            if message_counts is not None:
                count_suffix = f" ({message_counts.get(chat_name, 0)})"
            html += f'''    <a href="#" onclick="showChat('{safe_chat_name}')" data-chat-name="{chat_name}" {tooltip}>{chat_name}{count_suffix}</a>
'''
        
        html += '  </div>\n'
        return html
    
    def _generate_message_html(self, msg: Dict, chat_type: str = 'chat') -> str:
        """
        Generate HTML for a single message.
        
        Args:
            msg (Dict): Message object from Teams API
            chat_type (str): Type of chat ('chat' or 'channel')
            
        Returns:
            str: HTML content for the message
        """
        try:
            # Safely extract message data with proper null checks
            message_id = msg.get('id', '')
            
            # Safely extract sender information
            from_data = msg.get('from')
            if from_data and isinstance(from_data, dict):
                user_data = from_data.get('user')
                if user_data and isinstance(user_data, dict):
                    sender = user_data.get('displayName', 'Unknown')
                else:
                    sender = 'Unknown'
            else:
                sender = 'Unknown'
            
            # Safely extract message content
            body_data = msg.get('body')
            if body_data and isinstance(body_data, dict):
                raw_content = body_data.get('content', '')
            else:
                raw_content = ''
            
            # Process message content
            clean_content = process_message_content(
                raw_content, message_id, self.access_token, self.image_folder, DOWNLOAD_IMAGES
            )
            
            # Skip messages with no content after processing
            # Check for meaningful content: either text or images
            content_check = re.sub(r'<[^>]+>', '', clean_content or '').strip()
            has_text = bool(content_check)
            has_images = bool(re.search(r'<img[^>]+>', clean_content or ''))
            
            if not clean_content or (not has_text and not has_images):
                return ""
            
            # Format timestamp
            timestamp = msg.get('lastModifiedDateTime', msg.get('createdDateTime', ''))
            formatted_timestamp = format_timestamp(timestamp)
            
            # Track statistics: message date and participant count
            try:
                if timestamp:
                    # Parse and store message date
                    self.message_dates.append(timestamp)
                    # Count messages per participant (excluding own messages)
                    if sender and sender != self.user_display_name:
                        self.participant_message_counts[sender] = self.participant_message_counts.get(sender, 0) + 1
            except Exception:
                pass  # Silently ignore stats tracking errors
            
            # Determine CSS class based on sender
            css_class = 'mine' if sender == self.user_display_name else 'theirs'

            # Safely include sender as data attribute for client-side filtering
            sender_attr = (sender or 'Unknown').replace('"', '&quot;')
            
            # Generate avatar HTML with initials and color (only for received messages)
            initials = ''.join([n[0].upper() for n in (sender or '?').split() if n][:2]) or '?'
            avatar_html = f'<div class="avatar" style="background-color: hsl({hash(sender) % 360}, 70%, 55%);">{initials}</div>' if css_class == 'theirs' else ''
            
            # Generate unique message ID for copy functionality
            message_uid = f"msg-{message_id}"
            
            # Extract plain text from HTML content for copy
            plain_text = re.sub(r'<[^>]+>', '', clean_content or '').strip()
            # Escape HTML entities for data attribute
            plain_text_escaped = plain_text.replace('&', '&amp;').replace('"', '&quot;').replace("'", '&#39;').replace('<', '&lt;').replace('>', '&gt;')
            
            return f'<div class="clearfix"><div class="message {css_class}" data-sender="{sender_attr}" id="{message_uid}"><div class="message-content"><div class="message-header">{avatar_html}<div class="meta">{sender} • {formatted_timestamp}<button class="copy-btn" data-message-id="{message_uid}" data-message-text="{plain_text_escaped}" onclick="copyMessageFromBtn(this)" title="Copy message">📋</button></div></div><div class="text">{clean_content}</div></div></div></div>\n'
        except Exception as e:
            # Silently handle errors and return empty string to avoid console spam
            # Uncomment the line below if you want to see the error details for debugging
            # print(f"Error processing message: {e}")
            return ""

    def _calculate_statistics(self, html_content: str = None) -> Dict:
        """
        Calculate statistics for the export.
        
        Args:
            html_content (str): Optional HTML content string to calculate size. 
                              If None, will estimate based on images only.
        
        Returns:
            Dict: Statistics including total messages, chats, date range, storage size, and top participants
        """
        import os
        from datetime import datetime
        
        # Count total messages
        total_messages = (sum(self.chat_message_counts.get('oneonone', {}).values()) +
                         sum(self.chat_message_counts.get('group', {}).values()) +
                         sum(self.chat_message_counts.get('meeting', {}).values()) +
                         sum(self.channel_message_counts.values()))
        
        # Count total chats
        total_chats = (len(self.chats_one_on_one) + len(self.chats_group) + 
                      len(self.chats_meeting) + len(self.channel_message_counts))
        
        # Calculate storage size and count images
        total_size = 0
        total_images = 0
        
        # Add HTML content size (if provided, use actual size; otherwise estimate)
        if html_content:
            total_size += len(html_content.encode('utf-8'))
        
        # Add image folder size
        image_folder_path = os.path.join(self.output_folder, self.image_folder)
        if os.path.exists(image_folder_path):
            for dirpath, dirnames, filenames in os.walk(image_folder_path):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    total_size += os.path.getsize(filepath)
                    total_images += 1
        
        # Convert to readable format
        if total_size < 1024:
            size_str = f"{total_size} B"
        elif total_size < 1024 * 1024:
            size_str = f"{total_size / 1024:.1f} KB"
        else:
            size_str = f"{total_size / (1024 * 1024):.1f} MB"
        
        # Extract date range and count messages per participant
        earliest_date = None
        latest_date = None
        participant_counts = {}
        
        # Process collected dates to find range
        if self.message_dates:
            self.message_dates.sort()
            earliest_iso = self.message_dates[0]
            latest_iso = self.message_dates[-1]
            
            try:
                earliest_dt = datetime.fromisoformat(earliest_iso.replace('Z', '+00:00'))
                latest_dt = datetime.fromisoformat(latest_iso.replace('Z', '+00:00'))
                earliest_date = earliest_dt.strftime('%B %d, %Y')
                latest_date = latest_dt.strftime('%B %d, %Y')
            except Exception:
                earliest_date = None
                latest_date = None
        
        # Get top 10 most active participants
        top_participants = []
        if self.participant_message_counts:
            sorted_participants = sorted(
                self.participant_message_counts.items(),
                key=lambda x: x[1],
                reverse=True
            )[:10]
            top_participants = [
                {'name': name, 'count': count}
                for name, count in sorted_participants
            ]
        
        return {
            'total_messages': total_messages,
            'total_chats': total_chats,
            'total_size': size_str,
            'total_images': total_images,
            'one_on_one': len(self.chats_one_on_one),
            'group': len(self.chats_group),
            'meeting': len(self.chats_meeting),
            'channels': len(self.channel_message_counts),
            'earliest_date': earliest_date,
            'latest_date': latest_date,
            'top_participants': top_participants
        }

    def _get_date_label(self, timestamp: str) -> str:
        """Return a friendly date label (day-level) for separators."""
        try:
            dt = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
            return dt.strftime('%A, %b %d, %Y')
        except Exception:
            return 'Unknown date'
    
    def _generate_participant_list_html(self, top_participants: List[Dict]) -> str:
        """
        Generate HTML for top participants list.
        
        Args:
            top_participants (List[Dict]): List of participant dicts with 'name' and 'count' keys
            
        Returns:
            str: HTML for participant items
        """
        html = ''
        colors = [
            '#ec4899',  # pink
            '#f97316',  # orange
            '#eab308',  # yellow
            '#22c55e',  # green
            '#06b6d4',  # cyan
            '#3b82f6',  # blue
            '#8b5cf6',  # purple
            '#6366f1',  # indigo
            '#14b8a6',  # teal
            '#f43f5e',  # rose
        ]
        
        for idx, participant in enumerate(top_participants):
            color = colors[idx % len(colors)]
            name = participant.get('name', 'Unknown')
            count = participant.get('count', 0)
            
            # Truncate long names
            display_name = name[:35] + '...' if len(name) > 35 else name
            
            html += f'''<div style="padding: 12px; background: color-mix(in srgb, {color} 5%, var(--input-bg)); border-radius: 8px; border: 1px solid color-mix(in srgb, {color} 20%, var(--input-border));">
              <div style="font-size: 0.85em; color: var(--text); font-weight: 600; margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="{name}">#{idx+1} {display_name}</div>
              <div style="font-size: 0.9em; color: var(--meta);">{count:,} messages</div>
            </div>'''
        
        return html
    
    def _generate_chat_html(self, chat_name: str, chat_id: str, chat_type: str,
                           member_lists: Optional[Dict[str, str]] = None,
                           search_index_exporter: Optional[FragmentedExporter] = None) -> Tuple[str, int]:
        """
        Generate HTML for a chat section.
        
        Args:
            chat_name (str): The display name of the chat
            chat_id (str): The unique identifier of the chat
            chat_type (str): Type of chat ('group', 'meeting', or 'oneonone')
            member_lists (Optional[Dict[str, str]]): Member lists for display
            
        Returns:
            Tuple[str, int]: HTML content for the chat section and message count
        """
        safe_chat_name = urllib.parse.quote(str(chat_name or ''), safe='')
        html = f'<div id="{safe_chat_name}" class="chat-section">\n'
        
        # Add chat header
        html += f'  <h2 style="margin-top:0">{chat_name}</h2>\n'
        
        # Add member list for group and meeting chats
        if member_lists and chat_name in member_lists:
            formatted_members = format_member_list_for_display(member_lists[chat_name])
            if formatted_members:
                html += f'  <div class="chat-members" style="color:#555;font-size:0.95em;margin-bottom:10px;">Members: {formatted_members}</div>\n'
        
        # Check if this chat should be ignored
        if chat_name in IGNORED_CHATS:
            print(f'#   Chat "{chat_name}" is in ignore list - skipping message fetch')
            # Add ignore placeholder message
            ignored_html = f'<div class="clearfix"><div class="message theirs"><div class="meta">System • {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div><div class="text">This chat has been ignored due to configuration settings.<br><br>Chat: {chat_name}<br><br>To enable message fetching, remove this chat from the IGNORED_CHATS list in config.py.</div></div></div>'
            html += ignored_html
            html += '</div>\n'
            return html, 0
        
        # Fetch and process messages
        messages = self._fetch_chat_messages(chat_name, chat_id)
        
        # Sort messages chronologically (oldest first)
        messages.sort(key=lambda x: x.get('createdDateTime', ''))
        
        # Filter out system messages and empty messages
        filtered_messages = []
        for msg in messages:
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
            
            # Skip Unknown users and empty content
            if sender == 'Unknown':
                continue
                
            body = msg.get('body') or {}
            raw_content = body.get('content') or ''
            if not raw_content or not raw_content.strip():
                continue
                
            # Skip system event messages and other common empty content patterns
            # Also skip messages that only contain attachment tags without meaningful content
            raw_stripped = raw_content.strip()
            if ('<systemEventMessage/>' in raw_content or 
                raw_stripped in ['', '<p></p>', '<div></div>'] or
                (raw_stripped.startswith('<attachment id=') and raw_stripped.endswith('</attachment>') and len(raw_stripped) < 100) or
                re.match(r'^<attachment id="[^"]*"></attachment>$', raw_stripped)):
                continue
            
            # Apply date range filter
            if not is_message_in_date_range(msg, MESSAGE_DATE_FROM, MESSAGE_DATE_TO):
                continue
                
            filtered_messages.append(msg)
        
        # First pass: generate all message HTML and filter out empty ones
        messages_with_html = []
        skipped_count = 0
        for msg in filtered_messages:
            message_html = self._generate_message_html(msg, chat_type)
            if message_html.strip():  # Only keep messages with non-empty HTML
                # Use lastModifiedDateTime if available, otherwise fall back to createdDateTime
                timestamp = msg.get('lastModifiedDateTime') or msg.get('createdDateTime', '')
                date_label = self._get_date_label(timestamp)
                messages_with_html.append((date_label, message_html))

                # Add message to global search index for cross-chat filtering.
                if search_index_exporter:
                    from_data = msg.get('from') or {}
                    user_data = from_data.get('user') if isinstance(from_data, dict) else {}
                    sender = (user_data.get('displayName', 'Unknown') if isinstance(user_data, dict) else 'Unknown')
                    search_index_exporter.add_message_to_search_index(
                        sender,
                        message_html,
                        chat_name,
                        chat_id,
                        timestamp,
                    )
            else:
                skipped_count += 1

        # Second pass: add date separators only for messages that will actually be displayed
        last_date_label = None
        added_message_count = 0
        for date_label, message_html in messages_with_html:
            if date_label != last_date_label:
                html += f'<div class="date-separator"><span>{date_label}</span></div>'
                last_date_label = date_label
            html += message_html
            added_message_count += 1

        if skipped_count > 0:
            print(f'#   Skipped {skipped_count} empty messages in {chat_name}')
        html += '</div>\n'
        return html, added_message_count

    def _generate_channel_html(self, team_name: str, channel_name: str, messages: List[Dict],
                              search_index_exporter: Optional[FragmentedExporter] = None) -> Tuple[str, int]:
        """
        Generate HTML for a channel section.
        
        Args:
            team_name (str): The name of the team
            channel_name (str): The name of the channel
            messages (List[Dict]): List of message objects
            
        Returns:
            Tuple[str, int]: HTML content for the channel section and message count
        """
        safe_channel_id = urllib.parse.quote(f"{team_name or ''}|||{channel_name or ''}", safe='')
        html = f'<div id="{safe_channel_id}" class="chat-section">\n'
        html += f'  <h2 style="margin-top:0">{team_name} / {channel_name}</h2>\n'
        
        # Sort messages chronologically (oldest first)
        messages.sort(key=lambda x: x.get('createdDateTime', ''))
        
        # Check if this is an ignored channel - if so, skip filtering and use messages as-is
        if (team_name, channel_name) in IGNORED_CHANNELS:
            # For ignored channels, just use the messages directly (they should be ignore placeholder messages)
            filtered_messages = messages
        else:
            # Filter out system messages and empty messages for normal channels
            filtered_messages = []
            for msg in messages:
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
                
                # Skip Unknown users and empty content
                if sender == 'Unknown':
                    continue
                    
                body = msg.get('body') or {}
                raw_content = body.get('content') or ''
                if not raw_content or not raw_content.strip():
                    continue
                    
                # Skip system event messages and other common empty content patterns
                # Also skip messages that only contain attachment tags without meaningful content
                raw_stripped = raw_content.strip()
                if ('<systemEventMessage/>' in raw_content or 
                    raw_stripped in ['', '<p></p>', '<div></div>'] or
                    (raw_stripped.startswith('<attachment id=') and raw_stripped.endswith('</attachment>') and len(raw_stripped) < 100) or
                    re.match(r'^<attachment id="[^"]*"></attachment>$', raw_stripped)):
                    continue
                
                # Apply date range filter
                if not is_message_in_date_range(msg, MESSAGE_DATE_FROM, MESSAGE_DATE_TO):
                    continue
                    
                filtered_messages.append(msg)
        
        # First pass: generate all message HTML and filter out empty ones
        messages_with_html = []
        skipped_count = 0
        for msg in filtered_messages:
            # Use lastModifiedDateTime if available, otherwise fall back to createdDateTime
            timestamp = msg.get('lastModifiedDateTime') or msg.get('createdDateTime', '')
            message_html = self._generate_message_html(msg, 'channel')
            if message_html.strip():  # Only keep messages with non-empty HTML
                date_label = self._get_date_label(timestamp)
                messages_with_html.append((date_label, message_html))

                # Add channel message to global search index.
                if search_index_exporter:
                    from_data = msg.get('from') or {}
                    user_data = from_data.get('user') if isinstance(from_data, dict) else {}
                    sender = (user_data.get('displayName', 'Unknown') if isinstance(user_data, dict) else 'Unknown')
                    full_channel_name = f"{team_name} / {channel_name}"
                    combined_channel_id = f"{team_name}|||{channel_name}"
                    search_index_exporter.add_message_to_search_index(
                        sender,
                        message_html,
                        full_channel_name,
                        combined_channel_id,
                        timestamp,
                    )
            else:
                skipped_count += 1

        # Second pass: add date separators only for messages that will actually be displayed
        last_date_label = None
        added_message_count = 0
        for date_label, message_html in messages_with_html:
            if date_label != last_date_label:
                html += f'<div class="date-separator"><span>{date_label}</span></div>'
                last_date_label = date_label
            html += message_html
            added_message_count += 1
        
        if skipped_count > 0:
            print(f'#   Skipped {skipped_count} empty messages in {team_name}/{channel_name}')
        html += '</div>\n'
        return html, added_message_count
    
    def generate_html_export(self):
        """Generate the multi-file SPA export."""
        print('##  Generating multi-file SPA export')
        
        # Initialize fragmented exporter
        exporter = FragmentedExporter(self.output_folder, self.access_token, self.user_display_name)
        
        # Create shared assets (CSS, JS)
        exporter.create_assets()
        
        # Reset message count tracking
        self.chat_message_counts = {'oneonone': {}, 'group': {}, 'meeting': {}}
        self.channel_message_counts = {}

        # Build chat sections first to capture message counts
        print('##  Processing chat messages')
        
        # Process one-on-one chats
        for chat_name, chat_id in self.chats_one_on_one.items():
            chat_html, count = self._generate_chat_html(
                chat_name, chat_id, 'oneonone', search_index_exporter=exporter
            )
            self.chat_message_counts['oneonone'][chat_name] = count
            
            if count > 0:
                filename = exporter.create_chat_file(chat_id, chat_name, 'oneonone', chat_html)
                print(f'✅ Created chat file: {filename}')

        # Process group chats
        for chat_name, chat_id in self.chats_group.items():
            chat_html, count = self._generate_chat_html(
                chat_name, chat_id, 'group', self.group_full_member_lists, exporter
            )
            self.chat_message_counts['group'][chat_name] = count
            
            if count > 0:
                members_html = ''
                if chat_name in self.group_full_member_lists:
                    members_html = f'Members: {format_member_list_for_display(self.group_full_member_lists[chat_name])}'
                filename = exporter.create_chat_file(chat_id, chat_name, 'group', chat_html, members_html)
                print(f'✅ Created chat file: {filename}')

        # Process meeting chats
        for chat_name, chat_id in self.chats_meeting.items():
            chat_html, count = self._generate_chat_html(
                chat_name, chat_id, 'meeting', self.meeting_full_member_lists, exporter
            )
            self.chat_message_counts['meeting'][chat_name] = count
            
            if count > 0:
                members_html = ''
                if chat_name in self.meeting_full_member_lists:
                    members_html = f'Members: {format_member_list_for_display(self.meeting_full_member_lists[chat_name])}'
                filename = exporter.create_chat_file(chat_id, chat_name, 'meeting', chat_html, members_html)
                print(f'✅ Created chat file: {filename}')

        # Process channels
        for (team_name, channel_name), messages in self.channel_messages.items():
            channel_html, count = self._generate_channel_html(
                team_name, channel_name, messages, exporter
            )
            self.channel_message_counts[(team_name, channel_name)] = count
            
            if count > 0:
                # For channels, use team_name/channel_name as unique ID
                channel_id = f"{team_name}|||{channel_name}"
                full_channel_name = f"{team_name} / {channel_name}"
                filename = exporter.create_chat_file(channel_id, full_channel_name, 'channel', channel_html)
                print(f'✅ Created channel file: {filename}')
        
        # Generate sidebar HTML
        sidebar_html = exporter.generate_sidebar_sections(
            self.chats_one_on_one,
            self.chats_group,
            self.chats_meeting,
            self.channels_by_team,
            self.chat_message_counts,
            self.channel_message_counts,
            {**self.group_full_member_lists, **self.meeting_full_member_lists}
        )
        
        # Create index.html with sidebar
        exporter.create_index(sidebar_html)
        
        # Calculate statistics
        stats = self._calculate_statistics()
        stats['script_version'] = self.script_version
        
        # Create stats.html
        exporter.create_stats_page(stats)
        
        # Save search index
        exporter.save_search_index()
        
        print('✅ Multi-file SPA export completed successfully!')
        
        return {}  # Return empty dict as we're not using the traditional single-file approach
    
    def save_export(self, html_content: str = None):
        """Multi-file export - no single file to save."""
        # With the new multi-file approach, export is already saved during generate_html_export()
        print("✅ Multi-file SPA export completed successfully!")
    
    def run_export(self):
        """Run the complete export process."""
        print('### Microsoft Teams Chat Export Started')
        print(f'Script Version: {self.script_version}')
        
        # Validate access token before proceeding
        if not self._validate_access_token():
            sys.exit(1)
        
        try:
            # Step 1: Fetch all chats
            self.fetch_all_chats()
            
            # Step 2: Fetch teams and channels
            self.fetch_teams_and_channels()
            
            # Step 3: Generate multi-file SPA export
            self.generate_html_export()
            
            print('### Export completed successfully!')
            
        except Exception as e:
            print(f"❌ Export failed: {e}")
            sys.exit(1)


def main():
    """Main entry point for the script."""
    if not ACCESS_TOKEN or ACCESS_TOKEN == 'YOUR_ACCESS_TOKEN_HERE':
        print("❌ Error: Please configure your access token in config.py")
        print("   Set ACCESS_TOKEN to your valid Microsoft Graph API token")
        sys.exit(1)
    
    exporter = TeamsExporter()
    exporter.run_export()


if __name__ == "__main__":
    main()
