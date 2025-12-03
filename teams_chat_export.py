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
Version: v0.1.1
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
    GRAPH_API_BASE_URL, ITEMS_PER_PAGE, OUTPUT_HTML_FILE, NAVIGATION_BUTTONS_HTML,
    CHAT_TYPE_ONE_ON_ONE, CHAT_TYPE_GROUP, CHAT_TYPE_MEETING, MESSAGE_TYPE_USER,
    MAX_MEMBERS_IN_CHAT_NAME, IGNORED_CHANNELS, IGNORED_CHATS, get_api_headers
)
from html_template import html_content
from teams_utils import (
    chat_has_messages, sort_chats_by_name, process_message_content,
    format_timestamp, create_member_list_display, format_member_list_for_display,
    validate_message_content
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
        
        # Setup output directory
        self._setup_output_directory()
    
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
    
    def _make_paginated_request(self, url: str, progress_callback=None, context="") -> List[Dict]:
        """
        Make paginated requests to Microsoft Graph API.
        
        Args:
            url (str): The initial API URL
            progress_callback (callable): Optional callback for progress updates
            context (str): Context string for progress display
            
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
                return response.json().get('value', [])
        except Exception as e:
            print(f"Error fetching members for chat {chat_id}: {e}")
        return []
    
    def _process_one_on_one_chat(self, chat: Dict):
        """Process a one-on-one chat."""
        chat_id = chat.get('id')
        chat_name = chat.get('topic', '')
        
        if not chat_has_messages(chat_id, self.headers):
            return
        
        # If chat name is empty, use the other member's display name
        if not chat_name:
            members = self._get_chat_members(chat_id)
            other_member = next(
                (m for m in members if m.get('displayName') != self.user_display_name), 
                None
            )
            display_name = other_member.get('displayName', 'Unknown') if other_member else 'Unknown'
            chat_name = f"Chat with {display_name}"
        
        print(f'#   Processing: {chat_name}')
        self.chats_one_on_one[chat_name] = chat_id
    
    def _process_group_chat(self, chat: Dict):
        """Process a group chat."""
        chat_id = chat.get('id')
        chat_name = chat.get('topic', '')
        
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
        chat_name = chat.get('topic', '')
        
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
        
        for team in teams:
            team_id = team.get('id')
            team_name = team.get('displayName', f"Team {team_id}")
            print(f'#   Processing team: {team_name}')
            
            # Get channels for this team
            channels_url = f'{GRAPH_API_BASE_URL}/teams/{team_id}/channels'
            channels = self._make_paginated_request(channels_url)
            
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
        messages = self._make_paginated_request(url, chat_progress_callback, f"chat_{chat_name}")
        
        print(f'\r#   Fetching messages for chat: {chat_name}...({len(messages)} messages)')
        return messages
    
    def _generate_sidebar_section(self, section_id: str, title: str, chats: Dict[str, str], 
                                 member_lists: Optional[Dict[str, str]] = None) -> str:
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
        html = f'''
  <div class="sidebar-section-header" onclick="toggleSection('{section_id}')">{title}</div>
  <div id="{section_id}" class="sidebar-section-content" style="display:none;">
'''
        
        for chat_name, chat_id in chats.items():
            safe_chat_name = urllib.parse.quote(chat_name, safe='')
            tooltip = ""
            
            if member_lists and chat_name in member_lists:
                tooltip = f'title="{member_lists[chat_name]}"'
            
            html += f'''    <a href="#" onclick="showChat('{safe_chat_name}')" data-chat-name="{chat_name}" {tooltip}>{chat_name}</a>
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
                raw_content, message_id, self.access_token, self.image_folder
            )
            
            # Skip messages with no content after processing
            # Remove common empty HTML patterns and check if there's meaningful content
            content_check = re.sub(r'<[^>]+>', '', clean_content or '').strip()
            if not clean_content or not clean_content.strip() or not content_check:
                return ""
            
            # Format timestamp
            timestamp = msg.get('lastModifiedDateTime', msg.get('createdDateTime', ''))
            formatted_timestamp = format_timestamp(timestamp)
            
            # Determine CSS class based on sender
            css_class = 'mine' if sender == self.user_display_name else 'theirs'
            
            return f'<div class="clearfix"><div class="message {css_class}"><div class="meta">{sender} • {formatted_timestamp}</div><div class="text">{clean_content}</div></div></div>'
        except Exception as e:
            # Silently handle errors and return empty string to avoid console spam
            # Uncomment the line below if you want to see the error details for debugging
            # print(f"Error processing message: {e}")
            return ""
    
    def _generate_chat_html(self, chat_name: str, chat_id: str, chat_type: str, 
                           member_lists: Optional[Dict[str, str]] = None) -> str:
        """
        Generate HTML for a chat section.
        
        Args:
            chat_name (str): The display name of the chat
            chat_id (str): The unique identifier of the chat
            chat_type (str): Type of chat ('group', 'meeting', or 'oneonone')
            member_lists (Optional[Dict[str, str]]): Member lists for display
            
        Returns:
            str: HTML content for the chat section
        """
        safe_chat_name = urllib.parse.quote(chat_name, safe='')
        html = f'<div id="{safe_chat_name}" class="chat-section">\n'
        
        # Add chat header
        html += f'  <h2 style="margin-top:0">{chat_name}</h2>\n'
        
        # Add member list for group and meeting chats
        if member_lists and chat_name in member_lists:
            formatted_members = format_member_list_for_display(member_lists[chat_name])
            if formatted_members:
                html += f'  <div style="color:#555;font-size:0.95em;margin-bottom:10px;">Members: {formatted_members}</div>\n'
        
        # Check if this chat should be ignored
        if chat_name in IGNORED_CHATS:
            print(f'#   Chat "{chat_name}" is in ignore list - skipping message fetch')
            # Add ignore placeholder message
            ignored_html = f'<div class="clearfix"><div class="message theirs"><div class="meta">System • {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div><div class="text">This chat has been ignored due to configuration settings.<br><br>Chat: {chat_name}<br><br>To enable message fetching, remove this chat from the IGNORED_CHATS list in config.py.</div></div></div>'
            html += ignored_html
            html += '</div>\n'
            return html
        
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
                
            raw_content = msg.get('body', {}).get('content', '')
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
                
            filtered_messages.append(msg)
        
        # Generate HTML for each valid message
        for msg in filtered_messages:
            message_html = self._generate_message_html(msg, 'chat')
            if message_html.strip():  # Only add non-empty HTML
                html += message_html
        
        html += '</div>\n'
        return html
    
    def _generate_channel_html(self, team_name: str, channel_name: str, messages: List[Dict]) -> str:
        """
        Generate HTML for a channel section.
        
        Args:
            team_name (str): The name of the team
            channel_name (str): The name of the channel
            messages (List[Dict]): List of message objects
            
        Returns:
            str: HTML content for the channel section
        """
        safe_channel_id = urllib.parse.quote(f"{team_name}|||{channel_name}", safe='')
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
                    
                raw_content = msg.get('body', {}).get('content', '')
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
                    
                filtered_messages.append(msg)
        
        # Generate HTML for each valid message
        for msg in filtered_messages:
            message_html = self._generate_message_html(msg, 'channel')
            if message_html.strip():  # Only add non-empty HTML
                html += message_html
        
        html += '</div>\n'
        return html
    
    def generate_html_export(self):
        """Generate the complete HTML export."""
        print('##  Generating HTML export')
        
        # Start with the base template
        export_html = html_content
        
        # Add sidebar sections
        export_html += '<div class="sidebar-section">\n'
        
        # One-on-one chats
        export_html += self._generate_sidebar_section(
            'oneonone-section', '▶ One on One Chats', self.chats_one_on_one
        )
        
        # Group chats
        export_html += self._generate_sidebar_section(
            'group-section', '▶ Group Chats', self.chats_group, self.group_full_member_lists
        )
        
        # Meeting chats
        export_html += self._generate_sidebar_section(
            'meeting-section', '▶ Meeting Chats', self.chats_meeting
        )
        
        # Channel chats
        export_html += '''  <div class="sidebar-section-header" onclick="toggleSection('channel-section')">▶ Channel Chats</div>
  <div id="channel-section" class="sidebar-section-content" style="display:none;">
'''
        
        # Add team sections
        for team_name, channels in self.channels_by_team.items():
            safe_team_id = urllib.parse.quote(team_name, safe='')
            export_html += f'    <div class="sidebar-section-header" onclick="toggleSection(\'team-{safe_team_id}\')">▶&nbsp;&nbsp;&nbsp;&nbsp;{team_name}</div>\n'
            export_html += f'    <div id="team-{safe_team_id}" class="sidebar-section-content" style="display:none;">\n'
            
            for channel_name, team_id, channel_id in channels:
                safe_channel_id = urllib.parse.quote(f"{team_name}|||{channel_name}", safe='')
                export_html += f'      <a href="#" onclick="showChat(\'{safe_channel_id}\')" data-chat-name="{channel_name}">{channel_name}</a>\n'
            
            export_html += '    </div>\n'
        
        export_html += '  </div>\n'
        export_html += '</div>\n'
        export_html += '</div>\n'
        export_html += '</div><div class="content">\n'
        
        # Add cover page
        export_html += '''  <div id="cover-page" style="display: block; text-align: center; padding: 60px 20px 40px 20px; color: #23272e;">
    <h1 style="font-size:2.5em;margin-bottom:0.2em;">Teams Chat Export</h1>
    <p style="font-size:1.2em;max-width:600px;margin:0 auto 1.5em auto;">
      Welcome!<br>
      This file contains all your exported Microsoft Teams chats, meetings, and channel messages.<br>
      <span style="color:#888;">Exported on <span id="cover-export-date"><!--EXPORT_DATE--></span></span>
    </p>
    <ul style="text-align:left;display:inline-block;margin-bottom:1.5em;">
      <li>Browse all your chats and channels using the sidebar.</li>
      <li>Click on a chat or channel to view its messages.</li>
      <li>Use the search box to filter messages across all chats.</li>
      <li>Click images to view them in a lightbox.</li>
    </ul>
    <div style="margin-top:2em;font-size:0.9em;color:#aaa;">Powered by Teams Chat Export Script <!--SCRIPT_VERSION--></div>
  </div>
'''
        
        # Add chat content
        print('##  Processing chat messages')
        
        # Process one-on-one chats
        for chat_name, chat_id in self.chats_one_on_one.items():
            export_html += self._generate_chat_html(chat_name, chat_id, 'oneonone')
        
        # Process group chats
        for chat_name, chat_id in self.chats_group.items():
            export_html += self._generate_chat_html(chat_name, chat_id, 'group', self.group_full_member_lists)
        
        # Process meeting chats
        for chat_name, chat_id in self.chats_meeting.items():
            export_html += self._generate_chat_html(chat_name, chat_id, 'meeting', self.meeting_full_member_lists)
        
        # Process channel messages
        for (team_name, channel_name), messages in self.channel_messages.items():
            export_html += self._generate_channel_html(team_name, channel_name, messages)
        
        # Replace placeholders
        export_date = datetime.now().strftime("%Y-%m-%d at %H:%M:%S")
        export_html = export_html.replace("<!--EXPORT_DATE-->", f"{export_date} with {self.script_version}")
        export_html = export_html.replace("<!--SCRIPT_VERSION-->", self.script_version)
        
        # Add navigation buttons and close tags
        export_html += NAVIGATION_BUTTONS_HTML
        export_html += '</div></body></html>'
        
        return export_html
    
    def save_export(self, html_content: str):
        """Save the HTML export to file."""
        output_file = os.path.join(self.output_folder, OUTPUT_HTML_FILE)
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"✅ Successfully exported to '{output_file}'")
        except Exception as e:
            print(f"❌ Error saving export: {e}")
            sys.exit(1)
    
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
            
            # Step 3: Generate HTML export
            html_export = self.generate_html_export()
            
            # Step 4: Save export
            self.save_export(html_export)
            
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
