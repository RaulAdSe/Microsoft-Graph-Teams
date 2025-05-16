#!/usr/bin/env python3
# Servitec OneNote Notebook Extraction
# This script extracts OneNote notebooks from SharePoint document libraries associated with Teams

import requests
import json
import time
import os
import sys
from dotenv import load_dotenv
from pprint import pprint

# Load environment variables
load_dotenv()

# Access token from .env
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
# NEVER hardcode tokens in source code

if not ACCESS_TOKEN:
    print("Error: ACCESS_TOKEN not found in .env file")
    sys.exit(1)

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Accept": "application/json"
}

# Microsoft Graph API base URL
graph_base_url = "https://graph.microsoft.com/v1.0"

# Debug level (0-3): 0=minimal, 1=normal, 2=detailed, 3=verbose with raw data
DEBUG_LEVEL = 2

# Set this to test a specific team instead of all teams
# e.g., "1d092b0f-f8eb-459c-b391-f4487e66680f"
TEST_TEAM_ID = ""  # Leave empty to process all teams

def debug_print(level, message, data=None):
    """Print debug messages based on the current debug level"""
    if level <= DEBUG_LEVEL:
        print(message)
        if data is not None and DEBUG_LEVEL >= 3:
            if isinstance(data, dict) or isinstance(data, list):
                pprint(data)
            else:
                print(data)

def rate_limited_request(url, method="GET", retry_after=2):
    """Make a request with rate limiting to handle API throttling"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            debug_print(1, f"Making request to: {url} (attempt {attempt+1}/{max_retries})")
            
            if method == "GET":
                response = requests.get(url, headers=headers)
            else:
                debug_print(1, f"Unsupported method: {method}")
                return None
                
            if response.status_code == 200:
                data = response.json()
                debug_print(2, f"Request successful, status code: {response.status_code}")
                debug_print(3, "Raw response data:", data)
                return data
            elif response.status_code == 429:  # Too Many Requests
                # Get retry-after header or use default
                retry_after_val = int(response.headers.get('Retry-After', retry_after))
                debug_print(0, f"⚠️ Rate limited! Waiting for {retry_after_val} seconds before retrying...")
                time.sleep(retry_after_val)
                # Increase backoff for next potential retry
                retry_after = min(retry_after * 2, 60)  # Double the wait time, max 60 seconds
            else:
                debug_print(1, f"Error: {response.status_code} - {response.text}")
                if "forbidden" in response.text.lower() or "unauthorized" in response.text.lower():
                    debug_print(0, f"⚠️ Permission issue! Check that your token has the required permissions.")
                return None
                
        except Exception as e:
            debug_print(1, f"Exception making request: {str(e)}")
            time.sleep(retry_after)
    
    debug_print(0, f"❌ Failed after {max_retries} attempts to access: {url}")
    return None

def make_request(url, method="GET"):
    """Make a request to the Microsoft Graph API"""
    return rate_limited_request(url, method)

def get_team_details(team_id):
    """Get detailed information about a specific team"""
    debug_print(1, f"Getting details for team ID: {team_id}")
    team_url = f"{graph_base_url}/teams/{team_id}"
    
    response = make_request(team_url)
    if response:
        debug_print(2, f"Team details retrieved successfully: {response.get('displayName', 'Unknown')}")
        return response
    
    debug_print(1, "Could not retrieve team details")
    return None

def get_all_teams():
    """Get all teams the user is a member of"""
    debug_print(1, "Fetching all teams...")
    teams_url = f"{graph_base_url}/me/joinedTeams"
    
    response = make_request(teams_url)
    if response and "value" in response:
        teams = response["value"]
        debug_print(1, f"Found {len(teams)} teams")
        for team in teams:
            debug_print(2, f"  - Team: {team.get('displayName')} (ID: {team.get('id')})")
        return teams
    
    debug_print(1, "No teams found or error occurred")
    return []

def get_sharepoint_site_for_team(team_id):
    """Get the SharePoint site associated with a team"""
    debug_print(1, f"Getting SharePoint site for team ID: {team_id}")
    site_url = f"{graph_base_url}/groups/{team_id}/sites/root"
    
    response = make_request(site_url)
    if response and "id" in response:
        site_id = response.get("id")
        site_name = response.get("displayName", "Unknown")
        web_url = response.get("webUrl", "")
        debug_print(1, f"Found SharePoint site: {site_name} (ID: {site_id})")
        debug_print(1, f"SharePoint URL: {web_url}")
        debug_print(2, "SharePoint site details:", response)
        return response
    
    debug_print(1, "SharePoint site not found or permission error")
    if "403" in str(response):
        debug_print(1, "⚠️ Permission issue: Your token likely lacks Sites.Read.All permissions")
    
    return None

def get_document_library(site_id):
    """Get the document library (usually 'Documents') for a SharePoint site"""
    debug_print(1, f"Getting document library for site ID: {site_id}")
    drive_url = f"{graph_base_url}/sites/{site_id}/drives"
    
    response = make_request(drive_url)
    if response and "value" in response:
        drives = response["value"]
        debug_print(2, f"Found {len(drives)} drives in the site")
        
        # Look for the document library, typically called "Documents"
        for drive in drives:
            drive_name = drive.get("name", "")
            drive_id = drive.get("id", "")
            drive_type = drive.get("driveType", "")
            debug_print(2, f"  - Drive: {drive_name} (ID: {drive_id}, Type: {drive_type})")
            
            if "document" in drive_name.lower() or drive_type == "documentLibrary":
                debug_print(1, f"Found document library: {drive_name} (ID: {drive_id})")
                return drive
        
        # If we didn't find a "Documents" library, just return the first drive
        if drives:
            debug_print(1, f"Using first available drive: {drives[0].get('name')} (ID: {drives[0].get('id')})")
            return drives[0]
    
    debug_print(1, "No document library found or error occurred")
    return None

def explore_drive_structure(drive_id, folder_path="/", folder_id=None, depth=0):
    """Recursively explore the drive structure (for debugging)"""
    if depth > 3:  # Limit recursion depth
        return
    
    indent = "  " * depth
    if folder_id:
        debug_print(1, f"{indent}Exploring folder: {folder_path} (ID: {folder_id})")
        folder_url = f"{graph_base_url}/drives/{drive_id}/items/{folder_id}/children"
    else:
        debug_print(1, f"{indent}Exploring drive root: {drive_id}")
        folder_url = f"{graph_base_url}/drives/{drive_id}/root/children"
    
    response = make_request(folder_url)
    if response and "value" in response:
        items = response["value"]
        debug_print(1, f"{indent}Found {len(items)} items in {folder_path}")
        
        for item in items:
            name = item.get("name", "")
            item_id = item.get("id", "")
            item_type = "Folder" if item.get("folder") else "File"
            web_url = item.get("webUrl", "")
            
            # Print details of the current item
            debug_print(1, f"{indent}- {item_type}: {name} (ID: {item_id})")
            debug_print(2, f"{indent}  URL: {web_url}")
            
            # If it's a folder and we're not at max depth, recurse
            if item.get("folder") and depth < 2:
                new_path = f"{folder_path}/{name}"
                explore_drive_structure(drive_id, new_path, item_id, depth + 1)
                
            # Special handling for OneNote files
            if not item.get("folder") and (".one" in name.lower() or "onenote" in name.lower()):
                debug_print(1, f"{indent}  📓 Found OneNote file: {name}")
                
                # Get additional details if needed
                if DEBUG_LEVEL >= 2:
                    item_details_url = f"{graph_base_url}/drives/{drive_id}/items/{item_id}"
                    item_details = make_request(item_details_url)
                    if item_details:
                        debug_print(2, f"{indent}  OneNote details:", item_details)

def get_site_library_folder(drive_id):
    """Get the 'Site Library' folder within the Documents library"""
    debug_print(1, f"Looking for 'Site Library' folder in drive ID: {drive_id}")
    root_items_url = f"{graph_base_url}/drives/{drive_id}/root/children"
    
    response = make_request(root_items_url)
    if response and "value" in response:
        items = response["value"]
        debug_print(2, f"Found {len(items)} items at root level")
        
        # Print all folders at the root level for analysis
        debug_print(1, "All folders at root level:")
        for item in items:
            if item.get("folder"):
                name = item.get("name", "")
                item_id = item.get("id", "")
                folder_size = item.get("folder", {}).get("childCount", 0)
                debug_print(1, f"  - Folder: {name} (ID: {item_id}, Items: {folder_size})")
        
        # Look for a folder named "Site Library" or similar
        for item in items:
            name = item.get("name", "")
            item_id = item.get("id", "")
            
            if item.get("folder") and ("site library" in name.lower() or "sitelibrary" in name.lower()):
                debug_print(1, f"Found Site Library folder: {name} (ID: {item_id})")
                return item
        
        # If we can't find a specific "Site Library" folder, check for other common names
        for item in items:
            name = item.get("name", "")
            item_id = item.get("id", "")
            
            if item.get("folder") and ("channels" in name.lower() or "team" in name.lower()):
                debug_print(1, f"Found potential site library folder: {name} (ID: {item_id})")
                return item
                
        # If we still can't find a relevant folder, just return any folder that might contain channels
        for item in items:
            if item.get("folder"):
                debug_print(1, f"Using folder as potential site library: {item.get('name')} (ID: {item.get('id')})")
                return item
    
    debug_print(1, "No 'Site Library' folder found")
    return None

def get_channel_folders(drive_id, parent_folder_id=None):
    """Get folders within the Site Library that might correspond to Teams channels"""
    if parent_folder_id:
        debug_print(1, f"Getting channel folders from parent folder ID: {parent_folder_id}")
        folder_items_url = f"{graph_base_url}/drives/{drive_id}/items/{parent_folder_id}/children"
    else:
        debug_print(1, f"Getting channel folders from drive root: {drive_id}")
        folder_items_url = f"{graph_base_url}/drives/{drive_id}/root/children"
    
    response = make_request(folder_items_url)
    if response and "value" in response:
        items = response["value"]
        debug_print(2, f"Found {len(items)} items in parent folder")
        
        # Print all items for debugging
        debug_print(2, "All items in this folder:")
        for item in items:
            item_type = "Folder" if item.get("folder") else "File"
            debug_print(2, f"  - {item_type}: {item.get('name')} (ID: {item.get('id')})")
        
        # Look for a "General" folder (default channel) first
        general_folder = None
        channel_folders = []
        
        for item in items:
            name = item.get("name", "")
            item_id = item.get("id", "")
            if item.get("folder") and "general" in name.lower():
                general_folder = item
                debug_print(1, f"Found General channel folder: {name} (ID: {item_id})")
            elif item.get("folder"):
                channel_folders.append(item)
                debug_print(1, f"Found potential channel folder: {name} (ID: {item_id})")
        
        # Combine General and other channel folders
        if general_folder:
            channel_folders.insert(0, general_folder)
        
        debug_print(1, f"Total channel folders found: {len(channel_folders)}")
        return channel_folders
    
    debug_print(1, "No folders found or error occurred")
    return []

def find_onenote_files(drive_id, folder_id):
    """Find OneNote files within a folder"""
    debug_print(1, f"Looking for OneNote files in folder ID: {folder_id}")
    folder_items_url = f"{graph_base_url}/drives/{drive_id}/items/{folder_id}/children"
    
    response = make_request(folder_items_url)
    if response and "value" in response:
        items = response["value"]
        debug_print(2, f"Found {len(items)} items in folder")
        
        # Print all items in the folder for debugging
        if DEBUG_LEVEL >= 2:
            debug_print(2, "All items in this folder:")
            for item in items:
                item_type = "Folder" if item.get("folder") else "File"
                mime_type = item.get("file", {}).get("mimeType", "N/A") if not item.get("folder") else "N/A"
                debug_print(2, f"  - {item_type}: {item.get('name')} (MIME: {mime_type})")
        
        onenote_files = []
        
        for item in items:
            name = item.get("name", "")
            item_id = item.get("id", "")
            file_type = item.get("file", {}).get("mimeType", "")
            
            if ".one" in name.lower() or "onenote" in file_type.lower():
                debug_print(1, f"Found OneNote file: {name} (ID: {item_id})")
                onenote_files.append(item)
                
                # Try to get the notebook ID from the file properties
                item_details_url = f"{graph_base_url}/drives/{drive_id}/items/{item_id}"
                item_details = make_request(item_details_url)
                if item_details:
                    web_url = item_details.get("webUrl", "")
                    debug_print(1, f"OneNote web URL: {web_url}")
                    debug_print(2, "OneNote file details:", item_details)
                    # Try to extract the notebook ID from the URL if possible
                    # This will be used later to get notebook sections via OneNote API
                    
        debug_print(1, f"Total OneNote files found: {len(onenote_files)}")
        return onenote_files
    
    debug_print(1, "No files found or error occurred")
    return []

def get_notebooks_from_onenote_api(site_id):
    """Get all OneNote notebooks in a SharePoint site using OneNote API"""
    debug_print(1, f"Getting notebooks in SharePoint site ID: {site_id} via OneNote API")
    notebooks_url = f"{graph_base_url}/sites/{site_id}/onenote/notebooks"
    
    response = make_request(notebooks_url)
    if response and "value" in response:
        notebooks = response["value"]
        debug_print(1, f"Found {len(notebooks)} notebooks in SharePoint site via OneNote API")
        for notebook in notebooks:
            debug_print(2, f"  - Notebook: {notebook.get('displayName')} (ID: {notebook.get('id')})")
        return notebooks
    
    debug_print(1, "No notebooks found or error occurred")
    if "403" in str(response) or "401" in str(response):
        debug_print(1, "⚠️ Permission issue: Your token lacks Sites.Read.All and/or Notes.Read.All permissions")
    
    return []

def get_sections_for_notebook(notebook_id, site_id=None):
    """Get sections for a specific notebook using OneNote API"""
    debug_print(1, f"Getting sections for notebook ID: {notebook_id}")
    
    # Try different endpoints to get sections
    sections = None
    
    # Try SharePoint site path if we have site_id
    if site_id:
        debug_print(1, f"Trying site endpoint for sections...")
        url = f"{graph_base_url}/sites/{site_id}/onenote/notebooks/{notebook_id}/sections"
        response = make_request(url)
        
        if response and "value" in response:
            sections = response["value"]
            debug_print(1, f"Found {len(sections)} sections via site endpoint")
            for section in sections:
                debug_print(2, f"  - Section: {section.get('displayName')} (ID: {section.get('id')})")
            return sections
        
        debug_print(1, "Site endpoint failed")
    
    # Try personal endpoint as fallback
    if not sections:
        debug_print(1, f"Trying personal endpoint for sections...")
        url = f"{graph_base_url}/me/onenote/notebooks/{notebook_id}/sections"
        response = make_request(url)
        
        if response and "value" in response:
            sections = response["value"]
            debug_print(1, f"Found {len(sections)} sections via personal endpoint")
            for section in sections:
                debug_print(2, f"  - Section: {section.get('displayName')} (ID: {section.get('id')})")
            return sections
        
        debug_print(1, "Personal endpoint failed")
    
    debug_print(1, "No sections found for this notebook")
    return []

def extract_notebook_id_from_weburl(web_url):
    """Attempt to extract notebook ID from OneNote web URL"""
    # This is a fallback method - the OneNote API is more reliable when available
    debug_print(2, f"Attempting to extract notebook ID from URL: {web_url}")
    if "notebooks/" in web_url:
        parts = web_url.split("notebooks/")
        if len(parts) > 1:
            id_part = parts[1].split("/")[0]
            debug_print(2, f"Extracted notebook ID: {id_part}")
            return id_part
    debug_print(2, "Could not extract notebook ID from URL")
    return None

def get_team_channels(team_id):
    """Get all channels for a specific team"""
    debug_print(1, f"Getting channels for team ID: {team_id}")
    channels_url = f"{graph_base_url}/teams/{team_id}/channels"
    
    response = make_request(channels_url)
    if response and "value" in response:
        channels = response["value"]
        debug_print(2, f"Found {len(channels)} channels in team")
        
        # Create mappings for matching
        channel_map = {}
        channel_name_map = {}
        
        for channel in channels:
            channel_id = channel.get("id", "")
            channel_name = channel.get("displayName", "")
            channel_map[channel_id] = channel
            channel_name_map[channel_name.lower()] = channel
            debug_print(2, f"  - Channel: {channel_name} (ID: {channel_id})")
        
        return channels, channel_map, channel_name_map
    
    debug_print(1, "No channels found or error occurred")
    return [], {}, {}

def match_notebook_to_channel(notebook_name, channels, team_name):
    """Match notebook name to a channel based on name similarity"""
    debug_print(0, f"Trying to match notebook '{notebook_name}' to a channel in team '{team_name}'")
    
    channel_name = "Unknown Channel"
    channel_id = "unknown"
    simplified_notebook_name = notebook_name.lower()
    
    # Handle case where notebook name contains team name
    team_name_lower = team_name.lower()
    team_name_parts = team_name_lower.split()
    
    # Remove common prefixes and team name for better matching
    clean_notebook_name = simplified_notebook_name
    
    # First list all prefixes we want to remove
    prefixes = [
        "bloc de notas de ", 
        "notas_ ", 
        "notas de ", 
        "notebook ", 
        "notes - ", 
        "notes_ ", 
        "bloc de notes de ",
        "cuaderno de ",
        "onenote - "
    ]
    
    # Remove prefixes
    for prefix in prefixes:
        if clean_notebook_name.startswith(prefix):
            clean_notebook_name = clean_notebook_name[len(prefix):]
    
    # Check if the remaining notebook name is basically the team name
    # If so, always default to the "General" channel if it exists
    if clean_notebook_name == team_name_lower or clean_notebook_name.replace(" ", "") == team_name_lower.replace(" ", ""):
        debug_print(0, f"  📌 Notebook name matches team name after prefix removal: '{clean_notebook_name}' ≈ '{team_name_lower}'")
        general_channel = next((c for c in channels if c.get("displayName", "").lower() == "general"), None)
        if general_channel:
            channel_name = general_channel.get("displayName")
            channel_id = general_channel.get("id")
            debug_print(0, f"  ✅ Defaulting to General channel: {channel_name} ({channel_id})")
            return channel_name, channel_id
    
    debug_print(0, f"  📝 Cleaned notebook name for matching: '{clean_notebook_name}'")
    
    # Check if notebook name has multiple parts that could match with channels
    notebook_parts = clean_notebook_name.split()
    
    # List all channels for debugging
    debug_print(0, f"  Available channels in team '{team_name}':")
    for i, channel in enumerate(channels):
        ch_name = channel.get("displayName", "")
        ch_id = channel.get("id", "")
        debug_print(0, f"    {i+1}. {ch_name} (ID: {ch_id})")
    
    # Try to match with channel names - multiple matching strategies
    
    # STRATEGY 1: Try exact match first
    for channel in channels:
        ch_name = channel.get("displayName", "").lower()
        if not ch_name:  # Skip channels with no name
            continue
            
        # Check for exact match
        if ch_name == clean_notebook_name:
            channel_name = channel.get("displayName")
            channel_id = channel.get("id")
            debug_print(0, f"  ✅ Found exact channel match: {channel_name} ({channel_id})")
            return channel_name, channel_id
    
    # STRATEGY 2: Check for full containment
    for channel in channels:
        ch_name = channel.get("displayName", "").lower()
        if not ch_name:
            continue
            
        # Check if channel name contains notebook name or vice versa
        if ch_name in clean_notebook_name or clean_notebook_name in ch_name:
            channel_name = channel.get("displayName")
            channel_id = channel.get("id")
            debug_print(0, f"  ✅ Found containment match: {channel_name} ({channel_id})")
            return channel_name, channel_id
            
    # STRATEGY 3: Token matching with scoring
    best_match = None
    best_match_score = 0
    
    for channel in channels:
        ch_name = channel.get("displayName", "").lower()
        if not ch_name:
            continue
        
        # Calculate similarity score based on shared words
        ch_tokens = set(ch_name.split())
        nb_tokens = set(clean_notebook_name.split())
        
        # Give more weight to longer matches
        common_tokens = ch_tokens.intersection(nb_tokens)
        score = len(common_tokens)
        
        # Add bonus for longer matching words
        for token in common_tokens:
            if len(token) > 3:  # Bonus for words longer than 3 chars
                score += 1
        
        if score > best_match_score:
            best_match_score = score
            best_match = channel
    
    # Use best partial match if found
    if best_match and best_match_score > 0:
        channel_name = best_match.get("displayName")
        channel_id = best_match.get("id")
        debug_print(0, f"  ✅ Found partial match by token scoring: {channel_name} ({channel_id}) - score: {best_match_score}")
        return channel_name, channel_id
    
    # STRATEGY 4: Initial letters matching (for acronyms)
    for channel in channels:
        ch_name = channel.get("displayName", "").lower()
        if not ch_name:
            continue
            
        # Get initials from channel name
        ch_initials = ''.join([word[0] for word in ch_name.split() if word])
        
        # Check if notebook name contains these initials or vice versa
        if (len(ch_initials) > 1 and ch_initials in clean_notebook_name) or \
           (len(clean_notebook_name) > 0 and clean_notebook_name[0] == ch_name[0]):
            channel_name = channel.get("displayName")
            channel_id = channel.get("id")
            debug_print(0, f"  ✅ Found match by initials: {channel_name} ({channel_id})")
            return channel_name, channel_id
    
    # STRATEGY 5: If we have a General channel, use it as fallback
    general_channel = next((c for c in channels if c.get("displayName", "").lower() == "general"), None)
    if general_channel:
        channel_name = general_channel.get("displayName")
        channel_id = general_channel.get("id")
        debug_print(0, f"  ⚠️ No specific match found, using General channel: {channel_name} ({channel_id})")
        return channel_name, channel_id
    
    # STRATEGY 6: Last resort - just use the first channel if we have one
    if channels:
        channel_name = channels[0].get("displayName")
        channel_id = channels[0].get("id")
        debug_print(0, f"  ⚠️ No match found, using first available channel: {channel_name} ({channel_id})")
        return channel_name, channel_id
        
    debug_print(0, f"  ❌ NO CHANNEL FOUND for notebook '{notebook_name}' in team '{team_name}'")
    return channel_name, channel_id

def test_single_team(team_id):
    """Test the SharePoint structure for a single team"""
    debug_print(0, f"\n{'='*80}")
    debug_print(0, f"TESTING SHAREPOINT STRUCTURE FOR TEAM: {team_id}")
    debug_print(0, f"{'='*80}")
    
    # Get team details
    team_details = get_team_details(team_id)
    team_name = team_details.get("displayName", "Unknown Team") if team_details else "Unknown Team"
    debug_print(0, f"Team name: {team_name}")
    
    # Get SharePoint site
    site = get_sharepoint_site_for_team(team_id)
    if not site:
        debug_print(0, "Cannot access SharePoint site for this team")
        return
    
    site_id = site.get("id")
    site_name = site.get("displayName", "Unknown")
    debug_print(0, f"SharePoint site: {site_name} (ID: {site_id})")
    
    # Get document library
    document_library = get_document_library(site_id)
    if not document_library:
        debug_print(0, "Cannot find document library for this site")
        return
    
    drive_id = document_library.get("id")
    drive_name = document_library.get("name", "Unknown")
    debug_print(0, f"Document library: {drive_name} (ID: {drive_id})")
    
    # Explore the full drive structure for debugging
    debug_print(0, "\n== Exploring Drive Structure ==")
    explore_drive_structure(drive_id)
    
    # Find the Site Library folder
    debug_print(0, "\n== Looking for Site Library Folder ==")
    site_library = get_site_library_folder(drive_id)
    
    # Get the OneNote notebooks from API
    debug_print(0, "\n== Getting Notebooks from OneNote API ==")
    onenote_notebooks = get_notebooks_from_onenote_api(site_id)
    
    # Test channel folder detection
    debug_print(0, "\n== Testing Channel Folder Detection ==")
    if site_library:
        site_library_id = site_library.get("id")
        site_library_name = site_library.get("name")
        debug_print(0, f"Using '{site_library_name}' as the location for channel folders")
        channel_folders = get_channel_folders(drive_id, site_library_id)
    else:
        debug_print(0, "No Site Library folder found, looking for channels at root level")
        channel_folders = get_channel_folders(drive_id)
    
    # Test OneNote file detection in the first channel folder
    if channel_folders:
        first_folder = channel_folders[0]
        debug_print(0, "\n== Testing OneNote File Detection ==")
        debug_print(0, f"Looking for OneNote files in folder: {first_folder.get('name')}")
        onenote_files = find_onenote_files(drive_id, first_folder.get("id"))
    
    debug_print(0, "\n== Test Complete ==")
    debug_print(0, "Review the output above to understand the structure of your SharePoint")

def extract_onenote_from_sharepoint():
    """Extract OneNote notebooks from SharePoint document libraries associated with Teams"""
    if not ACCESS_TOKEN:
        debug_print(0, "ACCESS_TOKEN is not set. Please set it in the .env file or directly in the script.")
        return
    
    debug_print(0, "\n" + "="*80)
    debug_print(0, "Starting OneNote notebook extraction from SharePoint document libraries...")
    debug_print(0, "="*80)
    
    debug_print(0, "\n⚠️ NOTE: This script works best with the following Microsoft Graph permissions:")
    debug_print(0, "  - TeamSettings.Read.All: To access Teams")
    debug_print(0, "  - Sites.Read.All: To access SharePoint sites")
    debug_print(0, "  - Files.Read.All: To access document libraries")
    debug_print(0, "  - Notes.Read.All: To access OneNote notebooks")
    debug_print(0, "If you're seeing permission errors, ensure your token includes these scopes.\n")
    
    # Check if we're testing a specific team
    if TEST_TEAM_ID:
        debug_print(0, f"⚠️ Testing with specific team ID: {TEST_TEAM_ID}")
        test_single_team(TEST_TEAM_ID)
        return
    
    # Store for all discovered notebooks
    notebooks_data = []
    
    # Step 1: Get all teams
    teams = get_all_teams()
    if not teams:
        debug_print(0, "No teams found or error occurred.")
        return
    
    # Step 2: Process each team
    for team in teams:
        team_id = team.get("id")
        team_name = team.get("displayName", "Unknown Team")
        debug_print(0, f"\n{'='*50}")
        debug_print(0, f"🔍 Processing team: {team_name} ({team_id})")
        debug_print(0, f"{'='*50}")
        
        # Get team channels for proper channel ID mapping
        debug_print(0, f"Getting channels for team: {team_name}")
        channels, channel_map, channel_name_map = get_team_channels(team_id)
        debug_print(0, f"Found {len(channels)} channels in team")
        
        # Get SharePoint site for this team
        site = get_sharepoint_site_for_team(team_id)
        if not site:
            debug_print(0, f"Cannot find SharePoint site for team: {team_name}. Skipping team.")
            continue
            
        site_id = site.get("id")
        site_name = site.get("displayName", "Unknown")
        debug_print(0, f"SharePoint site: {site_name} (ID: {site_id})")
        
        # Step 3: Try to get notebooks directly from OneNote API first (most reliable)
        debug_print(0, f"Getting notebooks via OneNote API...")
        onenote_notebooks = get_notebooks_from_onenote_api(site_id)
        
        if not onenote_notebooks:
            debug_print(0, f"No notebooks found via OneNote API. Skipping team.")
            continue
            
        debug_print(0, f"Found {len(onenote_notebooks)} notebooks via OneNote API:")
        for notebook in onenote_notebooks:
            debug_print(0, f"  - Notebook: {notebook.get('displayName')} (ID: {notebook.get('id')})")
        
        # Create mapping of notebooks by name/id for later matching
        notebook_mapping = {}
        name_to_notebook = {}
        for notebook in onenote_notebooks:
            notebook_id = notebook.get("id")
            notebook_name = notebook.get("displayName", "")
            notebook_mapping[notebook_id] = notebook
            
            # Create a simplified version of the name for matching with folders
            simplified_name = notebook_name.lower()
            for prefix in ["bloc de notas de ", "notas_ ", "notas de ", "notebook "]:
                if simplified_name.startswith(prefix):
                    simplified_name = simplified_name[len(prefix):]
            
            name_to_notebook[simplified_name] = notebook
            debug_print(1, f"  Simplified name for matching: '{simplified_name}'")
        
        # Step 4: Get document library
        document_library = get_document_library(site_id)
        if not document_library:
            debug_print(0, f"Cannot find document library for site: {site_name}. Skipping site.")
            continue
            
        drive_id = document_library.get("id")
        drive_name = document_library.get("name", "Unknown")
        debug_print(0, f"Document library: {drive_name} (ID: {drive_id})")
        
        # Step 5: Get all root folders - these might be channels
        debug_print(0, f"Getting root folders which may represent channels...")
        root_folders_url = f"{graph_base_url}/drives/{drive_id}/root/children"
        root_response = make_request(root_folders_url)
        
        if not root_response or "value" not in root_response:
            debug_print(0, f"Cannot access root folders. Skipping team.")
            continue
            
        root_items = root_response["value"]
        root_folders = [item for item in root_items if item.get("folder")]
        
        debug_print(0, f"Found {len(root_folders)} root folders (potential channels):")
        for folder in root_folders:
            folder_name = folder.get("name", "Unknown")
            folder_id = folder.get("id", "")
            folder_size = folder.get("folder", {}).get("childCount", 0)
            debug_print(0, f"  - Folder: {folder_name} (ID: {folder_id}, Items: {folder_size})")
        
        # Step 6: Check each root folder for matching with a notebook
        processed_notebook_ids = set()
        
        for folder in root_folders:
            folder_name = folder.get("name", "Unknown")
            folder_id = folder.get("id", "")
            
            # Try to find a matching notebook by name similarity
            matched_notebook = None
            matched_similarity = 0
            simplified_folder_name = folder_name.lower()
            
            debug_print(1, f"Looking for notebook match for folder: '{simplified_folder_name}'")
            
            # First try exact match with simplified names
            for name, notebook in name_to_notebook.items():
                if name == simplified_folder_name or name in simplified_folder_name or simplified_folder_name in name:
                    matched_notebook = notebook
                    debug_print(0, f"  ✅ Found exact match between folder '{folder_name}' and notebook '{notebook.get('displayName')}'")
                    break
            
            # If no exact match, try partial match
            if not matched_notebook:
                for name, notebook in name_to_notebook.items():
                    # Calculate similarity - simple token overlap for now
                    folder_tokens = set(simplified_folder_name.split())
                    name_tokens = set(name.split())
                    common_tokens = folder_tokens.intersection(name_tokens)
                    
                    if common_tokens and len(common_tokens) > matched_similarity:
                        matched_similarity = len(common_tokens)
                        matched_notebook = notebook
                
                if matched_notebook:
                    debug_print(0, f"  ✅ Found partial match between folder '{folder_name}' and notebook '{matched_notebook.get('displayName')}'")
            
            # Step A: Process the folder as a channel with a matching notebook
            if matched_notebook:
                notebook_id = matched_notebook.get("id")
                notebook_name = matched_notebook.get("displayName")
                
                # Get sections for this notebook
                debug_print(0, f"Getting sections for notebook: {notebook_name}")
                sections = get_sections_for_notebook(notebook_id, site_id)
                
                # Default to folder as channel, but check if it's a generic name
                channel_name, channel_id = match_notebook_to_channel(notebook_name, channels, team_name)
                
                # Create data structure
                notebook_data = {
                    "team_name": team_name,
                    "team_id": team_id,
                    "channel_name": channel_name,
                    "channel_id": channel_id,
                    "notebook_name": notebook_name,
                    "notebook_id": notebook_id,
                    "sections": []
                }
                
                # Add sections
                for section in sections:
                    section_id = section.get("id")
                    section_name = section.get("displayName", "Unnamed Section")
                    
                    notebook_data["sections"].append({
                        "section_id": section_id,
                        "section_name": section_name
                    })
                    
                    debug_print(0, f"    - Section: {section_name} (ID: {section_id})")
                
                # Add to result list
                notebooks_data.append(notebook_data)
                processed_notebook_ids.add(notebook_id)
            
            # Step B: Also look for actual OneNote files in the folder
            debug_print(0, f"Looking for OneNote files in folder: {folder_name}")
            onenote_files = find_onenote_files(drive_id, folder_id)
            
            for onenote_file in onenote_files:
                file_name = onenote_file.get("name", "").replace(".one", "")
                file_id = onenote_file.get("id", "")
                web_url = onenote_file.get("webUrl", "")
                
                debug_print(0, f"  📓 Found OneNote file: {file_name}")
                
                # Try to match this file with a notebook from the OneNote API
                matched_api_notebook = None
                
                # Try matching by name
                simplified_file_name = file_name.lower()
                for name, notebook in name_to_notebook.items():
                    if name == simplified_file_name or name in simplified_file_name or simplified_file_name in name:
                        matched_api_notebook = notebook
                        debug_print(0, f"    ✅ Matched with notebook from OneNote API: {notebook.get('displayName')}")
                        break
                
                # If no match by name, try to extract notebook ID from URL
                if not matched_api_notebook:
                    extracted_id = extract_notebook_id_from_weburl(web_url)
                    if extracted_id and extracted_id in notebook_mapping:
                        matched_api_notebook = notebook_mapping[extracted_id]
                        debug_print(0, f"    ✅ Matched with notebook from URL extraction: {matched_api_notebook.get('displayName')}")
                
                # Skip if we already processed this notebook
                if matched_api_notebook and matched_api_notebook.get("id") in processed_notebook_ids:
                    debug_print(0, f"    ⚠️ Skipping notebook as it was already processed: {matched_api_notebook.get('displayName')}")
                    continue
                
                # Process the notebook if found
                if matched_api_notebook:
                    notebook_id = matched_api_notebook.get("id")
                    notebook_name = matched_api_notebook.get("displayName")
                    
                    # Get sections
                    sections = get_sections_for_notebook(notebook_id, site_id)
                    
                    # Try to match with a channel if folder_name is generic or unclear
                    channel_name, channel_id = match_notebook_to_channel(notebook_name, channels, team_name)
                    
                    # Create data structure
                    notebook_data = {
                        "team_name": team_name,
                        "team_id": team_id,
                        "channel_name": channel_name,
                        "channel_id": channel_id,
                        "notebook_name": notebook_name,
                        "notebook_id": notebook_id,
                        "match_source": "file_in_folder",
                        "sections": []
                    }
                    
                    # Add sections
                    for section in sections:
                        section_id = section.get("id")
                        section_name = section.get("displayName", "Unnamed Section")
                        
                        notebook_data["sections"].append({
                            "section_id": section_id,
                            "section_name": section_name
                        })
                        
                        debug_print(0, f"    - Section: {section_name} (ID: {section_id})")
                    
                    # Add to result list
                    notebooks_data.append(notebook_data)
                    processed_notebook_ids.add(notebook_id)
        
        # Step 7: Process any remaining notebooks that weren't matched to folders
        for notebook in onenote_notebooks:
            notebook_id = notebook.get("id")
            
            if notebook_id not in processed_notebook_ids:
                notebook_name = notebook.get("displayName")
                debug_print(0, f"Processing unmatched notebook: {notebook_name}")
                
                # Get sections
                sections = get_sections_for_notebook(notebook_id, site_id)
                
                # Try to match notebook with a channel
                channel_name, channel_id = match_notebook_to_channel(notebook_name, channels, team_name)
                
                # Create data structure
                notebook_data = {
                    "team_name": team_name,
                    "team_id": team_id,
                    "channel_name": channel_name,
                    "channel_id": channel_id,
                    "notebook_name": notebook_name,
                    "notebook_id": notebook_id,
                    "match_source": "api_only",
                    "sections": []
                }
                
                # Add sections
                for section in sections:
                    section_id = section.get("id")
                    section_name = section.get("displayName", "Unnamed Section")
                    
                    notebook_data["sections"].append({
                        "section_id": section_id,
                        "section_name": section_name
                    })
                    
                    debug_print(0, f"    - Section: {section_name} (ID: {section_id})")
                
                # Add to result list
                notebooks_data.append(notebook_data)
                processed_notebook_ids.add(notebook_id)
    
    # Final pass: Fix any remaining unknown channel IDs
    debug_print(0, f"\n{'='*50}")
    debug_print(0, f"Final pass: Checking for notebooks with unknown channel IDs")
    debug_print(0, f"{'='*50}")
    
    # Count how many unknown channel IDs we have
    unknown_count_before = len([entry for entry in notebooks_data if entry["channel_id"] == "unknown"])
    debug_print(0, f"Found {unknown_count_before} notebooks with unknown channel IDs")
    
    # Retry matching with more aggressive approach
    for entry in notebooks_data:
        if entry["channel_id"] == "unknown":
            team_id = entry["team_id"]
            team_name = entry["team_name"]
            notebook_name = entry["notebook_name"]
            
            debug_print(0, f"\nRetrying channel matching for notebook: {notebook_name}")
            debug_print(0, f"Team: {team_name}")
            
            # Get channels for this team
            channels, _, _ = get_team_channels(team_id)
            
            if channels:
                # Force a match by using the first available channel if all else fails
                channel_name, channel_id = match_notebook_to_channel(notebook_name, channels, team_name)
                
                if channel_id != "unknown":
                    entry["channel_name"] = channel_name
                    entry["channel_id"] = channel_id
                    debug_print(0, f"✅ Successfully matched notebook to channel: {channel_name} ({channel_id})")
                else:
                    # Ultimate fallback: use first channel
                    if channels:
                        entry["channel_name"] = channels[0].get("displayName", "First Channel")
                        entry["channel_id"] = channels[0].get("id", "unknown")
                        debug_print(0, f"⚠️ Using first available channel as fallback: {entry['channel_name']} ({entry['channel_id']})")
            else:
                debug_print(0, f"❌ No channels found for team: {team_name}")
    
    # Count how many unknown channel IDs remain
    unknown_count_after = len([entry for entry in notebooks_data if entry["channel_id"] == "unknown"])
    debug_print(0, f"\nFixed {unknown_count_before - unknown_count_after} unknown channel IDs, {unknown_count_after} remain")
    
    # If we still have unknown channel_ids, list them for debugging
    if unknown_count_after > 0:
        debug_print(0, "\nNotebooks with remaining unknown channel IDs:")
        for entry in notebooks_data:
            if entry["channel_id"] == "unknown":
                debug_print(0, f"  - {entry['notebook_name']} (Team: {entry['team_name']})")
    
    # Save all notebooks data to a JSON file
    output_file = "servitec_notebooks_data.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(notebooks_data, f, ensure_ascii=False, indent=4)
    
    # Print summary
    debug_print(0, f"\n{'='*80}")
    debug_print(0, f"✅ Extraction complete! Data saved to {output_file}")
    debug_print(0, f"📊 Total notebook entries found: {len(notebooks_data)}")
    
    # Group by teams for better summary
    teams_summary = {}
    for entry in notebooks_data:
        team_name = entry["team_name"]
        if team_name not in teams_summary:
            teams_summary[team_name] = {
                "notebooks": 0,
                "sections": 0,
                "channels": set()
            }
        
        teams_summary[team_name]["notebooks"] += 1
        teams_summary[team_name]["sections"] += len(entry["sections"])
        teams_summary[team_name]["channels"].add(entry["channel_name"])
    
    # Print detailed summary
    for team_name, summary in teams_summary.items():
        debug_print(0, f"  📊 Team '{team_name}': {summary['notebooks']} notebooks, {len(summary['channels'])} channels, {summary['sections']} total sections")
    
    debug_print(0, "\n⚠️ If you got permission errors, make sure your token has these Graph API permissions:")
    debug_print(0, "  - TeamSettings.Read.All")
    debug_print(0, "  - Sites.Read.All")
    debug_print(0, "  - Files.Read.All")
    debug_print(0, "  - Notes.Read.All")
    debug_print(0, "="*80)

if __name__ == "__main__":
    # Check for command line args to specify a team ID to test
    if len(sys.argv) > 1:
        TEST_TEAM_ID = sys.argv[1]
        debug_print(0, f"Using command line team ID: {TEST_TEAM_ID}")
    
    # Run the script
    extract_onenote_from_sharepoint()
