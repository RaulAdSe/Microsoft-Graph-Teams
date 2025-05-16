#!/usr/bin/env python3
# Microsoft Teams Notebook Extraction
# This script extracts OneNote notebooks from Teams channel tabs via direct API calls

import requests
import json
import time
from dotenv import load_dotenv
import os
from pprint import pprint

# Load environment variables
load_dotenv()

# Access token from .env
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")

# If no token in .env, show error message
if not ACCESS_TOKEN:
    print("ERROR: No ACCESS_TOKEN found in environment variables")
    print("Please create a .env file with your ACCESS_TOKEN or set it as an environment variable")
    print("You can obtain a token from Microsoft Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer")
    print("Example .env file content: ACCESS_TOKEN=your_token_here")
    import sys
    sys.exit(1)

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Accept": "application/json"
}

# Microsoft Graph API base URL
graph_base_url = "https://graph.microsoft.com/v1.0"

def make_request(url, method="GET"):
    """Make a request to the Microsoft Graph API"""
    try:
        print(f"Making request to: {url}")
        if method == "GET":
            response = requests.get(url, headers=headers)
        else:
            print(f"Unsupported method: {method}")
            return None
            
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Exception making request: {str(e)}")
        return None

def get_all_teams():
    """Get all teams the user is a member of"""
    print("Fetching all teams...")
    teams_url = f"{graph_base_url}/me/joinedTeams"
    
    response = make_request(teams_url)
    if response and "value" in response:
        teams = response["value"]
        print(f"Found {len(teams)} teams")
        return teams
    
    print("No teams found or error occurred")
    return []

def get_team_channels(team_id):
    """Get all channels for a specific team"""
    print(f"Getting channels for team ID: {team_id}")
    channels_url = f"{graph_base_url}/teams/{team_id}/channels"
    
    response = make_request(channels_url)
    if response and "value" in response:
        channels = response["value"]
        print(f"Found {len(channels)} channels")
        return channels
    
    print("No channels found or error occurred")
    return []

def get_channel_tabs(team_id, channel_id):
    """Get all tabs in a specific channel"""
    print(f"Getting tabs for channel ID: {channel_id}")
    tabs_url = f"{graph_base_url}/teams/{team_id}/channels/{channel_id}/tabs"
    
    response = make_request(tabs_url)
    if response and "value" in response:
        tabs = response["value"]
        print(f"Found {len(tabs)} tabs")
        return tabs
    
    print("No tabs found or error occurred")
    return []

def get_sharepoint_site_for_team(team_id):
    """Get the SharePoint site associated with a team"""
    print(f"Getting SharePoint site for team ID: {team_id}")
    site_url = f"{graph_base_url}/groups/{team_id}/sites/root"
    
    response = make_request(site_url)
    if response and "id" in response:
        site_id = response.get("id")
        site_name = response.get("displayName", "Unknown")
        print(f"Found SharePoint site: {site_name} (ID: {site_id})")
        return response
    
    print("SharePoint site not found or permission error")
    if "403" in str(response):
        print("⚠️ Permission issue: Your token likely lacks Sites.Read.All permissions")
    
    return None

def get_notebooks_in_group(group_id):
    """Get all OneNote notebooks in a group (team) directly via API"""
    print(f"Getting notebooks in group/team ID: {group_id}")
    notebooks_url = f"{graph_base_url}/groups/{group_id}/onenote/notebooks"
    
    response = make_request(notebooks_url)
    if response and "value" in response:
        notebooks = response["value"]
        print(f"Found {len(notebooks)} notebooks in group/team")
        return notebooks
    
    print("No notebooks found in group or error occurred")
    if "403" in str(response) or "401" in str(response):
        print("⚠️ Permission issue: Your token lacks Group.Read.All and/or Notes.Read.All permissions")
    
    return []

def get_notebooks_in_site(site_id):
    """Get all OneNote notebooks in a SharePoint site"""
    print(f"Getting notebooks in SharePoint site ID: {site_id}")
    notebooks_url = f"{graph_base_url}/sites/{site_id}/onenote/notebooks"
    
    response = make_request(notebooks_url)
    if response and "value" in response:
        notebooks = response["value"]
        print(f"Found {len(notebooks)} notebooks in SharePoint site")
        return notebooks
    
    print("No notebooks found in site or error occurred")
    if "403" in str(response) or "401" in str(response):
        print("⚠️ Permission issue: Your token lacks Sites.Read.All and/or Notes.Read.All permissions")
    
    return []

def get_tab_notebook_info(tab):
    """Extract notebook info from a OneNote tab using the tab's properties"""
    tab_name = tab.get("displayName", "Unknown")
    tab_id = tab.get("id", "Unknown")
    print(f"\nExamining OneNote tab: {tab_name} (ID: {tab_id})")
    
    # Print configuration to help debug
    configuration = tab.get("configuration", {})
    if configuration:
        print("Tab configuration:")
        pprint(configuration)
        
        # Look for entityId or contentUrl in the configuration
        entity_id = configuration.get("entityId")
        content_url = configuration.get("contentUrl")
        
        if entity_id and "notebook" in str(entity_id).lower():
            print(f"Found entity ID in tab configuration: {entity_id}")
            return {"notebook_id": entity_id, "source": "entity_id"}
        
        if content_url and "onenote" in content_url.lower():
            print(f"Found OneNote content URL in tab configuration: {content_url}")
            return {"content_url": content_url, "source": "content_url"}
    
    # If we couldn't get info from configuration, try the tab's webUrl
    web_url = tab.get("webUrl")
    if web_url:
        print(f"Using tab webUrl: {web_url}")
        return {"web_url": web_url, "source": "web_url"}
    
    print("No notebook information found in tab properties")
    return None

def get_sections_for_notebook(notebook_id, group_id=None, site_id=None):
    """Get sections for a specific notebook"""
    print(f"Getting sections for notebook ID: {notebook_id}")
    
    # Try different endpoints in priority order
    sections = None
    
    # 1. Try group/team path first (most common for Teams notebooks)
    if group_id:
        print(f"Trying group endpoint for sections...")
        url = f"{graph_base_url}/groups/{group_id}/onenote/notebooks/{notebook_id}/sections"
        response = make_request(url)
        
        if response and "value" in response:
            sections = response["value"]
            print(f"Found {len(sections)} sections via group endpoint")
            return sections
        
        print("Group endpoint failed")
    
    # 2. Try SharePoint site path if we have site_id
    if site_id and not sections:
        print(f"Trying site endpoint for sections...")
        url = f"{graph_base_url}/sites/{site_id}/onenote/notebooks/{notebook_id}/sections"
        response = make_request(url)
        
        if response and "value" in response:
            sections = response["value"]
            print(f"Found {len(sections)} sections via site endpoint")
            return sections
        
        print("Site endpoint failed")
    
    # 3. Try personal endpoint as fallback
    if not sections:
        print(f"Trying personal endpoint for sections...")
        url = f"{graph_base_url}/me/onenote/notebooks/{notebook_id}/sections"
        response = make_request(url)
        
        if response and "value" in response:
            sections = response["value"]
            print(f"Found {len(sections)} sections via personal endpoint")
            return sections
        
        print("Personal endpoint failed")
    
    # 4. Try direct filter as last resort
    if not sections:
        print(f"Trying direct filter endpoint for sections...")
        url = f"{graph_base_url}/me/onenote/sections?$filter=parentNotebook/id eq '{notebook_id}'"
        response = make_request(url)
        
        if response and "value" in response:
            sections = response["value"]
            print(f"Found {len(sections)} sections via filter endpoint")
            return sections
        
        print("Filter endpoint failed")
    
    print("No sections found for this notebook through any endpoint")
    return []

def get_notebook_details(notebook_id, group_id=None, site_id=None):
    """Get details for a specific notebook"""
    print(f"Getting details for notebook ID: {notebook_id}")
    
    # Try different endpoints in priority order
    notebook = None
    
    # 1. Try group/team path first (most common for Teams notebooks)
    if group_id:
        print(f"Trying group endpoint for notebook details...")
        url = f"{graph_base_url}/groups/{group_id}/onenote/notebooks/{notebook_id}"
        response = make_request(url)
        
        if response and "id" in response:
            notebook = response
            print(f"Found notebook details via group endpoint: {notebook.get('displayName')}")
            return notebook
        
        print("Group endpoint failed")
    
    # 2. Try SharePoint site path if we have site_id
    if site_id and not notebook:
        print(f"Trying site endpoint for notebook details...")
        url = f"{graph_base_url}/sites/{site_id}/onenote/notebooks/{notebook_id}"
        response = make_request(url)
        
        if response and "id" in response:
            notebook = response
            print(f"Found notebook details via site endpoint: {notebook.get('displayName')}")
            return notebook
        
        print("Site endpoint failed")
    
    # 3. Try personal endpoint as fallback
    if not notebook:
        print(f"Trying personal endpoint for notebook details...")
        url = f"{graph_base_url}/me/onenote/notebooks/{notebook_id}"
        response = make_request(url)
        
        if response and "id" in response:
            notebook = response
            print(f"Found notebook details via personal endpoint: {notebook.get('displayName')}")
            return notebook
        
        print("Personal endpoint failed")
    
    print("Could not get notebook details through any endpoint")
    return None

def is_onenote_tab(tab):
    """Check if a tab is a OneNote tab"""
    # Check tab name
    if "OneNote" in tab.get("displayName", ""):
        return True
    
    # Check teamsAppId (OneNote app ID)
    if tab.get("teamsAppId") == "0d820ecd-def2-4297-adad-78056cde7c78":
        return True
    
    # Check configuration
    config = tab.get("configuration", {})
    content_url = config.get("contentUrl", "")
    if content_url and "onenote" in content_url.lower():
        return True
    
    # Check webUrl
    web_url = tab.get("webUrl", "")
    if web_url and "onenote" in web_url.lower():
        return True
    
    return False

def extract_onenote_notebooks_from_teams():
    """Extract OneNote notebooks from Teams channel tabs via direct API calls"""
    if not ACCESS_TOKEN:
        print("ACCESS_TOKEN is not set. Please set it in the .env file or directly in the script.")
        return
    
    print("\n" + "="*80)
    print("Starting OneNote notebook extraction from Teams channel tabs...")
    print("="*80)
    
    print("\n⚠️ NOTE: This script works best with the following Microsoft Graph permissions:")
    print("  - TeamSettings.Read.All: To access Teams and channels")
    print("  - Sites.Read.All: To access SharePoint sites")
    print("  - Group.Read.All: To access Groups/Teams data")
    print("  - Notes.Read.All: To access OneNote notebooks")
    print("If you're seeing permission errors, ensure your token includes these scopes.\n")
    
    # Step 1: Get all teams
    teams = get_all_teams()
    if not teams:
        print("No teams found or error occurred.")
        return
    
    # Store for all discovered notebooks
    notebooks_data = []
    
    # Keep track of processed notebook IDs to avoid duplicates
    processed_notebook_ids = set()
    
    # Step 2: Process each team
    for team in teams:
        team_id = team.get("id")
        project_name = team.get("displayName", "Unknown Team")
        print(f"\n{'='*50}")
        print(f"🔍 Processing team: {project_name} ({team_id})")
        print(f"{'='*50}")
        
        # Get SharePoint site for this team
        site = get_sharepoint_site_for_team(team_id)
        site_id = site.get("id") if site else None
        
        # Get all notebooks for this team directly from the group API
        group_notebooks = get_notebooks_in_group(team_id)
        
        # Process group notebooks first (most reliable method)
        print(f"\n📚 Processing {len(group_notebooks)} notebooks found directly in group API")
        for notebook in group_notebooks:
            notebook_id = notebook.get("id")
            notebook_name = notebook.get("displayName", "Unnamed Notebook")
            
            # Skip if already processed
            if notebook_id in processed_notebook_ids:
                print(f"Skipping already processed notebook: {notebook_name}")
                continue
            
            processed_notebook_ids.add(notebook_id)
            print(f"\n  📕 Processing group notebook: {notebook_name} (ID: {notebook_id})")
            
            # Get sections for this notebook
            sections = get_sections_for_notebook(notebook_id, team_id, site_id)
            
            # Create notebook data structure
            notebook_data = {
                "notebook_id": notebook_id,
                "notebook_name": notebook_name,
                "project_name": project_name,
                "team_id": team_id,
                "source": "group_api",
                "sections": []
            }
            
            # Add sections to notebook data
            for section in sections:
                section_id = section.get("id")
                section_name = section.get("displayName", "Unnamed Section")
                
                notebook_data["sections"].append({
                    "section_id": section_id,
                    "section_name": section_name
                })
                
                print(f"    - Section: {section_name}")
            
            # Add notebook to final results
            notebooks_data.append(notebook_data)
        
        # Get all channels for this team
        channels = get_team_channels(team_id)
        
        # Step 3: Process each channel and its tabs
        for channel in channels:
            channel_id = channel.get("id")
            channel_name = channel.get("displayName", "Unknown Channel")
            print(f"\n  {'='*40}")
            print(f"  📊 Processing channel: {channel_name}")
            print(f"  {'='*40}")
            
            # Get all tabs for this channel
            tabs = get_channel_tabs(team_id, channel_id)
            
            # Step 4: Look for OneNote tabs
            for tab in tabs:
                if is_onenote_tab(tab):
                    tab_name = tab.get("displayName", "Unnamed Tab")
                    tab_id = tab.get("id")
                    print(f"\n    🔍 Found OneNote tab: {tab_name}")
                    
                    # Get notebook info from tab properties
                    notebook_info = get_tab_notebook_info(tab)
                    
                    if not notebook_info:
                        print(f"    ⚠️ Could not extract notebook information from tab")
                        continue
                    
                    # If we have a direct notebook ID from tab properties
                    if "notebook_id" in notebook_info:
                        notebook_id = notebook_info["notebook_id"]
                        
                        # Skip if already processed
                        if notebook_id in processed_notebook_ids:
                            print(f"    ⏭️ Skipping already processed notebook: {notebook_id}")
                            continue
                        
                        processed_notebook_ids.add(notebook_id)
                        
                        # Get notebook details
                        notebook_details = get_notebook_details(notebook_id, team_id, site_id)
                        
                        # If we couldn't get details, create minimal details
                        if not notebook_details:
                            print(f"    ⚠️ Using minimal notebook details based on tab name")
                            notebook_details = {
                                "id": notebook_id,
                                "displayName": tab_name.replace(" (OneNote)", "").strip()
                            }
                        
                        notebook_name = notebook_details.get("displayName", "Unnamed Notebook")
                        print(f"    📚 Notebook name: {notebook_name}")
                        print(f"    📚 Notebook ID: {notebook_id}")
                        
                        # Get sections for this notebook
                        print(f"    📑 Getting sections for notebook: {notebook_name}")
                        sections = get_sections_for_notebook(notebook_id, team_id, site_id)
                        
                        # Create notebook data structure
                        notebook_data = {
                            "notebook_id": notebook_id,
                            "notebook_name": notebook_name,
                            "project_name": project_name,
                            "team_id": team_id,
                            "channel_name": channel_name,
                            "channel_id": channel_id,
                            "tab_name": tab_name,
                            "source": "tab_configuration",
                            "sections": []
                        }
                        
                        # Add sections to notebook data
                        for section in sections:
                            section_id = section.get("id")
                            section_name = section.get("displayName", "Unnamed Section")
                            
                            notebook_data["sections"].append({
                                "section_id": section_id,
                                "section_name": section_name
                            })
                            
                            print(f"      - Section: {section_name}")
                        
                        # Add notebook to final results
                        notebooks_data.append(notebook_data)
    
    # Save all notebooks data to a JSON file
    output_file = "teams_notebooks_data.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(notebooks_data, f, ensure_ascii=False, indent=4)
    
    # Print summary
    print(f"\n{'='*80}")
    print(f"✅ Extraction complete! Data saved to {output_file}")
    print(f"📊 Total notebooks found in Teams: {len(notebooks_data)}")
    print(f"📊 Total teams processed: {len(teams)}")
    
    print("\n⚠️ If you got permission errors, make sure your token has these Graph API permissions:")
    print("  - TeamSettings.Read.All")
    print("  - Sites.Read.All")
    print("  - Group.Read.All")
    print("  - Notes.Read.All")
    print("="*80)

if __name__ == "__main__":
    extract_onenote_notebooks_from_teams() 