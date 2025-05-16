#!/usr/bin/env python3
"""
Simple script to access notebook sections in a specific Teams channel.
Uses direct API calls rather than URL parsing to access notebooks.
"""

import os
import json
import requests
import sys
from pprint import pprint
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get access token from environment variables
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")

# Verify access token is available
if not ACCESS_TOKEN:
    print("ERROR: No ACCESS_TOKEN found in environment variables")
    print("Please create a .env file with your ACCESS_TOKEN or set it as an environment variable")
    print("You can obtain a token from Microsoft Graph Explorer: https://developer.microsoft.com/en-us/graph/graph-explorer")
    sys.exit(1)

# Target team and channel IDs
TARGET_TEAM_ID = os.getenv("TARGET_TEAM_ID")
TARGET_CHANNEL_ID = os.getenv("TARGET_CHANNEL_ID")

# Verify required IDs are available
if not TARGET_TEAM_ID or not TARGET_CHANNEL_ID:
    print("WARNING: TARGET_TEAM_ID or TARGET_CHANNEL_ID not found in environment variables")
    print("Please set these in your .env file to specify which team/channel to explore")

# Headers for API requests
headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

def make_request(url):
    """Make a request to the Microsoft Graph API"""
    try:
        print(f"Making request to: {url}")
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Error: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Exception: {str(e)}")
        return None

def get_channel_tabs(team_id, channel_id):
    """Get all tabs in a channel"""
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/tabs"
    response = make_request(url)
    if response and "value" in response:
        return response["value"]
    return []

def get_group_notebooks(group_id):
    """Get all notebooks in a group/team directly via API"""
    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/onenote/notebooks"
    response = make_request(url)
    if response and "value" in response:
        print(f"Found {len(response['value'])} notebooks in group")
        return response["value"]
    print("No notebooks found in group or error occurred")
    return []

def get_tab_notebook_info(tab):
    """Extract notebook info from a OneNote tab using the tab's properties"""
    # Try to get the notebook information from the tab's configuration
    tab_name = tab.get("displayName", "Unknown")
    print(f"\nExamining tab: {tab_name}")
    
    # Print all tab properties to help debug
    print("Tab properties:")
    pprint(tab)
    
    # The configuration property should contain the notebook information
    configuration = tab.get("configuration", {})
    if configuration:
        print("Tab configuration:")
        pprint(configuration)
        
        # Look for entityId or contentUrl in the configuration
        entity_id = configuration.get("entityId")
        content_url = configuration.get("contentUrl")
        
        if entity_id and "notebook" in entity_id:
            print(f"Found entity ID: {entity_id}")
            # Entity ID often contains the notebook ID
            return {"notebook_id": entity_id, "source": "entity_id"}
        
        if content_url and "onenote" in content_url.lower():
            print(f"Found content URL: {content_url}")
            return {"content_url": content_url, "source": "content_url"}
    
    # If we couldn't get info from configuration, try other tab properties
    web_url = tab.get("webUrl")
    if web_url:
        print(f"Using tab webUrl as fallback: {web_url}")
        return {"web_url": web_url, "source": "web_url"}
    
    return None

def get_sections_for_notebook(notebook_id, group_id):
    """Get all sections in a notebook using different endpoints"""
    print(f"Getting sections for notebook ID: {notebook_id}")
    
    # Try group path first (most reliable for Teams notebooks)
    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/onenote/notebooks/{notebook_id}/sections"
    response = make_request(url)
    
    if response and "value" in response:
        print(f"Found {len(response['value'])} sections via group endpoint")
        return response["value"]
    
    # Try personal path as fallback
    print("Group endpoint failed, trying personal endpoint...")
    url = f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebook_id}/sections"
    response = make_request(url)
    
    if response and "value" in response:
        print(f"Found {len(response['value'])} sections via personal endpoint")
        return response["value"]
    
    print("No sections found for this notebook or errors occurred")
    return []

def get_onenote_tabs(tabs):
    """Filter tabs to find OneNote tabs"""
    onenote_tabs = []
    for tab in tabs:
        # OneNote tabs usually have "OneNote" in the name or a specific teamsAppId
        if ("OneNote" in tab.get("displayName", "") or 
            tab.get("teamsAppId") == "0d820ecd-def2-4297-adad-78056cde7c78"):
            onenote_tabs.append(tab)
    return onenote_tabs

def access_notebook_sections():
    """Access notebook sections in a specific Teams channel using direct API calls"""
    # Step 1: Get all tabs in the channel
    print(f"Getting tabs for team {TARGET_TEAM_ID}, channel {TARGET_CHANNEL_ID}")
    tabs = get_channel_tabs(TARGET_TEAM_ID, TARGET_CHANNEL_ID)
    
    # Step 2: Find OneNote tabs
    onenote_tabs = get_onenote_tabs(tabs)
    print(f"Found {len(onenote_tabs)} OneNote tabs")
    
    # Step 3: Get group notebooks directly (most reliable method)
    all_sections = []
    group_notebooks = get_group_notebooks(TARGET_TEAM_ID)
    
    for notebook in group_notebooks:
        notebook_id = notebook.get("id")
        notebook_name = notebook.get("displayName")
        
        print(f"\nProcessing group notebook: {notebook_name} (ID: {notebook_id})")
        sections = get_sections_for_notebook(notebook_id, TARGET_TEAM_ID)
        
        if sections:
            print(f"Found {len(sections)} sections")
            for section in sections:
                print(f"  - {section.get('displayName')} (ID: {section.get('id')})")
            
            all_sections.append({
                "notebook_name": notebook_name,
                "notebook_id": notebook_id,
                "source": "group_api",
                "sections": sections
            })
    
    # Step 4: For each OneNote tab, try to get notebook information directly
    for tab in onenote_tabs:
        notebook_info = get_tab_notebook_info(tab)
        
        if notebook_info:
            print(f"\nTab notebook info: {notebook_info}")
               
            # If we got a notebook ID directly, use it
            if "notebook_id" in notebook_info:
                notebook_id = notebook_info["notebook_id"]
                sections = get_sections_for_notebook(notebook_id, TARGET_TEAM_ID)
                
                if sections:
                    all_sections.append({
                        "tab_name": tab.get("displayName"),
                        "notebook_id": notebook_id,
                        "source": "tab_configuration",
                        "sections": sections
                    })
    
    # Step 5: Save results
    result_file = "notebook_sections_direct_api.json"
    with open(result_file, "w") as f:
        json.dump(all_sections, f, indent=4)
    
    print(f"\nSaved all sections to {result_file}")
    return all_sections

if __name__ == "__main__":
    access_notebook_sections() 