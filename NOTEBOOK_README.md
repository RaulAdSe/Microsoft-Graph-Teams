# Microsoft Teams Notebook Extraction

This script extracts notebook and section information from Microsoft Teams channels using the Microsoft Graph API.

## Setup

1. Make sure you have Python 3.6+ installed
2. Install required packages:
   ```bash
   pip install requests python-dotenv
   ```

3. Set up your access token:
   - Create a `.env` file in the same directory as the script
   - Add your Microsoft Graph API access token:
     ```
     ACCESS_TOKEN=your_access_token_here
     ```
   - Alternatively, you can add the token directly in the script

## How to Get an Access Token

To get a Microsoft Graph API access token:

1. Sign in to the [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Click on your profile icon and sign in with your Microsoft account
3. Grant the necessary permissions
4. Make a simple request (e.g., `/me`)
5. Open browser developer tools (F12)
6. Go to the Network tab
7. Look for a request to the Microsoft Graph API
8. Find the Authorization header which contains your token (starts with "Bearer ")
9. Copy the token part (without "Bearer ")

## Running the Script

Simply run the Python script:

```bash
python notebook_extraction.py
```

The script will:
1. Fetch all teams you have access to
2. For each team, get all channels
3. For each team, attempt to get all notebooks 
4. For each notebook, get all sections
5. Create records linking teams, channels, notebooks, and sections
6. Save the data to `teams_notebooks_data.json`

## Output Format

The script generates a JSON file with the following structure:

```json
[
  {
    "project_name": "Team Name",
    "team_id": "team-id-here",
    "channel_name": "Channel Name",
    "channel_id": "channel-id-here",
    "notebook_id": "notebook-id-here",
    "notebook_name": "Notebook Name",
    "sections": [
      {
        "section_id": "section-id-here",
        "section_name": "Section Name",
        "section_url": "section-url-here"
      }
    ]
  }
]
```

## Troubleshooting

- If the script fails to find notebooks, it will try an alternative API endpoint
- Make sure your access token is valid and has the necessary permissions:
  - TeamsActivity.Read.All
  - TeamsAppInstallation.ReadForTeam
  - TeamMember.Read.All
  - TeamsTab.Read.All
  - User.Read.All
  - Group.Read.All
  - Notes.Read.All

## Alternative Data Formats

If you need a different output format, modify the script as needed:
- To flatten the data and create one record per section, modify the `team_notebooks_data.append` call
- To include additional information, add more fields to the output dictionaries 