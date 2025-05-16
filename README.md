# Microsoft Teams Notebook Extraction

This repository contains scripts to extract notebooks from Microsoft Teams channels using the Microsoft Graph API. It helps map the relationships between Teams, Channels, and OneNote notebooks.

## 📌 **What This Code Does**
These scripts retrieve **Microsoft Teams OneNote Notebooks** and **Planner Tabs** from channels where you have access. It uses **Microsoft Graph API** to find **Notebooks** and **Planners (task lists)** linked to teams and extracts their IDs and content.

---

## 🔑 **Prerequisites**
To run this code, you need:

1. **A valid Microsoft Graph API Access Token**  
   - You can obtain this token **from the Microsoft Graph Explorer**:  
     - Go to [**Graph Explorer**](https://developer.microsoft.com/en-us/graph/graph-explorer)
     - Sign in with your **Microsoft account**
     - Click **"Access token"** (copy the token)

2. **Permissions Required**  
   Your Graph API token needs the following permissions:
   - `Team.ReadBasic.All` (to read team information)
   - `Channel.ReadBasic.All` (to read channels)
   - `Tab.Read.All` (to read Planner/OneNote tabs)
   - `Notes.Read.All` (to read notebook content)
   - `Group.Read.All` (to access team-related resources)

---

## 📂 **Files in this Repository**

### Core Extraction Scripts

- **`notebook_extraction.py`**: Main script that extracts OneNote notebooks from all your Teams channels. Creates a JSON file mapping Teams, channels, and notebooks. Uses the Microsoft Graph API to access Teams data and OneNote notebooks.

- **`explore_team_notebooks.py`**: Focused script that extracts notebook sections from a specific Teams channel. Uses direct API calls to access notebooks and their sections. Especially useful for exploring the structure of OneNote notebooks in Teams.

- **`servitec_notebook_extraction.py`**: Extended version of the notebook extraction script with additional features specific to Servitec's needs. Includes more detailed extraction of notebook content and metadata. In this company, each channel has an associated notebook to it, so it is first located in the Teams via the OneNote API and then the channel (SharePoint fodler) in which it is, is tracked.

### Data Processing

- **`planner_processing.ipynb`**: Jupyter notebook for processing and analyzing Tasks/Planner data from Teams. Helps visualize and manipulate task data extracted from Teams channels.

### Output Files (gitignored)

- **`teams_data.json`**: Contains extracted information about Teams and channels.
- **`teams_notebooks_data.json`**: Contains mappings between Teams, channels, and notebooks.
- **`notebook_sections_direct_api.json`**: Contains detailed information about notebook sections extracted via direct API calls.
- **`servitec_notebooks_data.json`**: Contains notebook data specific to Servitec teams.

### Documentation

- **`NOTEBOOK_README.md`**: Detailed documentation about the notebook extraction process, including setup instructions and output format.
- **`README.md`** (this file): Overview of the entire repository and its purpose.

### Configuration

- **`.env`**: Contains environment variables like the access token and target team/channel IDs. This file is not committed to git for security reasons.
- **`.gitignore`**: Specifies files that should not be tracked by git, including sensitive data files and environment variables.

---

## 🛠️ **How It Works**

### For Notebook Extraction
1. **Fetches All Teams you have access to**
2. **Retrieves All Channels in each Team**
3. **Identifies OneNote Tabs in each Channel**
4. **Extracts Notebook IDs and metadata**
5. **Accesses Notebook Sections through the API**
6. **Generates structured JSON output**

### For Planner/Tasks
1. **Fetches All Teams**  
2. **Fetches All Channels in Each Team**  
3. **Finds Planner Tabs in Each Channel**  
4. **Extracts the Planner ID from the URL**  
5. **Generates JSON Output**  

---

## 🚀 **How to Use**

1. **Create a `.env` file** with your access token:
   ```
   ACCESS_TOKEN=your_access_token_here
   TARGET_TEAM_ID=your_target_team_id  # optional, for specific team
   TARGET_CHANNEL_ID=your_target_channel_id  # optional, for specific channel
   ```

2. **Install required packages**:
   ```bash
   pip install requests python-dotenv
   ```

3. **Run the appropriate script**:
   ```bash
   python notebook_extraction.py  # For general notebook extraction
   # OR
   python explore_team_notebooks.py  # For detailed section extraction
   ```

The scripts will generate JSON files with the extracted data.

---

## 🎯 **Why This Is Important**

- **For Documentation**: Map all OneNote notebooks in your Teams for better knowledge management.
- **For Migration**: Prepare data for migration to other systems or backup.
- **For Analysis**: Understand how information is structured across teams and channels.
- **For Integration**: Build upon these scripts to create tools that integrate with other systems.

---

## 🔒 **Security Note**

- The `.env` file and JSON output files are excluded from git to prevent leaking sensitive information.
- Always protect your access tokens and never commit them to public repositories.
- The `.gitignore` file is configured to exclude sensitive data files (all `.json` files).
