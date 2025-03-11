## 📌 **What This Code Does**
This script retrieves **Microsoft Teams Planner Tabs** from all channels in all teams where you have access. It uses **Microsoft Graph API** to find **Planners (task lists)** linked to teams and extracts their **Planner IDs**.

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
   - `Tab.Read.All` (to read Planner tabs)

---

## 🛠️ **How It Works**
### 1️⃣ **Fetches All Teams**  
   - Uses `https://graph.microsoft.com/v1.0/me/joinedTeams`  
   - Retrieves the **team name** and **team ID**.

### 2️⃣ **Fetches All Channels in Each Team**  
   - Uses `https://graph.microsoft.com/v1.0/teams/{team_id}/channels`  
   - Gets **channel names** and **channel IDs**.

### 3️⃣ **Finds Planner Tabs in Each Channel**  
   - Uses `https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/tabs`  
   - Filters **Planner tabs** based on:
     - Tab name (e.g., "Tasks", "Planner", "Tareas")
     - **Web URL** containing `tasks.office.com`

### 4️⃣ **Extracts the Planner ID from the URL**  
   - The real **Planner ID** is found in:  
     ```
     https://tasks.office.com/.../PlanViews/{PLANNER_ID}?...  
     ```
   - The script extracts `{PLANNER_ID}` and saves it.

### 5️⃣ **Generates JSON Output**  
   - The final JSON includes:
     - **Team name & ID**
     - **Channel name & ID**
     - **Planner Tabs**
       - Tab name
       - Extracted **Planner ID**
       - **Planner URL** (for reference)

---

## 📥 **Example Output**
```json
[
    {
        "project_name": "ACTIV GROUP",
        "team_id": "d3a1f29a-40aa-4e27-b286-21f7dc5f3f7b",
        "channel_name": "General",
        "channel_id": "19:0020527fe41e4a35902ca528b6ca31d3@thread.skype",
        "planner_tabs": [
            {
                "tab_name": "Tasks ACTIV GROUP TERRASSA",
                "planner_id": "QqtnCj433E-tc_EuxF0YHZcAGsto",
                "planner_url": "https://tasks.office.com/.../PlanViews/QqtnCj433E-tc_EuxF0YHZcAGsto"
            }
        ]
    }
]
```
- ✅ **Planner ID is correctly extracted** from the URL.
- ✅ **Each Planner tab is listed under its corresponding channel & team**.

---

## 🚀 **How to Use in Jupyter Notebook**
### **Instructions on Top of Your Notebook**
```markdown
# 📌 Fetch Planner Tabs from Microsoft Teams
## **Requirements:**
1️⃣ Get a **Graph API Access Token** from [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
   - Sign in → Click "Access Token" → Copy the token

2️⃣ Paste the token into `headers`:
```python
headers = {
    "Authorization": "Bearer YOUR_ACCESS_TOKEN",
    "Content-Type": "application/json"
}
```
3️⃣ Run the script! It will:
   - Retrieve **Teams & Channels**
   - Find **Planner Tabs**
   - Extract **Real Planner IDs**
   - Generate a JSON output of all detected Planners
```

---

### **🎯 Why This is Useful?**
- **For Auditing:** List all active Planners in your Microsoft Teams.
- **For Automation:** Use this JSON in scripts to automate Planner-related tasks.
- **For Data Validation:** Ensure **Planner IDs are correct** before using them in another system.

Let me know if you need further refinements! 🚀
