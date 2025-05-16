import requests
from collections import defaultdict
import pandas as pd
import json
import logging
import sys
import os

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Get access token from environment variable instead of hardcoding it
ACCESS_TOKEN = os.environ.get("MS_ACCESS_TOKEN", "")
if not ACCESS_TOKEN:
    logger.error("ERROR: No access token found. Please set the MS_ACCESS_TOKEN environment variable.")
    logger.info("You can obtain a new token from Microsoft Graph Explorer or by using authentication libraries.")
    sys.exit(1)

PLAN_ID = "_PaccKc6iU28KYRTAdz7GpcACFmq"

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

base_url = "https://graph.microsoft.com/v1.0"

def get_buckets(plan_id):
    try:
        logger.info(f"Fetching buckets for plan: {plan_id}")
        url = f"{base_url}/planner/plans/{plan_id}/buckets"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        buckets = response.json().get("value", [])
        logger.info(f"Successfully retrieved {len(buckets)} buckets")
        return buckets
    except requests.exceptions.RequestException as e:
        logger.error(f"Error fetching buckets: {e}")
        raise

def get_tasks(plan_id):
    try:
        logger.info(f"Fetching tasks for plan: {plan_id}")
        url = f"{base_url}/planner/plans/{plan_id}/tasks"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        tasks = response.json().get("value", [])
        logger.info(f"Successfully retrieved {len(tasks)} tasks")
        return tasks
    except requests.exceptions.RequestException as e:
        logger.error(f"Error fetching tasks: {e}")
        raise

def get_task_details(task_id):
    try:
        logger.info(f"Fetching details for task: {task_id}")
        url = f"{base_url}/planner/tasks/{task_id}/details"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logger.error(f"Error fetching task details for task {task_id}: {e}")
        return {}  # Return empty dict to allow processing to continue

def create_hierarchical_rows(buckets, bucket_map):
    """
    Create hierarchical rows for Excel with visual spacing and formatting.
    """
    hierarchical_rows = []
    
    for bucket in buckets:
        # Add a blank row before each bucket for better separation
        hierarchical_rows.append({
            "Bucket": "",
            "Task": "",
            "Checklist Item": ""
        })
        
        # Add bucket as a header row
        hierarchical_rows.append({
            "Bucket": f"📁 {bucket['name']}",
            "Task": "",
            "Checklist Item": ""
        })
        
        bucket_tasks = bucket_map.get(bucket["id"], [])
        
        if not bucket_tasks:
            # Add empty row if no tasks
            hierarchical_rows.append({
                "Bucket": "",
                "Task": "(No tasks)",
                "Checklist Item": ""
            })
            continue
            
        for task in bucket_tasks:
            # Calculate completion percentage for the task
            details = get_task_details(task["id"])
            checklist = details.get("checklist", {})
            
            completion_text = ""
            if checklist:
                total_items = len(checklist)
                completed_items = sum(1 for item in checklist.values() if item["isChecked"])
                
                if total_items > 0:
                    completion_percentage = (completed_items / total_items) * 100
                    status_symbol = "🔴" if completion_percentage == 0 else "🟡" if completion_percentage < 100 else "🟢"
                    completion_text = f" {status_symbol} ({completed_items}/{total_items})"
            
            # Add task row with indentation and completion info
            hierarchical_rows.append({
                "Bucket": "",
                "Task": f"📌 {task['title']}{completion_text}",
                "Checklist Item": ""
            })
            
            if checklist:
                for item_id, item in checklist.items():
                    # Add checklist item with status
                    status = "✅" if item["isChecked"] else "⬜"
                    hierarchical_rows.append({
                        "Bucket": "",
                        "Task": "",
                        "Checklist Item": f"{status} {item['title']}"
                    })
            else:
                # Add a message if no checklist
                hierarchical_rows.append({
                    "Bucket": "",
                    "Task": "",
                    "Checklist Item": "(No checklist items)"
                })
    
    return hierarchical_rows

def export_to_excel(rows, filename="planner_modelo.xlsx"):
    try:
        logger.info(f"Exporting data to Excel: {filename}")
        df = pd.DataFrame(rows)
        
        # Create a styled Excel writer
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Planner Tasks', index=False)
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Planner Tasks']
            
            # Set column widths
            worksheet.column_dimensions['A'].width = 30  # Bucket column
            worksheet.column_dimensions['B'].width = 40  # Task column
            worksheet.column_dimensions['C'].width = 60  # Checklist item column
            
            # Style formatting is simplified to avoid compatibility issues
        
        logger.info(f"✅ Successfully exported to {filename}")
        return True
    except Exception as e:
        logger.error(f"Failed to export to Excel: {e}")
        return False

def main():
    try:
        logger.info("Starting Microsoft Planner export process")
        
        # Fetch data
        buckets = get_buckets(PLAN_ID)
        tasks = get_tasks(PLAN_ID)

        bucket_map = defaultdict(list)
        for task in tasks:
            bucket_map[task["bucketId"]].append(task)

        logger.info("Processing tasks and checklist items")
        
        # Create hierarchical rows for better visualization
        hierarchical_rows = create_hierarchical_rows(buckets, bucket_map)
        
        logger.info(f"Created {len(hierarchical_rows)} rows for hierarchical display")
        
        # Export to Excel with formatting
        export_to_excel(hierarchical_rows, "planner_modelo_hierarchical.xlsx")
        
        # Export to JSON (original format for compatibility)
        try:
            # Create regular rows for JSON export (keeping original format)
            regular_rows = []
            for bucket in buckets:
                for task in bucket_map.get(bucket["id"], []):
                    details = get_task_details(task["id"])
                    checklist = details.get("checklist", {})
                    if checklist:
                        for item in checklist.values():
                            regular_rows.append({
                                "Bucket": bucket["name"],
                                "Task": task["title"],
                                "Checklist item": item["title"],
                                "Completed": item["isChecked"]
                            })
                    else:
                        regular_rows.append({
                            "Bucket": bucket["name"],
                            "Task": task["title"],
                            "Checklist item": "(no checklist)",
                            "Completed": ""
                        })
            
            output_json = "planner_modelo.json"
            logger.info(f"Exporting data to JSON: {output_json}")
            with open(output_json, "w", encoding="utf-8") as f:
                json.dump(regular_rows, f, indent=2, ensure_ascii=False)
            logger.info(f"✅ Successfully exported to {output_json}")
        except Exception as e:
            logger.error(f"Failed to export to JSON: {e}")

        logger.info("Export process completed")
        
    except Exception as e:
        logger.error(f"An error occurred in the main process: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
