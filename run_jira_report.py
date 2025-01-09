import re
import requests
import pandas as pd
import json
import os
import configparser
import smtplib
import argparse
import logging
from logging.handlers import SysLogHandler
from email.message import EmailMessage
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, GradientFill

# Parse arguments
parser = argparse.ArgumentParser(description="Generate a Jira report and send it as an attachment")
parser.add_argument("--email", help="Email address to send the report to")
parser.add_argument("--debug", action="store_true", help="Output log messages to console")
parser.add_argument("--plain", action="store_true", help="Outputs non-formatted Excel file")
args = parser.parse_args()

# Retrieve and build configurations
config = configparser.ConfigParser()
config.read("config.ini")

JIRA_BASE_URL = config["credentials"]["JIRA_BASE_URL"]
API_EMAIL = config["credentials"]["API_EMAIL"]
API_KEY = config["credentials"]["API_KEY"]
PROJECT_KEY = config["credentials"]["PROJECT_KEY"]

SMTP_SERVER = config["smtp"]["SMTP_SERVER"]
SMTP_PORT = int(config["smtp"]["SMTP_PORT"])
SMTP_DEBUG = int(config["smtp"].get("SMTP_DEBUG", 0))
FROM_EMAIL = config["smtp"]["FROM_EMAIL"]
REPLY_TO = config["smtp"]["REPLY_TO"]
SUBJECT = config["smtp"]["SUBJECT"]


# Setup logging
logfile = os.path.join(os.path.dirname(os.path.abspath(__file__)), "script.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)s[%(process)d]: %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(logfile),
    ]
)

logger = logging.getLogger()

if args.debug:
    console_handler = logging.Streamhandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(logging.Formatter("%(asctime)s %(name)s[%(process)d]: %(levelname)s %(message)s"))
    logger.addHandler(console_handler)
    logger.setLevel(logging.DEBUG)
    logger.info("Debugging mode enabled")

# JQL to fetch Epics and related issues
JQL_QUERY = f'issuetype = Epic AND project = {PROJECT_KEY} ORDER BY key ASC' 

def fetch_issues(jql_query):
    url = f"{JIRA_BASE_URL}/rest/api/3/search"
    headers = {
        "Accept": "application/json",
    }
    auth = (API_EMAIL, API_KEY)
    params = {
        "jql": jql_query,
        "maxResults": 100,  
        "fields": "key,summary,comment, status"
    }
    logger.info(f"Fetching issues with JQL: {jql_query}")
    try:
        response = requests.get(url, headers=headers, auth=auth, params=params)
        response.raise_for_status()
        logger.info("Successfully fetched issues from Jira")
        return response.json()
    except Exception as e:
        logger.error(f"Error fetching issues: {e}")
        raise

def fetch_most_recent_comment(issue_key):
    url = f"{JIRA_BASE_URL}/rest/api/3/issue/{issue_key}/comment"
    headers = {
        "Accept": "application/json",
    }
    auth = (API_EMAIL, API_KEY)
    logger.info(f"Fetching most recent comment for issue: {issue_key}")
    try:
        response = requests.get(url, headers=headers, auth=auth)
        response.raise_for_status()
        comments = response.json()["comments"]
        if comments:
            logger.info(f"Found {len(comments)} comments for issue {issue_key}")
            recent_comment = comments[-1]
            comment_body = parse_comment(recent_comment["body"])
            comment_date = recent_comment["created"]
            try:
                dt = datetime.strptime(comment_date, '%Y-%m-%dT%H:%M:%S.%f%z')
            except ValueError:
                dt = datetime.strptime(comment_date, '%Y-%m-%dT%H:%M:%S%z')
            formatted_comment_date = dt.strftime('%Y-%m-%d %H:%M:%S %z')
            
            return comment_body, formatted_comment_date
        logger.info(f"No comments found for issue {issue_key}")
        return " ", None
    except Exception as e:
        logger.error(f"Error fetching comments for {issue_key}: {e}")
        raise

def parse_comment(comment_json):
    def parse_block(block):
        """Recursively parse a block of JSON and return its text representation."""
        if block["type"] == "paragraph":
            # Parse the content of a paragraph
            paragraph_text = []
            for content in block.get("content", []):
                if content["type"] == "text":
                    paragraph_text.append(content["text"])
            return " ".join(paragraph_text)
        elif block["type"] == "listItem":
            # Parse a list item
            bullet_text = []
            for content in block.get("content", []):
                if content["type"] == "paragraph":
                    bullet_text.append(parse_block(content))
            return f"- {' '.join(bullet_text)}"
        elif block["type"] == "orderedList" or block["type"] == "bulletList":
            # Parse ordered or bullet lists
            list_text = []
            for item in block.get("content", []):
                if item["type"] == "listItem":
                    list_text.append(parse_block(item))
            return "\n".join(list_text)
        return ""  # Fallback for unhandled block types

    if isinstance(comment_json, dict) and comment_json.get("type") == "doc":
        text = []
        for block in comment_json.get("content", []):
            parsed_block = parse_block(block)
            if parsed_block:
                text.append(parsed_block)
        return "\n".join(text)
    return json.dumps(comment_json, indent=2)  # Return formatted JSON for unhandled structures


def format_excel_file(filename):
    wb = load_workbook(filename)
    ws = wb.active
    header_fill = PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')  # Light steel blue
    header_font = Font(bold=True, size=11, color='000000')
    header_border = Border(
        bottom=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000')
    )
    
    epic_header_pattern = re.compile(r'.+\(NOOPT-\d+\)$')
    
    done_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Pastel green
    sub_task_spacer_fill = PatternFill(start_color="DCDCDC", end_color="DCDCDC", fill_type="solid")  # Light Gray
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # White
    
    current_row = 1
    logger.info("Formatting rows")
    while current_row <= ws.max_row:
        first_cell = ws[f'A{current_row}']
        cell_value = str(first_cell.value) if first_cell.value else ""
        row_key = str(ws[f'B{current_row}'].value)
        logger.info(f'Cell value: {cell_value}')

        if epic_header_pattern.match(cell_value): 
            logger.info(f'Found header row: {current_row}')
            ws.merge_cells(f'A{current_row}:F{current_row}')
            
            header_cell = ws[f'A{current_row}']
            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = Alignment(horizontal='left', vertical='center')
            
        else:
            if ws[f'G{current_row}'].value is not None:
                indent_level = int(ws[f'G{current_row}'].value)
            else:
                indent_level = 1
                row_key = ""

            if ws[f'A{current_row}'].value == "§":
                indent_level = 3
                
            logger.debug(f'Indent level: {indent_level} Row: {current_row}')
            if indent_level == 0:  # Task
                logger.info(f'Task row: {current_row}')
                ws.merge_cells(f'A{current_row}:B{current_row}')
                task_cell = ws[f'A{current_row}']
                task_cell.value = row_key
                task_cell.alignment = Alignment(horizontal='left', vertical='center')
            elif indent_level == 1:  # Sub-Task
                logger.info(f'Sub-Task row: {current_row}')
                spacer_cell = ws[f'A{current_row}']
                spacer_cell.value = ""
                spacer_cell.fill = sub_task_spacer_fill
                sub_task_cell = ws[f'B{current_row}']
                sub_task_cell.value = row_key
                sub_task_cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                spacer_cell = ws[f'A{current_row}']
                spacer_cell.value = ""
                spacer_cell.fill = PatternFill(fill_type=None)
            
            for col in ['C', 'D', 'E', 'F']:
                cell = ws[f'{col}{current_row}']
                cell.alignment = Alignment(
                    wrap_text=True, 
                    vertical='top',
                    horizontal='left',
                    indent=indent_level
                )
                
                if ws[f'D{current_row}'].value and ws[f'D{current_row}'].value.strip() == "Done":
                    for col_to_fill in ['A', 'B', 'C', 'D', 'E', 'F']:
                        ws[f'{col_to_fill}{current_row}'].fill = done_fill
        current_row += 1
    
    ws.column_dimensions['A'].width = 2  # Spacer column
    ws.column_dimensions['B'].width = 15  # Key column
    ws.column_dimensions['C'].width = 40  # Summary column
    ws.column_dimensions['D'].width = 20  # Status column
    ws.column_dimensions['E'].width = 25  # Comment Date column
    ws.column_dimensions['F'].width = 100  # Comment column
    
    ws.delete_cols(7)
    
    wb.save(filename)

def send_email(recipient, filename):
    logger.info(f"Preparing to send email to {recipient} with attachment {filename}")
    try:
        msg = EmailMessage()
        msg["From"] = FROM_EMAIL
        msg["To"] = recipient
        msg["Reply-To"] = REPLY_TO
        msg["Subject"] = SUBJECT
        msg.set_content("Please see the attached Jira report.")
    
        with open(filename, "rb") as f:
            file_data = f.read()
            file_name = os.path.basename(filename)
            msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
    
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            if SMTP_DEBUG > 0:
                smtp_debug_handler = logging.StreamHandler()
                smtp_debug_handler.setLevel(logging.DEBUG)
                smtp_debug_handler.setFormatter(logging.Formatter("%(asctime)s %(name)s[%(process)d]: %(levelname)s %(message)s"))
                debug_logger = logging.getLogger("SMTP")
                debug_logger.setLevel(logging.DEBUG)
                debug_logger.addHandler(smtp_debug_handler)

                # Redirect smtplib debug output to the logger
                class SMTPLoggerAdapter:
                    def write(self, msg):
                        if msg.strip():  # Avoid blank lines
                            debug_logger.debug(msg.strip())
                    def flush(self):  # Required for compatibility with `sys.stderr`
                        pass

                server.set_debuglevel(SMTP_DEBUG)
                server._debug_output = SMTPLoggerAdapter()

            server.send_message(msg)
        logger.info(f"Email sent successfully to {recipient}")
    except Exception as e:
        logger.error(f"Failed to send email to {recipient}: {e}")
        raise

def main():
    logger.info("Script started")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logger.info(f"Working directory: {script_dir}")
    
    data = []
    print("Fetching Epics...", end="", flush=True)
    logger.info("Fetching Epics...")
    epics = fetch_issues(JQL_QUERY)
    
    for epic in epics["issues"]:
        epic_key = epic["key"]
        epic_summary = epic["fields"]["summary"]
        logger.info(f"Processing Epic: {epic_key} - {epic_summary}")
        print("*", end="", flush=True)
        child_issues = fetch_issues(f'"Epic Link" = {epic_key}')
        
        if child_issues["issues"]:
            for child in child_issues["issues"]:
                child_key = child["key"]
                child_summary = child["fields"]["summary"]
                child_status = child["fields"]["status"]["name"]
                
                most_recent_comment, comment_date = fetch_most_recent_comment(child_key)
                
                if child_status == "Done":
                    done_task = {
                        "Epic Key": epic_key,
                        "Epic Summary": epic_summary,
                        "Child Key": child_key,
                        "Child Summary": child_summary,
                        "Child Status": child_status,
                        "Comment Date": comment_date,
                        "Most Recent Comment": most_recent_comment,
                        "Type": "Task",
                        "Indent": 0,
                    }
                    data.append(done_task)
                    continue 

                data.append({
                    "Epic Key": epic_key,
                    "Epic Summary": epic_summary,
                    "Child Key": child_key,
                    "Child Summary": child_summary,
                    "Child Status": child_status,
                    "Comment Date": comment_date,
                    "Most Recent Comment": most_recent_comment,
                    "Type": "Task",
                    "Indent": 0,
                })
                
                # Subtask handling
                sub_tasks = fetch_issues(f'"parent" = {child_key}')
                for sub_task in sub_tasks["issues"]:
                    sub_key = sub_task["key"]
                    sub_summary = sub_task["fields"]["summary"]
                    sub_status = sub_task["fields"]["status"]["name"]
                    sub_parent = child_key
                    
                    sub_comment, sub_comment_date = fetch_most_recent_comment(sub_key)
                    
                    data.append({
                        "Epic Key": epic_key,
                        "Epic Summary": epic_summary,
                        "Child Key": sub_key,
                        "Child Summary": sub_summary,
                        "Child Status": sub_status,
                        "Comment Date": sub_comment_date,
                        "Most Recent Comment": sub_comment,
                        "Type": "Sub-Task",
                        "Parent": sub_parent,
                        "Indent": 1,
                    })
               
        else:
            data.append({
                "Epic Key": epic_key,
                "Epic Summary": epic_summary,
                "Child Key": "",
                "Child Summary": "No tasks created",
                "Comment Date": "",
                "Most Recent Comment": "",
            })
            
    # Format Excel output
    print("Generating Excel report...")
    logger.info("Generating Excel report")
    excel_data = []
    for epic_key, group in pd.DataFrame(data).groupby("Epic Key"):
        epic_summary = group.iloc[0]["Epic Summary"]
        header_row = [f"{epic_summary} ({epic_key})", "", "", "", "", ""]
        excel_data.append(header_row)
        
        status_order = {"In Progress": 1, "Waiting": 2, "For approval": 3, "New": 4, "Done": 5}
        
        child_issues = group[group["Type"] == "Task"].copy()
        sub_tasks = group[group["Type"] == "Sub-Task"].copy()
        logger.debug("Child Issues DataFrame:")
        logger.debug(child_issues)
        logger.debug("Sub-Tasks DataFrame:")
        logger.debug(sub_tasks)
        
        logger.info("Sorting rows")
        child_issues.sort_values(
            by=["Child Status", "Comment Date"],
            key=lambda col: (
                col.map(status_order).fillna(5) if col.name == "Child Status" else col
            ),
            ascending=[True, False],
            inplace=True
        )
        
        sub_tasks.sort_values(
            by=["Child Status", "Comment Date", "Parent"],
            key=lambda col: (
                col.map(status_order).fillna(5) if col.name == "Child Status" else col
            ),
            ascending=[True, False, True],
            inplace=True
        )
        
        sorted_rows =[]
        for _, child_row in child_issues.iterrows():
            sorted_rows.append(child_row.to_dict())
            child_key = child_row["Child Key"]
            sub_task_rows = sub_tasks[sub_tasks["Parent"] == child_key]
            sorted_rows.extend([row.to_dict() for _, row in sub_task_rows.iterrows()])
            
        group = pd.DataFrame(sorted_rows)
            
        for _, row in group.iterrows():
            excel_data.append([
                "",
                row["Child Key"],
                row["Child Summary"],
                row["Child Status"],
                row["Comment Date"],
                row["Most Recent Comment"],
                row.get("Indent", 0),
            ])
        
        excel_data.append(["§", "", "", "", "", "", ""]) #Spacer row
    
    # Save to Excel
    timestamp = datetime.utcnow().strftime("%Y-%m-%d_%H-%M-%S")
    filename = os.path.join(script_dir, f"jira_report_{timestamp}.xlsx")
    df = pd.DataFrame(excel_data)

    df.to_excel(filename, index=False, header=False)
    if not getattr(args, 'plain', False): 
        print("Applying formatting...")
        logger.info("Applying formatting...")
        format_excel_file(filename)

    print(f"Report saved as '{filename}'")
    logger.info(f"Report saved as '{filename}'")
    return filename

if __name__ == "__main__":
    report_filename = main()
    if args.email:
        recipient = args.email
        print(f"Sending report to {recipient}...")
        send_email(recipient, report_filename)

