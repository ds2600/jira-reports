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
        return "No comments", None
    except Exception as e:
        logger.error(f"Error fetching comments for {issue_key}: {e}")
        raise

def parse_comment(comment_json):
    if isinstance(comment_json, dict) and comment_json.get("type") == "doc":
        text = []
        for block in comment_json.get("content", []):
            if block["type"] == "paragraph":
                for content in block.get("content", []):
                    if content["type"] == "text":
                        text.append(content["text"])
        return "\n".join(text)
    return comment_json  # Return raw JSON if not in expected format

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
    
    current_row = 1
    while current_row <= ws.max_row:
        first_cell = ws[f'A{current_row}']
        cell_value = str(first_cell.value) if first_cell.value else ""

        if epic_header_pattern.match(cell_value): 
            ws.merge_cells(f'A{current_row}:E{current_row}')
            
            header_cell = ws[f'A{current_row}']
            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = Alignment(horizontal='left', vertical='center')
            
        elif first_cell.value:
            for col in ['A', 'B', 'C', 'D', 'E']:
                cell = ws[f'{col}{current_row}']
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                
                if ws[f'C{current_row}'].value == "Done":
                    for col_to_fill in ['A', 'B', 'C', 'D', 'E']:
                        ws[f'{col_to_fill}{current_row}'].fill = done_fill
        current_row += 1
    
    ws.column_dimensions['A'].width = 15  # Key column
    ws.column_dimensions['B'].width = 40  # Summary column
    ws.column_dimensions['C'].width = 20  # Status column
    ws.column_dimensions['D'].width = 25  # Time column
    ws.column_dimensions['E'].width = 50  # Comment column
    
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
                data.append({
                    "Epic Key": epic_key,
                    "Epic Summary": epic_summary,
                    "Child Key": child_key,
                    "Child Summary": child_summary,
                    "Child Status": child_status,
                    "Comment Date": comment_date,
                    "Most Recent Comment": most_recent_comment,
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
    logger.info("Generating Excel report...")
    excel_data = []
    for epic_key, group in pd.DataFrame(data).groupby("Epic Key"):
        epic_summary = group.iloc[0]["Epic Summary"]
        header_row = [f"{epic_summary} ({epic_key})", "", "", "", ""]
        excel_data.append(header_row)
        
        status_order = {"In Progress": 1, "Waiting": 2, "For approval": 2, "New": 3, "Done": 4}
        
        group = group.sort_values(
            by=["Child Status", "Comment Date"],
            key=lambda col: col.map(status_order) if col.name == "Child Status" else col,
            ascending=[True, False]
        )
            
        for _, row in group.iterrows():
            excel_data.append([
                row["Child Key"],
                row["Child Summary"],
                row["Child Status"],
                row["Comment Date"],
                row["Most Recent Comment"],
            ])
        
        excel_data.append(["", "", "", "", ""])  # Add a blank row for spacing
    
    # Save to Excel
    timestamp = datetime.utcnow().strftime("%Y-%m-%d_%H-%M-%S")
    filename = os.path.join(script_dir, f"jira_report_{timestamp}.xlsx")
    df = pd.DataFrame(excel_data)
    df.to_excel(filename, index=False, header=False)

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

