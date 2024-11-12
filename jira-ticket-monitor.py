import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import configparser
import win32com.client as win32
import time
import os


# Load configuration from ini file
config = configparser.ConfigParser()
config.read('config.ini')

# Jira API URL and authentication details
JIRA_URL = config['JIRA']['URL']
AUTH = HTTPBasicAuth(config['JIRA']['USERNAME'], config['JIRA']['API_TOKEN'])
HEADERS = {"Accept": "application/json"}

MAX_RESULTS = int(config['SETTINGS']['MAX_RESULTS'])
# Number of days to highlight recent comments
HIGHLIGHT_DAYS = int(config['SETTINGS']['HIGHLIGHT_DAYS'])

# Number of recent comments to include in the Excel report
RECENT_COMMENTS_COUNT = int(config['SETTINGS']['RECENT_COMMENTS_COUNT'])

# File name prefix for the Excel report
FILE_NAME_PREFIX = config['SETTINGS']['FILE_NAME_PREFIX']
# File name postfix for the Excel report
FILE_NAME_POSTFIX = config['SETTINGS']['FILE_NAME_POSTFIX']
# 获取保存目录
save_directory = config.get('Paths', 'save_directory')

# Cache for storing comments of issues
comments_cache = {}

# Debug switch for timing
DEBUG_TIMING = False

def fetch_issues(jql_query):
    """Fetch all issues from Jira based on the JQL query."""
    if DEBUG_TIMING:
        start_time = time.time()
    start_at = 0
    all_issues = []
    while True:
        params = {'jql': jql_query, 'maxResults': MAX_RESULTS, 'startAt': start_at}
        response = requests.get(JIRA_URL, headers=HEADERS, auth=AUTH, params=params)
        response.raise_for_status()
        issues = response.json().get('issues', [])
        if not issues:
            break
        all_issues.extend(issues)
        start_at += MAX_RESULTS
    if DEBUG_TIMING:
        end_time = time.time()
        print(f"fetch_issues took {end_time - start_time:.2f} seconds")
    return all_issues

def fetch_comments(issue_key):
    """Fetch the last few comments for a given Jira issue."""
    if DEBUG_TIMING:
        start_time = time.time()
    if issue_key in comments_cache:
        return comments_cache[issue_key]

    comments_url = f"https://metainfra.atlassian.net/rest/api/3/issue/{issue_key}"
    response = requests.get(comments_url, headers=HEADERS, auth=AUTH)
    response.raise_for_status()
    comments_data = response.json().get('fields', {}).get('comment', {}).get('comments', [])
    comments_list = []

    def extract_text(content):
        """Recursively extract text from the content."""
        if isinstance(content, list):
            return ''.join(extract_text(item) for item in content)

        if not isinstance(content, dict):
            return str(content)

        if content['type'] == 'text':
            return content['text']
        elif content['type'] == 'mention':
            return content['attrs']['text']
        elif content['type'] == 'hardBreak':
            return "\n"
        elif 'content' in content:
            return ''.join(extract_text(item) for item in content['content'])
        return ""

    # Only take the last few comments based on RECENT_COMMENTS_COUNT
    for comment in comments_data[-RECENT_COMMENTS_COUNT:][::-1]:
        try:
            author = comment['updateAuthor']['displayName']
            created_time = comment['created']
            comment_body = extract_text(comment['body']['content'])

            # Combine the full comment content
            # Convert created time to local timezone
            created_time = datetime.strptime(created_time, '%Y-%m-%dT%H:%M:%S.%f%z')
            local_created_time = created_time.astimezone().strftime('%Y-%m-%d %H:%M:%S')
            full_comment = f"**[{local_created_time}, {author}]**\n{comment_body}"
            comments_list.append(full_comment)
        except (KeyError, IndexError) as e:
            comments_list.append(f"Error parsing comment: {str(e)}")

    comments_cache[issue_key] = comments_list  # Cache the comments
    if DEBUG_TIMING:
        end_time = time.time()
        print(f"fetch_comments for {issue_key} took {end_time - start_time:.2f} seconds")
    return comments_list  # Return the list of comments, not a combined string

def extract_labels(issue, prefix):
    """Extract labels from an issue based on a given prefix."""
    labels = [label[len(prefix):] for label in issue['fields']['labels'] if label.startswith(prefix)]
    return ','.join(labels)

def create_excel(queries):
    """Create an Excel file with the fetched Jira issues and their details."""
    if DEBUG_TIMING:
        start_time = time.time()
    #excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel = win32.Dispatch('Excel.Application')
    
    current_time = datetime.now().strftime("%m-%d_%H.%M")
    file_name = f"{FILE_NAME_PREFIX}_jira_issues_{FILE_NAME_POSTFIX}.xlsx"
    file_path = rf"{save_directory}\{file_name}"
    
    if os.path.exists(file_path):
        wb = excel.Workbooks.Open(file_path)
    else:
        wb = excel.Workbooks.Add()
    
    for sheet_name, jql_query in queries.items():
        if DEBUG_TIMING:
            sheet_start_time = time.time()
        
        # Check if the worksheet already exists
        try:
            ws = wb.Worksheets(sheet_name)
        except win32.com_error:
            ws = wb.Worksheets.Add()
            ws.Name = sheet_name
            # Add the header row for the issues
            headers = ["Jira Ticket ID", "Summary", "PIC", "Status", "Priority", "Update Time", "Sensor Issue Category", "Gerrit ID", "Comments", "Remark"]
            for col_num, header in enumerate(headers, 1):
                ws.Cells(2, col_num).Value = header
                ws.Cells(2, col_num).Interior.Color = 65535
                ws.Cells(2, col_num).Font.Bold = True

        if DEBUG_TIMING:
            fetch_issues_start_time = time.time()
        issues = fetch_issues(jql_query)
        if DEBUG_TIMING:
            fetch_issues_end_time = time.time()
            print(f"fetch_issues for {sheet_name} took {fetch_issues_end_time - fetch_issues_start_time:.2f} seconds")

        if DEBUG_TIMING:
            fetch_comments_start_time = time.time()
        rows = []
        with ThreadPoolExecutor(max_workers=20) as executor:
            future_to_issue = {executor.submit(fetch_comments, issue['key']): issue for issue in issues}
            print(f"Processing JQL Query: {sheet_name}")
            for index, future in enumerate(as_completed(future_to_issue), start=1):
                issue = future_to_issue[future]
                issue_key = issue['key']
                summary = issue['fields']['summary']
                assignee = issue['fields']['assignee']['displayName'] if issue['fields']['assignee'] else 'Unassigned'
                status = issue['fields']['status']['name']
                priority = issue['fields']['priority']['name'] if issue['fields']['priority'] else 'None'
                updated = issue['fields']['updated']
                comments = future.result()  # 這裡獲取的是評論列表

                # Extract labels
                sensor_issue_category = extract_labels(issue, 'issue-category:')
                gerrit_id = extract_labels(issue, 'gerrit:')

                # Convert update time to local timezone
                update_time = datetime.strptime(updated, '%Y-%m-%dT%H:%M:%S.%f%z')
                local_update_time = update_time.astimezone().strftime('%Y-%m-%d %H:%M:%S')

                # Combine all comments into one cell
                combined_comments = "\n\n".join(comments)
                rows.append([issue_key, summary, assignee, status, priority, local_update_time, sensor_issue_category, gerrit_id, combined_comments, ""])

                progress = (index / len(issues)) * 100
                print(f"Processed {index}/{len(issues)} issues ({progress:.2f}%)")

        if DEBUG_TIMING:
            fetch_comments_end_time = time.time()
            print(f"fetch_comments for {sheet_name} took {fetch_comments_end_time - fetch_comments_start_time:.2f} seconds")

        # Update existing rows or add new rows
        last_row = 2  # Start after the header row
        if ws.UsedRange is not None:
            last_row = ws.UsedRange.Rows.Count
        for row in rows:
            issue_key = row[0]
            found = False
            for r in range(3, last_row + 1):
                if ws.Cells(r, 1).Value == issue_key:
                    for col in range(2, 10):
                        ws.Cells(r, col).Value = str(row[col - 1])
                    found = True
                    break
            if not found:
                last_row += 1
                for col in range(1, 10):
                    ws.Cells(last_row, col).Value = row[col - 1]

        # Set hyperlinks and format comments
        for row_num in range(3, last_row + 1):
            ws.Hyperlinks.Add(Anchor=ws.Cells(row_num, 1), Address=f"https://metainfra.atlassian.net/browse/{ws.Cells(row_num, 1).Value}", TextToDisplay=ws.Cells(row_num, 1).Value)
            cell_value = ws.Cells(row_num, 9).Value
            if cell_value:
                for comment in cell_value.split("\n\n"):
                    if '**' in comment:
                        start = cell_value.find(comment)
                        bold_start = comment.find('**') + 2
                        bold_end = comment.find('**', bold_start)
                        if bold_end != -1:
                            ws.Cells(row_num, 9).GetCharacters(Start=start + bold_start + 1, Length=bold_end - bold_start).Font.Bold = True
                    if "**[" in comment and "]**" in comment:
                        comment_time_str = comment.split("**[")[1].split(", ")[0]
                        try:
                            comment_time = datetime.strptime(comment_time_str, '%Y-%m-%d %H:%M:%S')
                            if (datetime.now(comment_time.tzinfo) - comment_time).days <= HIGHLIGHT_DAYS:
                                start = cell_value.find(comment)
                                ws.Cells(row_num, 9).GetCharacters(Start=start + 1, Length=len(comment)).Font.Color = 16711680
                                if DEBUG_TIMING:
                                    print(f"Highlighted comment: {comment}")
                            else:
                                if DEBUG_TIMING:
                                    print(f"Comment not highlighted (older than {HIGHLIGHT_DAYS} days): {comment}")
                        except (ValueError, IndexError) as e:
                            if DEBUG_TIMING:
                                print(f"Error parsing comment time: {e}")
                                print(f"Comment: {comment}")

        format_excel(ws)
        if DEBUG_TIMING:
            sheet_end_time = time.time()
            print(f"Processing sheet {sheet_name} took {sheet_end_time - sheet_start_time:.2f} seconds")
    
    save_excel(wb, excel)
    if DEBUG_TIMING:
        end_time = time.time()
        print(f"create_excel took {end_time - start_time:.2f} seconds")

def format_excel(ws):
    """Format the Excel sheet with appropriate styles and widths."""
    
    # 設置單元格對齊方式為自動換行和頂部對齊
    ws.Cells.WrapText = True
    ws.Cells.VerticalAlignment = win32.constants.xlTop

    # 調整每一欄的寬度
    ws.Columns.AutoFit()

    # 設置特定欄的寬度
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 50
    ws.Columns(4).ColumnWidth = 15
    ws.Columns(7).ColumnWidth = 20
    ws.Columns(8).ColumnWidth = 8
    ws.Columns(9).ColumnWidth = 100
    ws.Columns(10).ColumnWidth = 50

    # 設置單元格邊框和字體
    used_range = ws.UsedRange
    used_range.Borders.LineStyle = win32.constants.xlContinuous
    used_range.Borders.Weight = win32.constants.xlThin
    used_range.Font.Name = 'Calibri'

def save_excel(wb, excel):
    """Save the Excel workbook to a file."""
    current_time = datetime.now().strftime("%m-%d_%H.%M")
    file_name = f"{FILE_NAME_PREFIX}_jira_issues_{FILE_NAME_POSTFIX}.xlsx"
    file_path = rf"{save_directory}\{file_name}"
    
    if os.path.exists(file_path):
        existing_wb = excel.Workbooks.Open(file_path)
        for sheet in wb.Worksheets:
            sheet.Copy(Before=existing_wb.Sheets(1))
        existing_wb.Save()
        existing_wb.Close(SaveChanges=True)
    else:
        wb.SaveAs(file_path)
    
    wb.Close(SaveChanges=True)
    excel.Quit()
    print(f"Jira issues have been written to {file_name}")

def main():
    """Main function to fetch Jira issues and create an Excel report."""
    print("Sending request to Jira...")
    try:
        queries = {key: value for key, value in config['QUERY'].items()}
        create_excel(queries)
    except requests.RequestException as e:
        print(f"Failed to fetch issues: {e}")

if __name__ == "__main__":
    main()
