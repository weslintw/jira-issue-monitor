import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import configparser
import win32com.client as win32


# Load configuration from ini file
config = configparser.ConfigParser()
config.read('config.ini')

# Jira API URL and authentication details
JIRA_URL = config['JIRA']['URL']
AUTH = HTTPBasicAuth(config['JIRA']['USERNAME'], config['JIRA']['API_TOKEN'])
HEADERS = {"Accept": "application/json"}

# Jira Query Language (JQL) query to fetch issues
JQL_QUERY = config['QUERY']['JQL_QUERY']


MAX_RESULTS = int(config['SETTINGS']['MAX_RESULTS'])
# Number of days to highlight recent comments
HIGHLIGHT_DAYS = int(config['SETTINGS']['HIGHLIGHT_DAYS'])

# Number of recent comments to include in the Excel report
RECENT_COMMENTS_COUNT = int(config['SETTINGS']['RECENT_COMMENTS_COUNT'])

# File name prefix for the Excel report
FILE_NAME_PREFIX = config['SETTINGS']['FILE_NAME_PREFIX']
# 获取保存目录
save_directory = config.get('Paths', 'save_directory')

def fetch_issues(jql_query):
    """Fetch all issues from Jira based on the JQL query."""
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
    return all_issues

def fetch_comments(issue_key):
    """Fetch the last few comments for a given Jira issue."""
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
            update_time = comment['updated']
            comment_body = extract_text(comment['body']['content'])

            # Combine the full comment content
            # Convert update time to local timezone
            update_time = datetime.strptime(update_time, '%Y-%m-%dT%H:%M:%S.%f%z')
            local_update_time = update_time.astimezone().strftime('%Y-%m-%d %H:%M:%S')
            full_comment = f"**{author} ({local_update_time}):**\n{comment_body}"
            comments_list.append(full_comment)
        except (KeyError, IndexError) as e:
            comments_list.append(f"Error parsing comment: {str(e)}")

    return comments_list  # Return the list of comments, not a combined string

def extract_labels(issue, prefix):
    """Extract labels from an issue based on a given prefix."""
    labels = [label[len(prefix):] for label in issue['fields']['labels'] if label.startswith(prefix)]
    return ','.join(labels)

def create_excel(queries):
    """Create an Excel file with the fetched Jira issues and their details."""
    excel = win32.Dispatch('Excel.Application')
    wb = excel.Workbooks.Add()
    
    for sheet_name, jql_query in queries.items():
        ws = wb.Worksheets.Add()
        ws.Name = sheet_name
        
        # Insert the JQL query from config.ini into the first row
        ws.Cells(1, 1).Value = f"JQL Query: {jql_query}"
        ws.Cells(1, 1).Interior.Color = 65535
        
        # Merge the cells for the JQL query row
        ws.Range(ws.Cells(1, 1), ws.Cells(1, 9)).Merge()
        
        # Add the header row for the issues
        headers = ["Jira Ticket ID", "Summary", "PIC", "Status", "Priority", "Update Time", "Sensor Issue Category", "Gerrit ID", "Comments"]
        for col_num, header in enumerate(headers, 1):
            ws.Cells(2, col_num).Value = header
            ws.Cells(2, col_num).Interior.Color = 65535
            ws.Cells(2, col_num).Font.Bold = True

        issues = fetch_issues(jql_query)
        with ThreadPoolExecutor(max_workers=20) as executor:
            future_to_issue = {executor.submit(fetch_comments, issue['key']): issue for issue in issues}
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

                # 第一行
                start_row = ws.UsedRange.Rows.Count + 1
                ws.Cells(start_row, 1).Value = issue_key
                ws.Cells(start_row, 2).Value = summary
                ws.Cells(start_row, 3).Value = assignee
                ws.Cells(start_row, 4).Value = status
                ws.Cells(start_row, 5).Value = priority
                ws.Cells(start_row, 6).Value = local_update_time
                ws.Cells(start_row, 7).Value = sensor_issue_category
                ws.Cells(start_row, 8).Value = gerrit_id

                # Combine all comments into one cell
                combined_comments = "\n\n".join(comments)
                ws.Cells(start_row, 9).Value = combined_comments

                # Set hyperlink
                ws.Hyperlinks.Add(Anchor=ws.Cells(start_row, 1), Address=f"https://metainfra.atlassian.net/browse/{issue_key}", TextToDisplay=issue_key)
                # Bold the author and timestamp in comments
                for comment in comments:
                    if '**' in comment:
                        start = combined_comments.find(comment)
                        end = start + len(comment)
                        bold_start = comment.find('**') + 2
                        bold_end = comment.find('**', bold_start)
                        if bold_end != -1:
                            ws.Cells(start_row, 9).GetCharacters(Start=start + bold_start + 1, Length=bold_end - bold_start).Font.Bold = True
                # Highlight recent comments
                for comment in comments:
                    if '(' in comment and ')' in comment:
                        comment_time_str = comment.split('(')[1].split(')')[0]
                        try:
                            comment_time = datetime.strptime(comment_time_str, '%Y-%m-%d %H:%M:%S')
                            if (datetime.now(comment_time.tzinfo) - comment_time).days <= HIGHLIGHT_DAYS:
                                start = combined_comments.find(comment)
                                end = start + len(comment)
                                ws.Cells(start_row, 9).GetCharacters(Start=start + 1, Length=len(comment)).Font.Color = 16711680
                        except (ValueError, IndexError):
                            pass

                progress = (index / len(issues)) * 100
                print(f"Processed {index}/{len(issues)} issues ({progress:.2f}%)")

        format_excel(ws)
    
    save_excel(wb, excel)

def format_excel(ws):
    """Format the Excel sheet with appropriate styles and widths."""
    
    # 設置單元格對齊方式為自動換行和頂部對齊
    ws.Cells.WrapText = True
    ws.Cells.VerticalAlignment = win32.constants.xlTop

    # 調整每一欄的寬度
    for col in range(1, 10):
        ws.Columns(col).AutoFit()

    # 設置特定欄的寬度
    # ws.Columns(1).ColumnWidth = 14.86
    ws.Columns(2).ColumnWidth = 50
    # ws.Columns(7).ColumnWidth = 30
    # ws.Columns(8).ColumnWidth = 30
    ws.Columns(9).ColumnWidth = 100

    # 設置單元格邊框
    for row in range(1, ws.UsedRange.Rows.Count + 1):
        for col in range(1, 10):
            cell = ws.Cells(row, col)
            cell.Borders.LineStyle = win32.constants.xlContinuous
            cell.Borders.Weight = win32.constants.xlThin
            # 設置所有單元格的字體為 Calibri
            ws.Cells.Font.Name = 'Calibri'

def save_excel(wb, excel):
    """Save the Excel workbook to a file."""
    current_time = datetime.now().strftime("%m-%d_%H.%M")
    file_name = f"{FILE_NAME_PREFIX}_jira_issues_{current_time}.xlsx"
    file_path = rf"{save_directory}\{file_name}"
    wb.SaveAs(file_path)
    wb.Close()
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
