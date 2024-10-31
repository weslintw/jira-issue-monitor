import requests
from requests.auth import HTTPBasicAuth
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import configparser

# Load configuration from ini file
config = configparser.ConfigParser()
config.read('config.ini')

# Jira API URL and authentication details
JIRA_URL = config['JIRA']['URL']
AUTH = HTTPBasicAuth(config['JIRA']['USERNAME'], config['JIRA']['API_TOKEN'])
HEADERS = {"Accept": "application/json"}

# Jira Query Language (JQL) query to fetch issues
JQL_QUERY = config['JIRA']['JQL_QUERY']
MAX_RESULTS = int(config['JIRA']['MAX_RESULTS'])

def fetch_issues():
    """Fetch all issues from Jira based on the JQL query."""
    start_at = 0
    all_issues = []
    while True:
        params = {'jql': JQL_QUERY, 'maxResults': MAX_RESULTS, 'startAt': start_at}
        response = requests.get(JIRA_URL, headers=HEADERS, auth=AUTH, params=params)
        response.raise_for_status()
        issues = response.json().get('issues', [])
        if not issues:
            break
        all_issues.extend(issues)
        start_at += MAX_RESULTS
    return all_issues

def fetch_comments(issue_key):
    """Fetch the last three comments for a given Jira issue."""
    comments_url = f"https://metainfra.atlassian.net/rest/api/3/issue/{issue_key}/comment"
    response = requests.get(comments_url, headers=HEADERS, auth=AUTH)
    response.raise_for_status()
    comments_data = response.json().get('comments', [])
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

    # 只取最後三個評論
    for comment in comments_data[-3:][::-1]:
        try:
            author = comment['updateAuthor']['displayName']
            update_time = comment['updated']
            comment_body = extract_text(comment['body']['content'])

            # 組合完整的評論內容
            full_comment = f"**{author} ({update_time}):**\n{comment_body}"
            comments_list.append(full_comment)
        except (KeyError, IndexError) as e:
            comments_list.append(f"Error parsing comment: {str(e)}")

    return comments_list  # 返回評論列表，而不是合併的字串

def create_excel(issues):
    """Create an Excel file with the fetched Jira issues and their details."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jira Issues"
    ws.append(["Jira Ticket ID", "Summary", "PIC", "Status", "Priority", "Update Time", "Comments"])
    wrap_alignment = Alignment(wrap_text=True)

    with ThreadPoolExecutor(max_workers=10) as executor:
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

            # 第一行
            start_row = ws.max_row + 1
            ws.append([issue_key, summary, assignee, status, priority, updated, comments[0] if comments else ""])

            # 設置超連結
            cell = ws.cell(row=start_row, column=1)
            cell.value = issue_key
            cell.hyperlink = f"https://metainfra.atlassian.net/browse/{issue_key}"
            cell.style = "Hyperlink"

            # 第一個評論設為藍色
            comment_cell = ws.cell(row=start_row, column=7)
            comment_cell.font = Font(color="0000FF")

            # 添加其餘評論
            for comment in comments[1:]:
                ws.append(["", "", "", "", "", "", comment])

            # 合併相同 issue 的儲存格
            end_row = ws.max_row
            for col in range(1, 7):  # 合併 A 到 F 欄
                if start_row != end_row:  # 只在有多行時才合併
                    ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)

            progress = (index / len(issues)) * 100
            print(f"Processed {index}/{len(issues)} issues ({progress:.2f}%)")

    format_excel(ws)
    save_excel(wb)

def format_excel(ws):
    """Format the Excel sheet with appropriate styles and widths."""
    wrap_alignment = Alignment(wrap_text=True, vertical='top')
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = wrap_alignment
            cell.font = Font(name='Calibri', color=cell.font.color, bold=cell.font.bold, italic=cell.font.italic, underline=cell.font.underline)

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        adjusted_width = (max_length + 2)
        ws.column_dimensions[col[0].column_letter].width = adjusted_width

    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['G'].width = 100

    for cell in ws[1]:
        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        cell.font = Font(bold=True, color=cell.font.color)

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

    # Apply hyperlink style to the first column (A) except the header
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.font = Font(color="0000FF", underline="single", name=cell.font.name, bold=cell.font.bold, italic=cell.font.italic)


def save_excel(wb):
    """Save the Excel workbook to a file."""
    current_time = datetime.now().strftime("%m-%d_%H.%M")
    file_name = f"jira_issues_{current_time}.xlsx"
    file_path = rf"C:\Users\Wes\Documents\PlatformIO\Projects\jira-ticket-monitor\{file_name}"
    wb.save(file_path)
    print(f"Jira issues have been written to {file_name}")

def main():
    """Main function to fetch Jira issues and create an Excel report."""
    print("Sending request to Jira...")
    try:
        issues = fetch_issues()
        print(f"Total issues to process: {len(issues)}")
        create_excel(issues)
    except requests.RequestException as e:
        print(f"Failed to fetch issues: {e}")

if __name__ == "__main__":
    main()
