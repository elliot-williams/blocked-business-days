import streamlit as st
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import numpy as np
from datetime import datetime, timezone
from io import BytesIO

st.title("JIRA Blocked Issues Report")

JIRA_URL = "https://maersk-tools.atlassian.net"

SEARCH_URL = f"{JIRA_URL}/rest/api/3/search"
HEADERS = {"Accept": "application/json"}
MAX_RESULTS = 50


def get_issues(jira_email, jira_token, team_name):
    jql_dict = {
        "Reliance": '''
        "Team[Team]" in (92aa14a1-a594-471e-9b9f-162d0d038010-554)
        AND issuetype in (Story, Support)
        AND status in ("Blocked Internal", "Blocked External")
        ORDER BY created ASC
        ''',
        "Abbey Road": '''
        "Team[Team]" in (abbey-road-team-id)
        AND issuetype in (Story, Support)
        AND status in ("Blocked Internal", "Blocked External")
        ORDER BY created ASC
        ''',
        "Team Tigers": '''
        "Team[Team]" in (team-tigers-team-id)
        AND issuetype in (Story, Support)
        AND status in ("Blocked Internal", "Blocked External")
        ORDER BY created ASC
        ''',
        "TbM Ocean": '''
        "Team[Team]" in (
            92aa14a1-a594-471e-9b9f-162d0d038010-554,
            92aa14a1-a594-471e-9b9f-162d0d038010-298,
            b4d52324-fe3a-451f-ab59-89efbbbcd2ee
        )
        AND issuetype in (Story, Support)
        AND status in ("Blocked Internal", "Blocked External")
        ORDER BY created ASC
        '''
    }
    JQL = jql_dict.get(team_name, jql_dict["Reliance"])
    issues = []
    start_at = 0
    while True:
        params = {
            "jql": JQL,
            "startAt": start_at,
            "maxResults": MAX_RESULTS,
            "fields": "summary,assignee,status,customfield_12220,customfield_12221,created",
            "expand": "changelog"
        }
        response = requests.get(
            SEARCH_URL,
            headers=HEADERS,
            params=params,
            auth=HTTPBasicAuth(jira_email, jira_token)
        )
        response.raise_for_status()
        data = response.json()
        issues.extend(data.get("issues", []))
        if start_at + MAX_RESULTS >= data.get("total", 0):
            break
        start_at += MAX_RESULTS
    return issues


def calculate_days_in_blocked(issue):
    key = issue.get("key")
    fields = issue.get("fields", {})
    summary = fields.get("summary", "")
    assignee = fields.get("assignee", {}).get("displayName", "Unassigned")
    status = fields.get("status", {}).get("name", "")
    function = (fields.get("customfield_12220") or {}).get("value", "")
    team = (fields.get("customfield_12221") or {}).get("value", "")
    created = fields.get("created", "")
    changelog = issue.get("changelog", {}).get("histories", [])

    # Sort changelog by time
    changelog.sort(key=lambda x: x["created"])
    last_blocked_entry = None
    for entry in changelog:
        for item in entry.get("items", []):
            if item.get("field") == "status" and item.get("toString") in ["Blocked Internal", "Blocked External"]:
                last_blocked_entry = entry["created"]

    # Calculate business days in current blocked status
    if last_blocked_entry:
        start = pd.to_datetime(last_blocked_entry).date()
        end = datetime.now(timezone.utc).date()
        blocked_days = np.busday_count(start, end)
    else:
        blocked_days = None

    return {
        "Key": key,
        "Summary": summary,
        "Assignee": assignee,
        "Status": status,
        "Function": function,
        "Team": team,
        "Created": created,
        "Business Days in Blocked": blocked_days,
        "Link": f"{JIRA_URL}/browse/{key}"
    }


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df_to_save = df.copy()
    # Replace 'Key' column with Excel HYPERLINK formula
    df_to_save["Key"] = df_to_save.apply(
        lambda row: f'=HYPERLINK("{row["Link"]}", "{row["Key"]}")', axis=1
    )
    df_to_save.drop(columns=["Link"], inplace=True)
    df_to_save.to_excel(writer, index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


with st.form("credentials_form"):
    team_name = st.selectbox("Select Your Team", ["Reliance", "Abbey Road", "Team Tigers", "TbM Ocean"])
    jira_email = st.text_input("JIRA Email")
    jira_token = st.text_input("JIRA API Token", type="password")
    submitted = st.form_submit_button("Generate Report")

if submitted:
    if not jira_email or not jira_token:
        st.error("Please enter both JIRA Email and API Token.")
    else:
        try:
            with st.spinner("Fetching issues..."):
                issues = get_issues(jira_email, jira_token, team_name)
                blocked_issues = [calculate_days_in_blocked(issue) for issue in issues]
                df = pd.DataFrame(blocked_issues)
                # Show dataframe without the Link column
                st.dataframe(df.drop(columns=["Link"]))
                excel_data = to_excel(df)
                st.download_button(
                    label="Download Excel Report",
                    data=excel_data,
                    file_name="blocked_issues.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except requests.exceptions.HTTPError as e:
            st.error(f"HTTP error: {e}")
        except Exception as e:
            st.error(f"An error occurred: {e}")