import streamlit as st
import json
from openai import OpenAI
from docx import Document
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

client = OpenAI(api_key=st.secrets["openai"]["api_key"])

# Get the credentials JSON from the secrets manager
credentials_info = json.loads(st.secrets["google"]["credentials_json"])

# Google Sheets Credentials
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1INz9LD7JUaZiIbY4uoGId0riIhkltlE6aroLKOAWtNo/edit#gid=0"

# Use the latest OpenAI GPT model
MODEL_NAME = "gpt-4"

# Initialize Google Sheets
def get_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
    client = gspread.authorize(creds)
    return client.open_by_url(GOOGLE_SHEET_URL)

def get_tab_names(sheet):
    """Fetch tab names from the Google Sheet."""
    return [worksheet.title for worksheet in sheet.worksheets()]

def get_rules(sheet, client, market):
    """Retrieve tone-of-voice rules for the selected client and market."""
    all_clients_tab = sheet.worksheet("ALL CLIENTS")
    all_clients_rules = pd.DataFrame(all_clients_tab.get_all_records())
    all_clients_rules["Source Client"] = "ALL CLIENTS"

    client_tab = sheet.worksheet(client)
    client_rules = pd.DataFrame(client_tab.get_all_records())
    client_rules["Source Client"] = client

    client_market_rules = client_rules[
        (client_rules["Market"] == market) | (client_rules["Market"] == "All")
    ]

    combined_rules = pd.concat([all_clients_rules, client_market_rules], ignore_index=True)
    return combined_rules

def read_docx(file):
    """Reads the content of a Word document."""
    doc = Document(file)
    text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text])
    return text

def check_compliance(document_text, rules_df):
    """Checks compliance of the document against the rules in chunks of 20."""
    rules_list = rules_df.to_dict("records")
    chunked_rules = [rules_list[i:i + 20] for i in range(0, len(rules_list), 20)]
    all_reports = []
    violation_number = 1

    for chunk in chunked_rules:
        rules_text = "\n".join(
            [f"- Rule Name: {rule['Rule Name']} | Rule: {rule['Rules']} (Client: {rule['Source Client']}, Market: {rule['Market']})"
             for rule in chunk]
        )

        messages = [
            {"role": "system", "content": "You are an expert in compliance and tone-of-voice review."},
            {
                "role": "user",
                "content": f"""
Document Content:
{document_text}

Tone of Voice Rules:
{rules_text}

Analyze the document for compliance with the rules. Identify any violations and state which rule is being violated. For each violation, list:
1. Rule name (number it continuously across chunks).
2. Client name.
3. Market name.
4. Details of the violation.
""",
            },
        ]

        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=messages,
                max_tokens=2000,
                temperature=0.5,
            )
            report = response.choices[0].message.content
            # Format and extract violations with continuous numbering
            violations = []
            for line in report.split("\n\n"):
                if line.strip():
                    formatted_violation = f"{violation_number}. {line.strip()}"
                    violations.append(formatted_violation)
                    violation_number += 1
            all_reports.extend(violations)
        except Exception as e:
            return f"An error occurred: {str(e)}"

    # Combine all reports into a single response
    compliance_status = "Compliant" if not all_reports else "Non-Compliant"
    combined_report = f"Document Status: {compliance_status}\n\n" + "\n\n".join(all_reports)
    return combined_report

# Streamlit App
st.title("Welcome to QAbot")

st.sidebar.title("Instructions")
st.sidebar.write("""
1. Select your client.
2. Select your market.
3. Upload a Word document (.docx).
4. Click the "Check Compliance" button to evaluate.
""")

# Google Sheets Initialization
try:
    sheet = get_google_sheets()
    tab_names = get_tab_names(sheet)
except Exception as e:
    st.error("Error connecting to Google Sheets. Please check your credentials.")
    st.stop()

# Dropdowns for client and market
selected_client = st.selectbox("Select a Client", [tab for tab in tab_names if tab != "ALL CLIENTS"])
if selected_client:
    market_tab = pd.DataFrame(sheet.worksheet(selected_client).get_all_records())
    available_markets = market_tab["Market"].unique().tolist()
    selected_market = st.selectbox("Select a Market", available_markets)

# File Upload
uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

if st.button("Check Compliance"):
    if uploaded_file and selected_client and selected_market:
        document_text = read_docx(uploaded_file)
        rules_df = get_rules(sheet, selected_client, selected_market)

        with st.spinner("Checking compliance..."):
            compliance_report = check_compliance(document_text, rules_df)
            if compliance_report.startswith("An error occurred"):
                st.error(compliance_report)
            else:
                st.success("Compliance check complete!")
                st.subheader("Compliance Report")
                st.text_area("Report", value=compliance_report, height=400, disabled=True)
    else:
        st.error("Please upload a document, select a client, and a market.")
