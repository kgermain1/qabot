import streamlit as st
from openai import OpenAI
from docx import Document
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

client = OpenAI(api_key=st.secrets["openai"]["api_key"])

# Google Sheets Credentials
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1INz9LD7JUaZiIbY4uoGId0riIhkltlE6aroLKOAWtNo/edit#gid=0"
CREDENTIALS_FILE = "zenith-ww-65cc590712fd.json"  # Replace with your Google Sheets API credentials file

# Use the latest OpenAI GPT model
MODEL_NAME = "gpt-4"

# Initialize Google Sheets
def get_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(credentials)
    return client.open_by_url(GOOGLE_SHEET_URL)

def get_tab_names(sheet):
    """Fetch tab names from the Google Sheet."""
    return [worksheet.title for worksheet in sheet.worksheets()]

def get_rules(sheet, client, market):
    """Retrieve tone-of-voice rules for the selected client and market."""
    # Fetch rules from ALL CLIENTS tab
    all_clients_tab = sheet.worksheet("ALL CLIENTS")
    all_clients_rules = pd.DataFrame(all_clients_tab.get_all_records())
    all_clients_rules["Source Client"] = "ALL CLIENTS"

    # Fetch rules from the selected client’s tab
    client_tab = sheet.worksheet(client)
    client_rules = pd.DataFrame(client_tab.get_all_records())
    client_rules["Source Client"] = client

    # Filter client-specific rules for the selected market or "All"
    client_market_rules = client_rules[
        (client_rules["Market"] == market) | (client_rules["Market"] == "All")
    ]

    # Combine rules, preserving the source information
    combined_rules = pd.concat([all_clients_rules, client_market_rules], ignore_index=True)
    return combined_rules

def read_docx(file):
    """Reads the content of a Word document."""
    doc = Document(file)
    text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text])
    return text

def check_compliance(document_text, rules_df):
    """Checks compliance of the document against the rules."""
    # Use the rule name, rule content, and source for the check
    tone_rules = rules_df[["Rule Name", "Rules", "Market", "Source Client"]].to_dict("records")
    rules_list = "\n".join(
        [f"- Rule Name: {rule['Rule Name']} | Rule: {rule['Rules']} (Client: {rule['Source Client']}, Market: {rule['Market']})" for rule in tone_rules]
    )

    messages = [
        {"role": "system", "content": "You are an expert in compliance and tone-of-voice review."},
        {
            "role": "user",
            "content": f"""
Document Content:
{document_text}

Tone of Voice Rules:
{rules_list}

Analyze the document for compliance with the rules. Identify any violations and state which rule is being violated. For each violation, list:
- The exact rule name from the rules list (Column B).
- The client name: the source of the rule as specified in the rules list.
- The market name: the source market as specified in the rules list.

Structure the report as follows:
1. State whether the document is "Compliant" or "Non-Compliant".
2. If non-compliant, provide details of each violation, including the rule name, client name, and market.
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
        return response.choices[0].message.content
    except Exception as e:
        return f"An error occurred: {str(e)}"

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
        # Process the uploaded file
        document_text = read_docx(uploaded_file)

        # Retrieve tone rules
        rules_df = get_rules(sheet, selected_client, selected_market)

        with st.spinner("Checking compliance..."):
            # Check compliance using the model
            compliance_report = check_compliance(document_text, rules_df)

            if compliance_report.startswith("An error occurred"):
                st.error(compliance_report)
            else:
                st.success("Compliance check complete!")
                st.subheader("Compliance Report")
                st.text_area("Report", value=compliance_report, height=300, disabled=True)
    else:
        st.error("Please upload a document, select a client, and a market.")
