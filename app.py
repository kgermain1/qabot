import streamlit as st
from openai import OpenAI
from docx import Document
import json
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Constants
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1QO6tjLTCTbM_O57wZk08d-A0UYcbGUvG90B7An9ibYQ/edit?gid=2007004764#gid=2007004764"
MODEL_NAME = "gpt-4"

# OpenAI client
openai_client = OpenAI(api_key=st.secrets["openai"]["api_key"])

# Load Google credentials
credentials_dict = json.loads(st.secrets["google"]["credentials_json"])

# Google Sheets Functions
def get_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gs_client = gspread.authorize(creds)
    return gs_client.open_by_url(GOOGLE_SHEET_URL)

@st.cache_data(ttl=300)
def get_tab_names(_sheet):
    return [worksheet.title for worksheet in _sheet.worksheets()]

def get_rules(sheet, client, market):
    """Retrieve rules for the selected client and market (excluding ALL CLIENTS)."""
    client_tab = sheet.worksheet(client)
    client_rules = pd.DataFrame(client_tab.get_all_records())

    filtered_client_rules = client_rules[
        (client_rules["Market"] == market) | (client_rules["Market"] == "All")
    ]

    if "Rule" not in filtered_client_rules or "Rule Name" not in filtered_client_rules:
        raise ValueError("The required 'Rule' or 'Rule Name' columns are missing in the Google Sheet.")

    return filtered_client_rules

def read_docx(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs if p.text])

def build_prompt_for_rule(document_text, rule_text, rule_name):
    return [
        {"role": "system", "content": "You are an expert in compliance and tone-of-voice review."},
        {
            "role": "user",
            "content": f"""
Document Content:
{document_text}

Rule:
{rule_text}

Analyze the document for compliance with this rule only. If non-compliant, provide a brief explanation referencing exactly the Rule Name: {rule_name}.

Format:
- If compliant, reply: "Compliant"
- If not compliant, reply only with the explanation text.

Do not mention any rules other than this one.
"""
        },
    ]

def check_rule_compliance(document_text, rule_text, rule_name):
    messages = build_prompt_for_rule(document_text, rule_text, rule_name)
    try:
        response = openai_client.chat.completions.create(
            model=MODEL_NAME,
            messages=messages,
            max_tokens=500,
            temperature=0.5,
        )
        answer = response.choices[0].message.content.strip()
        return answer
    except Exception as e:
        return f"Error: {str(e)}"

# Streamlit UI
st.title("Welcome to QAbot")

st.sidebar.title("Instructions")
st.sidebar.write("""
1. Select your client.
2. Select your market.
3. Upload a Word document (.docx).
4. Click the "Check Compliance" button to evaluate.
""")

# Load Google Sheet
try:
    sheet = get_google_sheets()
    tab_names = get_tab_names(sheet)
except Exception as e:
    st.error(f"Error connecting to Google Sheets. Please check your credentials.\n{str(e)}")
    st.stop()

# Client & Market Selection
selected_client = st.selectbox("Select a Client", tab_names)

selected_market = None
if selected_client:
    market_tab = pd.DataFrame(sheet.worksheet(selected_client).get_all_records())
    available_markets = market_tab["Market"].unique().tolist()
    selected_market = st.selectbox("Select a Market", available_markets)

# Document Upload
uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

if st.button("Check Compliance"):
    if not selected_client:
        st.error("Please select a client.")
    elif not selected_market:
        st.error("Please select a market.")
    elif not uploaded_file:
        st.error("Please upload a document.")
    else:
        document_text = read_docx(uploaded_file)
        rules_df = get_rules(sheet, selected_client, selected_market)

        violations = []

        progress_bar = st.progress(0)
        total_rules = len(rules_df)

        with st.spinner("Checking compliance..."):
            for idx, (_, row) in enumerate(rules_df.iterrows()):
                rule = row["Rule"]
                rule_name = row["Rule Name"]

                result = check_rule_compliance(document_text, rule, rule_name)

                if result.lower() != "compliant":
                    # Append violation details
                    violations.append({
                        "Rule": rule,
                        "Rule Name": rule_name,
                        "Explanation": result,
                    })

                progress_bar.progress((idx + 1) / total_rules)

        progress_bar.empty()

        if not violations:
            st.info("No violations found. Document is compliant.")
        else:
            st.subheader("Compliance Report")
            df_violations = pd.DataFrame(violations)
        
            # Keep only the 'Rule' and 'Explanation' columns for display
            df_violations = df_violations[["Rule", "Explanation"]]
        
            # Display table
            st.table(df_violations)


