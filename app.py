import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from langchain.chat_models import ChatOpenAI
from langchain.tools import tool
from langchain.agents import initialize_agent, AgentType

# --------------------------
# Google Sheets Setup
# --------------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
try:
    creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    gc = gspread.authorize(creds)
except:
    gc = None  # If no Google Sheets credentials

# --------------------------
# Tools for Excel & Google Sheets
# --------------------------

@tool
def read_excel(file_path: str) -> str:
    df = pd.read_excel(file_path)
    return f"Columns: {list(df.columns)}\n\nPreview:\n{df.head().to_string()}"

@tool
def filter_and_write_excel(file_path: str, condition: str, new_sheet: str) -> str:
    df = pd.read_excel(file_path)
    filtered = df.query(condition)
    with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        filtered.to_excel(writer, sheet_name=new_sheet, index=False)
    return f"✅ Wrote {len(filtered)} rows to sheet '{new_sheet}' in {file_path}."

@tool
def read_gsheet(sheet_id: str, worksheet: str = None) -> str:
    if not gc:
        return "❌ Google Sheets not configured."
    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(worksheet) if worksheet else sh.sheet1
    df = pd.DataFrame(ws.get_all_records())
    return f"Columns: {list(df.columns)}\n\nPreview:\n{df.head().to_string()}"

@tool
def filter_and_write_gsheet(sheet_id: str, condition: str, new_sheet: str) -> str:
    if not gc:
        return "❌ Google Sheets not configured."
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1
    df = pd.DataFrame(ws.get_all_records())

    filtered = df.query(condition)

    try:
        new_ws = sh.worksheet(new_sheet)
        sh.del_worksheet(new_ws)
    except:
        pass

    new_ws = sh.add_worksheet(title=new_sheet, rows=len(filtered)+1, cols=len(filtered.columns))
    new_ws.update([filtered.columns.values.tolist()] + filtered.values.tolist())
    return f"✅ Wrote {len(filtered)} rows to Google Sheet tab '{new_sheet}'."

# --------------------------
# LangChain Agent Setup
# --------------------------

from langchain.chat_models import ChatOpenAI
import streamlit as st

llm = ChatOpenAI(
    model="gpt-4o-mini",
    temperature=0,
    openai_api_key=st.secrets["OPENAI_API_KEY"]  # <- fetches key from Streamlit Secrets
)


tools = [read_excel, filter_and_write_excel, read_gsheet, filter_and_write_gsheet]

agent = initialize_agent(
    tools,
    llm,
    agent=AgentType.ZERO_SHOT_REACT_DESCRIPTION,
    verbose=True
)

# --------------------------
# Streamlit UI
# --------------------------

st.set_page_config(page_title="Spreadsheet AI Agent", page_icon="📊")

st.title("📊 Spreadsheet AI Agent")
st.write("Give natural language instructions to move/filter data in Excel or Google Sheets.")

user_input = st.text_input("Enter your command:")

if user_input:
    with st.spinner("Working..."):
        response = agent.run(user_input)
    st.success("Done!")
    st.write(response)
