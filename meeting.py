import os
import requests
import pandas as pd
from docx import Document
from datetime import datetime
import json

# ====== Load .env ======
from dotenv import load_dotenv
load_dotenv()
os.environ['GROQ_API_KEY'] = os.getenv('GROQ_API_KEY')
groq_api_key = os.getenv('GROQ_API_KEY')

# ====== CONFIG ======
GROQ_API_KEY = groq_api_key  # from console.groq.com
TRANSCRIPT_FILE = "Text_Summarizer.docx"  # path to Teams transcript file
EXCEL_FILE = "Meeting_summary_template.xlsx"  # output Excel tracker

# ====== JSON SCHEMA PROMPT ======
PROMPT_TEMPLATE = """
You are a meeting summarizer. 
Return the result ONLY in strict JSON following this schema, without extra text:

{
  "MeetingDetails": {
    "Date & Time": "",
    "Location": "",
    "Participants": []
  },
  "Objective": "",
  "AgendaItems": [],
  "KeyDiscussions": "",
  "DecisionsMade": "",
  "ActionItems": [
    {"Task": "", "Owner": "", "DueDate": ""}
  ],
  "NextSteps": "",
  "AdditionalNotes": ""
}
"""

# ====== STEP 1: Read transcript ======
def read_docx(file_path):
    doc = Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])

transcript = read_docx(TRANSCRIPT_FILE)

# ====== STEP 2: Call Groq API ======
url = "https://api.groq.com/openai/v1/chat/completions"
headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
payload = {
    "model": "llama-3.1-8b-instant",
    "messages": [
        {"role": "system", "content": "You are a meeting summarizer."},
        {"role": "user", "content": PROMPT_TEMPLATE + "\n\nTranscript:\n" + transcript}
    ],
    "temperature": 0.2,
    "response_format": {"type": "json_object"}  # ✅ Force JSON
}

response = requests.post(url, headers=headers, json=payload)
result = response.json()

# ====== STEP 3: Extract JSON summary ======
summary_json = json.loads(result["choices"][0]["message"]["content"])

# ====== STEP 4: Flatten JSON for Excel ======
def flatten_summary(summary):
    row = {
        "Meeting_DateTime": summary["MeetingDetails"].get("Date & Time", ""),
        "Meeting_Location": summary["MeetingDetails"].get("Location", ""),
        "Meeting_Participants": ", ".join(summary["MeetingDetails"].get("Participants", [])),
        "Objective": summary.get("Objective", ""),
        "AgendaItems": "; ".join(summary.get("AgendaItems", [])),
        "KeyDiscussions": summary.get("KeyDiscussions", ""),
        "DecisionsMade": summary.get("DecisionsMade", ""),
        "NextSteps": summary.get("NextSteps", ""),
        "AdditionalNotes": summary.get("AdditionalNotes", "")
    }

    # Handle multiple Action Items
    if "ActionItems" in summary and summary["ActionItems"]:
        for i, action in enumerate(summary["ActionItems"], start=1):
            row[f"ActionItem_{i}_Task"] = action.get("Task", "")
            row[f"ActionItem_{i}_Owner"] = action.get("Owner", "")
            row[f"ActionItem_{i}_DueDate"] = action.get("DueDate", "")
    return row

row = flatten_summary(summary_json)

# ====== STEP 5: Save to Excel ======
if os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
else:
    df = pd.DataFrame([row])

df.to_excel(EXCEL_FILE, index=False)
print("✅ Structured meeting summary added to", EXCEL_FILE)
