# 📝 Meeting Summarizer using Groq API

This project automates the process of summarizing meeting transcripts from Microsoft Teams using the **Groq LLM API** and stores the structured summary in an Excel tracker.

## 🔧 Features

- Reads `.docx` transcript files
- Sends transcript to Groq's LLaMA 3.1 model for summarization
- Extracts structured JSON summary with:
  - Meeting details
  - Agenda items
  - Key discussions
  - Decisions made
  - Action items
- Flattens JSON into tabular format
- Appends summary to an Excel tracker

## 📁 File Structure

- `meeting.py`: Main script to run the summarization pipeline
- `Text_Summarizer.docx`: Input transcript file (example)
- `Meeting_summary_template.xlsx`: Output Excel file with structured summaries

## 🚀 How to Run

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/meeting-summarizer.git
   cd meeting-summarizer

2. Install dependencies:
```
pip install -r requirements.txt
```

3. Add your .env file with API keys.
