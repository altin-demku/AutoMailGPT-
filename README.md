# AutoMailGPT-
# Outlook Email Summarizer

A Python script that fetches emails from Microsoft Outlook, summarizes their content using the OpenAI API, and saves the summaries to a text file on your desktop.

## Features
- Fetch emails from your Outlook inbox for a specified number of past days.
- Summarize email content automatically using OpenAI.
- Save summarized output to a text file on your desktop.

## Requirements
- Python 3.8+
- Microsoft Outlook installed and configured
- Python packages:
  - `pywin32`
  - `openai`

Install dependencies:

``bash
pip install pywin32 openai

##Setup
1. Clone the repository
git clone https://github.com/your-username/outlook-email-summarizer.git

2. Configure your OpenAI API key

For security, do not hardcode your API key in the script.
Instead, store it in an environment variable.

##Usage

Run the script with the number of days back you want to fetch emails:

python app.py 7


This fetches emails from the last 7 days, summarizes them, and saves the result as:

Desktop/Email_Summary.txt
