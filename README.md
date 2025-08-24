# AutoMailGPT-

This script fetches emails from Microsoft Outlook, summarizes their content using the OpenAI API, and saves the summaries to a text file on your desktop.

Features

Fetch emails from your Outlook inbox for a given number of past days.

Automatically generate concise summaries of email bodies using OpenAI.

Save the summarized output to a file on your desktop.

Requirements

Python 3.8+

Microsoft Outlook installed and configured

Dependencies:

pywin32 (for Outlook access)

openai

Install dependencies:

pip install pywin32 openai

Setup
1. Clone the repository
git clone https://github.com/your-username/outlook-email-summarizer.git
cd outlook-email-summarizer

2. Configure your OpenAI API key

For security, do not hardcode your API key in the script.
Instead, store it in an environment variable.

Windows (PowerShell):
setx OPENAI_API_KEY "your_api_key_here"

macOS/Linux (bash/zsh):
export OPENAI_API_KEY="your_api_key_here"


Then, update the script to read the key:

import os
openai.api_key = os.getenv("OPENAI_API_KEY")

Usage

Run the script with the number of days back you want to fetch emails:

python app.py 7


This fetches emails from the last 7 days, summarizes them, and saves the result as:

Desktop/Email_Summary.txt
