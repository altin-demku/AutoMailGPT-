import os
import win32com.client
import openai
import sys
from datetime import datetime, timedelta


def fetch_emails_from_outlook(days_back):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)

    since_date = datetime.now() - timedelta(days=days_back)

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    emails = []
    for message in messages:
        if message.ReceivedTime < since_date:
            break

        email = {
            "subject": message.Subject,
            "body": message.Body,
            "received_time": message.ReceivedTime
        }
        emails.append(email)

    return emails


def summarize_text(text):
    openai.api_key = "your_openai_api_key"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=f"Please summarize the following email content:\n{text}",
        max_tokens=100,
        temperature=0.5,
    )

    return response.choices[0].text.strip()


def save_file_to_desktop(content, filename):
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

    file_path = os.path.join(desktop_path, filename)

    with open(file_path, "w") as file:
        file.write(content)

    print(f"Summary saved to {file_path}")


def main():
    if len(sys.argv) != 2:
        print("Usage: app.exe <days_back>")
        sys.exit(1)

    days_back = int(sys.argv[1])

    emails = fetch_emails_from_outlook(days_back)

    if not emails:
        print("No emails found for the specified date range.")
        return

    summary = ""
    for email in emails:
        print(f"Summarizing email: {email['subject']}")
        summarized_text = summarize_text(email["body"])
        summary += f"Subject: {email['subject']}\nReceived Time: {email['received_time']}\nSummary: {summarized_text}\n\n"

    save_file_to_desktop(summary, "Email_Summary.txt")


if __name__ == "__main__":
    main()
