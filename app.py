import os
import win32com.client
import openai
import sys
from datetime import datetime, timedelta


# Function to get emails from Outlook
def fetch_emails_from_outlook(days_back):
    # Access Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox folder

    # Calculate the date range for filtering emails
    since_date = datetime.now() - timedelta(days=days_back)

    # Filter emails by the date range
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by received time

    emails = []
    for message in messages:
        # Filter based on the received date
        if message.ReceivedTime < since_date:
            break

        # Append the email details (subject, body, and received time)
        email = {
            "subject": message.Subject,
            "body": message.Body,
            "received_time": message.ReceivedTime
        }
        emails.append(email)

    return emails


# Function to summarize email content using OpenAI
def summarize_text(text):
    openai.api_key = "your_openai_api_key"

    response = openai.Completion.create(
        model="text-davinci-003",  # Or any other model you'd like to use
        prompt=f"Please summarize the following email content:\n{text}",
        max_tokens=100,
        temperature=0.5,
    )

    return response.choices[0].text.strip()


# Function to save the summary to the desktop
def save_file_to_desktop(content, filename):
    # Get the path to the user's desktop
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

    # Create the full path to the file
    file_path = os.path.join(desktop_path, filename)

    # Save the content to the file
    with open(file_path, "w") as file:
        file.write(content)

    print(f"Summary saved to {file_path}")


def main():
    # Check for command-line argument (days_back)
    if len(sys.argv) != 2:
        print("Usage: app.exe <days_back>")
        sys.exit(1)

    # Get the number of days back from the command-line argument
    days_back = int(sys.argv[1])

    # Fetch the emails
    emails = fetch_emails_from_outlook(days_back)

    if not emails:
        print("No emails found for the specified date range.")
        return

    # Generate the summary for each email and combine them
    summary = ""
    for email in emails:
        print(f"Summarizing email: {email['subject']}")
        summarized_text = summarize_text(email["body"])
        summary += f"Subject: {email['subject']}\nReceived Time: {email['received_time']}\nSummary: {summarized_text}\n\n"

    # Save the summary to a text file on the desktop
    save_file_to_desktop(summary, "Email_Summary.txt")


# Run the script
if __name__ == "__main__":
    main()
