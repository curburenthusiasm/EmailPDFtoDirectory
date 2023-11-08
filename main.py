import win32com.client
import datetime
import os
import re

directory = r""

# Create attachments directory if it does not exist
if not os.path.exists(directory):
    os.makedirs(directory)

# Shared mailbox email address
shared_mailbox_email = ""

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get shared mailbox
recipient = outlook.CreateRecipient(shared_mailbox_email)
recipient.Resolve()
shared_mailbox = outlook.GetSharedDefaultFolder(recipient, 6)

# Construct search criteria
today = datetime.date.today().strftime('%Y-%m-%d')
query = f"[ReceivedTime] > '{today}'"

messages = shared_mailbox.Items.Restrict(query)
print(f"Number of messages found: {len(messages)}")


# Remove invalid characters from file name
def sanitize_filename(filename):
    return re.sub(r'[\\/:"*?<>|]', "", filename)


# Loop through each message and extract attachments
for message in messages:
    print(f"Subject: {message.Subject}")

    # Remove colon (":") from subject header
    subject_header = message.Subject.replace(":", "")

    for attachment in message.Attachments:
        if attachment.FileName.endswith('.pdf'):
            # Sanitize and shorten the file name
            sanitized_subject_header = sanitize_filename(subject_header)[:50]
            sanitized_file_name = sanitize_filename(attachment.FileName)[:50]
            filename = os.path.join(directory, f"{sanitized_subject_header}_{sanitized_file_name}")

            try:
                attachment.SaveAsFile(filename)
                print(f"Attachment saved: {filename}")
            except Exception as e:
                print(f"Error saving attachment: {e}")