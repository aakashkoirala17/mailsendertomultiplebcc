
import pandas as pd
import webbrowser

# Load the Excel data into a pandas DataFrame
file_path = 'your_excel_file_name.yxlsx'
sheet_name = 'Sheet 1'
column_name = 'Email:'

# Read the Excel file and extract email addresses
data = pd.read_excel(file_path, sheet_name=sheet_name)
email_list = data[column_name]

# Print the extracted email addresses
print("Extracted Email Addresses:")
for email in email_list:
    print(email)

# Ask the user to select email addresses for BCC
selected_bcc_emails = []
for email in email_list:
    use_email = input(f"Do you want to use {email} as a BCC recipient? (yes/no): ")
    if use_email.lower() == 'yes':
        selected_bcc_emails.append(email)

if selected_bcc_emails:
    # Construct the Gmail compose URL with BCC recipients
    bcc_recipients = ','.join(selected_bcc_emails)
    gmail_url = f"https://mail.google.com/mail/u/0/?view=cm&fs=1&bcc={bcc_recipients}"

    # Open the Gmail compose URL in a web browser
    webbrowser.open_new_tab(gmail_url)
    print(f"A web browser has been opened to compose an email with BCC to {', '.join(selected_bcc_emails)}.")
else:
    print("No email addresses selected for BCC.")
