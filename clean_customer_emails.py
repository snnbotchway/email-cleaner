import pandas as pd
import re

# Load the list of customer emails
customer_emails_df = pd.read_excel('ALL Clients EMail.xlsx')

# Load the list of blacklisted emails
blacklisted_emails_df = pd.read_csv('all_blacklisted_emails.csv')

# Extract email addresses from the blacklisted emails using regex


def extract_email_address(x):
    matches = re.findall(r'[\w\.-]+@[\w\.-]+', x)
    if matches:
        return matches[0].lower()
    else:
        return ''


blacklisted_emails_df['email_address'] = blacklisted_emails_df['email_address'].apply(
    extract_email_address)

# Clean up blacklisted emails with extra characters
blacklisted_emails_df['email_address'] = blacklisted_emails_df['email_address'].apply(
    lambda x: x.lstrip('|'))

# Convert all emails to lowercase for easy comparison
customer_emails_df['EMAILS'] = customer_emails_df['EMAILS'].str.lower()

# Clean up customer emails with extra characters
customer_emails_df['EMAILS'] = customer_emails_df['EMAILS'].apply(
    lambda x: x.lstrip('|'))

# Get a list of blacklisted emails
blacklisted_emails = blacklisted_emails_df['email_address'].tolist()

# Remove all blacklisted emails from the customer email list
cleaned_customer_emails_df = customer_emails_df[~customer_emails_df['EMAILS'].apply(
    lambda x: x.startswith('|') or x.endswith('|') or x in blacklisted_emails)]

# Save the cleaned list of customer emails to a new Excel file
cleaned_customer_emails_df.to_excel(
    'cleaned_customer_emails.xlsx', index=False)
