import pandas as pd
import random
import re
import string
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def generate_unique_code(code_length=5):
    """Generate a random alphanumeric code."""
    return ''.join(random.choices('0123456789', k=code_length))

def fetch_emails_from_excel(file_path):
    """Load emails from Excel file and return as a list."""
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip().str.lower()  # Normalize column names
    return df

# Function to validate email with regular expression
def is_valid_email(email):
    """Validate if email is of the form 'something@kab.ac.ug'."""
    pattern = r"^[a-zA-Z0-9._%+-]+@kab\.ac\.ug$"
    return re.match(pattern, email) is not None

def send_verification_email(receiver_email, unique_code, sender_email, sender_password):
    """Send an email with a unique verification code."""
    subject = "Voter's Verification Code"
    message_body = f"""
    Dear Voter,

    Your unique verification code for the upcoming election is: {unique_code}
    
    Please keep this code confidential and do not share it with anyone.
        NOTE: voting will start exactly  at 1:00pm and ends at 4:00pm 
        venue :room 4 guild offices 
    
    Regards,
    Indabax Kabale University Club Election Committee
    """
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)

        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject
        message.attach(MIMEText(message_body, 'plain'))

        server.send_message(message)
        print(f"Verification email sent to {receiver_email}")
        
    except Exception as e:
        print(f"Error sending email to {receiver_email}: {e}")
    finally:
        server.quit()

def save_codes_to_excel(df, file_path):
    """Save the updated DataFrame with codes back to Excel."""
    df.to_excel(file_path, index=False)
    print(f"Codes saved to {file_path}")

# Example usage
file_path = 'emails.xlsx'
emails_df = fetch_emails_from_excel(file_path)

sender_email = "indabaxkabale@gmail.com"  # Replace with your email
sender_password = "frha hexn jeab uesj"  # Replace with app-specific password

# Generate a unique code for each email and send the email
emails_df['verification_code'] = emails_df['email'].apply(lambda x: generate_unique_code())  # Create a new column for codes

# Filter out emails that do not match the 'kab.ac.ug' domain
valid_emails_df = emails_df[emails_df['email'].apply(is_valid_email)]

# Send verification emails to valid email addresses only
for _, row in valid_emails_df.iterrows():
    email = row['email']
    unique_code = row['verification_code']
    # print(f"Email: {email} -> Unique Code: {unique_code}")
    
    send_verification_email(email, unique_code, sender_email, sender_password)

# Save the updated DataFrame with the codes back into the Excel file
save_codes_to_excel(emails_df, file_path)
