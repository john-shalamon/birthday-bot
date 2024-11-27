import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
import os
import yagmail

# Load environment variables from .env file
load_dotenv()

# Fetch email credentials
EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')

# Validate email credentials
if not EMAIL or not PASSWORD:
    raise ValueError("Email credentials are missing in the .env file.")

# Function to read employee data from Excel
def read_employee_data(file_path):
    try:
        data = pd.read_excel(file_path)
        print("Columns in the Excel file:", data.columns)  # Debugging: Print columns
        data.columns = data.columns.str.strip()  # Remove extra spaces
        data.rename(columns={'Date of Birth': 'DOB'}, inplace=True)  # Rename column
        return data
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")


# Function to filter today's birthdays
def get_todays_birthdays(data):
    today = datetime.now().strftime('%m-%d')
    try:
        data['DOB'] = pd.to_datetime(data['DOB'], errors='coerce')  # Convert DOB column to datetime
        data['Birthday'] = data['DOB'].dt.strftime('%m-%d')        # Extract month-day format
        return data[data['Birthday'] == today]
    except KeyError as e:
        raise KeyError("The required column 'DOB' is missing in the Excel file.")
    except Exception as e:
        raise Exception(f"Error processing data for birthdays: {e}")

# Function to send personalized birthday emails
def send_emails(birthday_people):
    try:
        yag = yagmail.SMTP(EMAIL, PASSWORD)  # Initialize email client
        for _, person in birthday_people.iterrows():
            name = person['Name']
            email = person['Email']
            subject = "Happy Birthday!"
            content = f"""
            Hi {name},
            
            Wishing you a very Happy Birthday! 🎉🎂
            May your year be filled with joy, success, and happiness.
            
            Best Regards,
            Your Company
            """
            yag.send(to=email, subject=subject, contents=content)
            print(f"Email sent to {name} ({email})")
    except Exception as e:
        raise Exception(f"Error sending emails: {e}")

# Main function
if __name__ == "__main__":
    FILE_PATH = "data/employees.xlsx"  # Path to the Excel file
    
    try:
        # Step 1: Read employee data
        print("Reading employee data...")
        data = read_employee_data(FILE_PATH)
        
        # Step 2: Get today's birthdays
        print("Checking for today's birthdays...")
        birthday_people = get_todays_birthdays(data)
        
        # Step 3: Send birthday emails
        if not birthday_people.empty:
            print("Sending birthday emails...")
            send_emails(birthday_people)
        else:
            print("No birthdays today!")
    except Exception as e:
        print(f"An error occurred: {e}")
