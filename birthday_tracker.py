import openpyxl
import smtplib
from datetime import datetime
import schedule
import time

# Step 1: Load birthdays from an Excel file
def load_birthdays(file_path):
    birthdays = []
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Assuming the columns are: Name, Birthday, Email (row 1 is the header)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name, birthday, email = row
        birthdays.append({
            "name": name,
            "birthday": birthday.strftime("%m-%d"),  # Convert date to MM-DD format
            "email": email
        })
    return birthdays


# Step 2: Send email
def send_email(to_email, subject, body, from_email, password):
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(from_email, password)
            message = f"Subject: {subject}\n\n{body}"
            server.sendmail(from_email, to_email, message)
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

# Step 3: Check for birthdays and send wishes
def check_and_send_wishes():
    # Load birthday data
    file_path = "Birthday_Tracker.xlsx"
    birthdays = load_birthdays(file_path)

    # Today's date
    today = datetime.now().strftime("%m-%d")

    # Sender credentials (use environment variables for security)
    from_email = "mgrsam778@gmail.com"
    password = "yqfph oadv ymdd kijw"

    for person in birthdays:
        if person["birthday"] == today:
            to_email = person["email"]
            name = person["name"]
            subject = "Happy Birthday!"
            body = f"Hi {name},\n\nWishing you a very Happy Birthday! Have an amazing day ahead!\n\nBest regards,\nSamiksha"
            send_email(to_email, subject, body, from_email, password)

# Step 4: Schedule the script to run daily
def main():
    schedule.every().day.at("08:00").do(check_and_send_wishes)
    print("Birthday Email Automator is running...")
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    main()
