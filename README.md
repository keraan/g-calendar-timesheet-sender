# Timesheet Automation Tool for Masman Group

The **Timesheet Automation Tool** is a Python script that automates the process of creating and sending timesheets via Gmail. It utilizes Google APIs for authentication, email sending, and Google Calendar integration to create a comprehensive timesheet report and deliver it to the specified recipient. Useful for Masman Group employees who want to 
streamline the process of submitting their timesheets using Google Calendar.

## Features

- Automatically generates a timesheet report for the past week (last 7 days) based on the user's worked hours on Google Calendar.
- Uses Google APIs to authenticate and send emails, ensuring secure and reliable delivery of the timesheet report.
- Uses Google APIs to fetch data from the user's Google Calendar.
- Customizable configurations through the `config.json` file, allowing users to define personal details, recipient's email, and file paths.

## Getting Started

Follow these steps to set up and use the Timesheet Automation Tool:

1. **Clone the Repository**: Clone this repository to your local machine.

2. **Install Dependencies**: Install the required Python packages by running the following command:

```
pip install -r requirements.txt
```

3. **Configure the Tool**: Edit the `config.json` file with your personal details, including your first name, last name, email, and recipient's email. You'll also need to provide the paths to your `client_secret_file.json` (for Google API authentication) and the `original_file` (the template timesheet file you want to use).

See: https://developers.google.com/workspace/guides/create-credentials#:~:text=In%20the%20Google%20Cloud%20console%2C%20go%20to%20Menu,%3E%20APIs%20%26%20Services%20%3E%20Credentials.&text=Click%20Create%20credentials%20%3E%20API%20key,new%20API%20key%20is%20displayed.

4. **Run the Script**: Run the `mailSender.py` script using the following command:

```
python3 mailSender.py
```

The script will automatically calculate your working hours from your Google Calendar.

5. **Email Delivery**: The tool will automatically create a timesheet report in Excel format and send it as an attachment via Gmail to the specified recipient.

6. **File Saving**: A copy of the timesheet report will be saved locally for your records.

## Note

- The script uses Google APIs for authentication. Ensure you have the necessary credentials (`client_secret_file.json`) and permissions to send emails using your Gmail account.

- The timesheet report is generated in Excel format and is attached to the email.

- Please ensure you have a stable internet connection while running the script, as it relies on Google services for authentication and email sending.

- Before using the tool, familiarize yourself with the Gmail API and Google Sheets API documentation.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

**Disclaimer**: This tool is provided as-is and without warranty. Use it responsibly and ensure you have the necessary permissions to access and modify the provided files.