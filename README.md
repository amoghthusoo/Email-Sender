# Email Sender Application

This is a simple Email Sender application built using KivyMD, a Material Design framework for Kivy. The application allows users to send emails to multiple recipients with their respective marks for CIA-I and CIA-II exams. Users can select an `.xls` file containing recipient details, compose the email, and send it to all recipients.

## Features

- User-friendly interface with Material Design elements.
- Allows users to select an `.xls` file with recipient details.
- Option to download a format `.xls` file to enter recipient details.
- Send emails to multiple recipients with their respective marks.
- Show/hide password functionality.
- Progress indicator while sending emails.
- Success and error dialogs for user feedback.

## Requirements

- Python 3.6+
- Kivy 2.1.0+
- KivyMD 1.1.1+
- xlrd 1.2.0
- xlwt 1.3.0
- smtplib (Standard Library)
- email.mime (Standard Library)

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/yourusername/email-sender.git
   cd email-sender
   ```

2. Install the required packages:
   ```sh
   pip install -r requirements.txt
   ```

3. Run the application:
   ```sh
   python main.py
   ```

## Usage

1. Launch the application.
2. Enter your email address and password.
3. Click on "Select File" to choose an `.xls` file containing the recipient details (Email, Name, CIA-I Marks, CIA-II Marks).
4. (Optional) Click on "Download Format" to download a format `.xls` file to enter recipient details.
5. Choose the marks to be sent (CIA-I, CIA-II, or both) using the checkboxes.
6. Compose the email subject and message. Use `{}` as placeholders for the recipient's name and marks.
7. Click on "Send" to send the emails to all recipients.
8. A progress spinner will indicate the sending process. Success or error dialogs will provide user feedback.

## File Format

The `.xls` file should have the following columns:
- Email
- Name
- CIA - I
- CIA - II

Example:

| Email                | Name       | CIA - I | CIA - II |
|----------------------|------------|---------|----------|
| student1@example.com | Student 1  | 18      | 19       |
| student2@example.com | Student 2  | 15      | 17       |

## Customization

You can customize the application by modifying the `UI` string in the code for different layout and design adjustments.

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/amoghthusoo/Email-Sender/blob/master/LICENSE.txt) file for details.