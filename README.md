# Automated Email Sender

This project is an automated Python script that extracts email addresses and messages from an Excel file and sends custom emails to them along with an attachment.

## Features

- Extracts email addresses and messages from an Excel file.
- Sends custom emails to each recipient in either French or English based on the specified language.
- Attaches a specified file to each email.

## Requirements

- Python 3.x
- `pandas` library
- `smtplib` library
- `email` library
- `openpyxl` library

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/aminemahd13/automail.git
    cd automail
    ```

2. Install the required libraries:
    ```sh
    pip install pandas openpyxl
    ```

## Usage

1. Prepare your Excel file (`email.xlsx`) with the following columns:
    - `Emails`: The recipient's email address.
    - `cnames`: The recipient's company name.
    - `language`: The language of the email (`FR` for French, `EN` for English).

2. Update the script with your email server settings and the path to the attachment files:
    ```python
    SenderAddress = "your-email@example.com"
    password = "your-email-password"
    folderpath = "path/to/your/attachments/"
    ```

3. Run the script:
    ```sh
    python automail.py
    ```


