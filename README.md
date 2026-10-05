# Email Management App

A Python and Streamlit application for managing business email workflows from a single dashboard.

The app combines structured data processing, email automation, AI-assisted information extraction, and reporting tools. It was built as a practical workflow application rather than a standalone demo.

## Features

- Send product-related emails to suppliers using records stored in Excel
- Send customer outreach emails with per-recipient confirmation
- Read recent inbox messages through IMAP
- Use OpenAI models to extract structured information from incoming emails
- Export sent-email and extracted-email results to Excel
- Provide an AI assistant interface inside Streamlit
- Embed a Power BI report in the application dashboard

## Tech Stack

- Python
- Streamlit
- pandas
- OpenAI API
- SMTP / IMAP
- Excel processing with openpyxl and xlsxwriter
- Power BI embedding

## Project Structure

- `app.py` - Streamlit user interface and application navigation
- `send_email_to_suppliers.py` - supplier email workflow
- `send_email_to_customers.py` - customer email workflow
- `get_email_and_extract.py` - inbox retrieval and AI-assisted information extraction
- `requirements.txt` - Python dependencies

## Local Setup

1. Create and activate a Python virtual environment.
2. Install the dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Configure the required credentials outside the repository. The application expects:

   ```text
   EMAIL_ADDRESS
   EMAIL_PASSWORD
   OPENAI_API_KEY
   ```

   For Streamlit, these can be provided through `.streamlit/secrets.toml`.

4. Provide the required local data file used by the email workflows.
5. Start the application:

   ```bash
   streamlit run app.py
   ```

## Security

Credentials and local virtual environments are intentionally excluded from version control. Do not commit `.env`, `.venv`, or Streamlit secret files.

## Notes

This repository reflects a hands-on application combining backend-style workflow logic, external APIs, data processing, and an interactive Streamlit interface.
