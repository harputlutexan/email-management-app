# get_email_and_extract.py
import imaplib
import email
import pandas as pd
import os
from dotenv import load_dotenv
from openai import OpenAI
from datetime import datetime
from email.header import decode_header
import streamlit as st

load_dotenv()

EMAIL = st.secrets["EMAIL_ADDRESS"]
PASSWORD = st.secrets("EMAIL_PASSWORD")
IMAP_SERVER = "imap.one.com"
IMAP_PORT = 993
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)

def extract_info_with_openai(body_text):
    prompt = f"""Aşağıdaki e-posta mesajında geçen ürün (product), fiyat (price), teslimat yeri (delivery address) ve teslimat şekli (CIF-FOB vs)
     ile ilgili bilgileri yapılandırılmış biçimde çıkart. Yoksa 'None' yaz:

---
{body_text}
---

Yanıtı şu formatta ver:

Product: ...
Price: ...
Delivery Address: ...
Delivery Condition: ...
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        content = response.choices[0].message.content
        return content
    except Exception as e:
        return f"OpenAI error: {e}"

def get_emails_from_inbox(limit=3):
    results = []
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        status, messages = mail.search(None, 'ALL')
        mail_ids = messages[0].split()[-limit:]

        for mail_id in mail_ids:
            status, msg_data = mail.fetch(mail_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])

            from_ = msg.get("From")
            subject = msg.get("Subject")
            date = msg.get("Date")

            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode(errors="ignore")
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")

            extracted_info = extract_info_with_openai(body.strip()[:2000])

            results.append({
                "From": from_,
                "Subject": subject,
                "Date": date,
                "Body": body.strip()[:300],
                "Extracted Info": extracted_info
            })

        return pd.DataFrame(results)

    except Exception as e:
        return pd.DataFrame([{"Error": str(e)}])
