# send_email_to_suppliers.py
import pandas as pd
import smtplib
import os
import unicodedata
import time
import random
from email.message import EmailMessage
from dotenv import load_dotenv
from datetime import datetime
import streamlit as st
load_dotenv()

SMTP_SERVER = "send.one.com"
SMTP_PORT = 465
SMTP_USER = st.secrets["EMAIL_ADDRESS"]
SMTP_PASSWORD = st.secrets["EMAIL_PASSWORD"]

BATCH_SIZE = 10
DELAY_BETWEEN_EMAILS = (6, 12)
DELAY_BETWEEN_BATCHES = 60


def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = text.replace("ı", "i").replace("İ", "i")
    text = unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("utf-8")
    return " ".join(text.split()).strip()


def create_email_body(row):
    return f"""\
Dear {row['company_name']},

We are planning to purchase the following product and would appreciate your offer:

Product: {row['product_name']}
Unit: {row['unit']}
Delivery Term: {row['delivery_term']}
Delivery Location: {row['delivery_location']}

Could you please provide us with your best price offer, estimated availability, and preparation/lead time?

- Product Description:
- Loading Quantity:
- FOB Price:
- Loading Port:
- CIF Price:
- Destination Port:
- Payment Term:

Best regards,  
Musab Uslu  
TheStar Purchasing Team
"""


def send_email_to_suppliers(query_product: str, test_email: str = None, progress_callback=None):
    # [Load dotenv, SMTP info etc. same as before]

    df = pd.read_excel("bazdata_final_with_emails_corrected_full2.xlsx")
    # [Clean and normalize columns, same as before]

    # Rename columns (required to make product_name etc. work)
    df = df.rename(columns={
        "Gönderici Adı Ünvanı": "company_name",
        "Ticari Tanım": "product_name",
        "Satışa Esas Miktar Birimi Kodu": "unit",
        "Teslim Şekli Kodu": "delivery_term",
        "Teslim Yeri": "delivery_location",
        "Email": "email"
    })

    df = df.dropna(subset=["email"])  # filter rows without email

    # ✅ THIS was missing!
    df["product_name_clean"] = df["product_name"].apply(normalize_text)


    df_filtered = df[df["product_name_clean"].str.contains(normalize_text(query_product), na=False)]

    sent_log = []
    total = len(df_filtered)

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SMTP_USER, SMTP_PASSWORD)

        for i, (_, row) in enumerate(df_filtered.iterrows()):
            msg = EmailMessage()
            msg["Subject"] = f"Offer Request – {row['product_name']}"
            msg["From"] = f"Musab Uslu <{SMTP_USER}>"
            msg["To"] = test_email if test_email else row["email"]
            msg.set_content(create_email_body(row))

            try:
                server.send_message(msg)
                sent_log.append({
                    "Email": row["email"],
                    "Product": row["product_name"],
                    "Date Sent": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })

                # 💾 Optional: also write a growing file if user stops mid-way
                pd.DataFrame(sent_log).to_excel("partial_log.xlsx", index=False)

            except Exception as e:
                print(f"❌ Error sending to {row['email']}: {e}")

            # ⏳ Show progress
            if progress_callback:
                progress_callback((i + 1) / total)

            time.sleep(random.uniform(*DELAY_BETWEEN_EMAILS))

    return f"{len(sent_log)} emails sent for product: {query_product}", sent_log

