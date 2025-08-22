# send_email_to_customers.py
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

def generate_description(product_name):
    if pd.isna(product_name):
        return "High-quality titanium dioxide pigment."
    name = product_name.lower()
    if "kronos" in name and "2160" in name:
        return "A high-performance pigment with excellent opacity and UV resistance."
    elif "kronos" in name and "2360" in name:
        return "A technical-grade TiO₂ suitable for industrial coatings."
    elif "billions tr52" in name:
        return "A chloride-process TiO₂ pigment known for its whiteness and dispersion."
    elif "rc 84" in name:
        return "A rutile titanium dioxide pigment offering high durability and gloss."
    elif "anatase" in name:
        return "An anatase-grade TiO₂ pigment ideal for plastics and rubber applications."
    elif "bey-az" in name or "beyaz" in name:
        return "A general-purpose white pigment with strong brightness and covering power."
    else:
        return "High-quality titanium dioxide pigment designed for versatile applications."

def create_email_body(row):

    description = generate_description(row['product_name'])
    
    return f"""\
Dear Valued Partner,

We are pleased to introduce one of our premium products, available for prompt delivery:

**Product:** {row['product_name']}  
**Description:** {description} 
**Unit:** {row['unit']}  
**Delivery Term:** {row['delivery_term']}  
**Delivery Location:** {row['delivery_location']}

Our {row['product_name']} products are trusted by manufacturers worldwide for their consistent quality, high whiteness, and reliable supply.

Whether you're looking to improve your current formulations or secure a new supplier, we would be glad to assist you.

For detailed specifications, a price quote, or sample requests, please feel free to contact us.

We look forward to hearing from you.

Best regards,  
Musab Uslu  
Sales Manager | TheStar Trading  
📧 {SMTP_USER}  
🌐 www.thestartrading.com
"""


def send_email_to_customers(query_product: str, test_email: str = None, progress_callback=None):
    df = pd.read_excel("proaktifPricing/bazdata_final_with_emails_corrected kopyası.xlsx")

    df = df.rename(columns={
        "Gönderici Adı Ünvanı": "company_name",
        "Ticari Tanım": "product_name",
        "Satışa Esas Miktar Birimi Kodu": "unit",
        "Teslim Şekli Kodu": "delivery_term",
        "Teslim Yeri": "delivery_location",
        "Email": "email"
    })
    df = df.dropna(subset=["email"])

    df["product_description"] = df["product_name"].apply(generate_description)
    df["product_name_clean"] = df["product_name"].apply(normalize_text)
    query_product_clean = normalize_text(query_product)
    df_filtered = df[df["product_name_clean"].str.contains(query_product_clean, na=False)]

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
                    "Company Name": row["company_name"],
                    "Email": row["email"],
                    "Product": row["product_name"],
                    "Date Sent": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })

                pd.DataFrame(sent_log).to_excel("partial_log_customers.xlsx", index=False)

            except Exception as e:
                print(f"❌ Error: {e}")

            if progress_callback:
                progress_callback((i + 1) / total)

            time.sleep(random.uniform(*DELAY_BETWEEN_EMAILS))

    return f"{len(sent_log)} customer emails sent for product: {query_product}", sent_log
