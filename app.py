import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import os
from send_email_to_customers import send_email_to_customers
from send_email_to_suppliers import send_email_to_suppliers
from get_email_and_extract import get_emails_from_inbox

st.set_page_config(page_title="Email Management Panel", layout="wide")
#st.set_page_config(page_title="Email Management Panel", layout="centered")
st.title("📧 Email Management Panel")

menu = st.sidebar.radio("Menü Seç", [
    "Tedarikçiye Mail Gönder",
    "Müşteriye Mail Gönder",
    "Gelen Mailleri Oku",
    "AI Chat (Assistant)",
    "Power BI Raporu"
])

if menu == "Tedarikçiye Mail Gönder":
    st.header("📤 Tedarikçiye Mail Gönder")

    query_product = st.text_input("Ürün Adı (örn. TİTANYUM)")
    test_email = st.text_input("Test E-mail (zorunlu değil)", value="")

    if st.button("Mail Gönder"):
        if query_product:
            with st.spinner("Mailler gönderiliyor..."):
                progress_bar = st.progress(0)

                result_text, sent_log = send_email_to_suppliers(
                    query_product,
                    test_email or None,
                    progress_callback=progress_bar.progress
                )

            st.success(result_text)

            if sent_log:
                for entry in sent_log:
                    st.markdown(f"- ✅ Sent to: **{entry['Email']}** – _{entry['Product']}_")

                df_log = pd.DataFrame(sent_log)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_log.to_excel(writer, index=False, sheet_name="Sent Emails")
                output.seek(0)

                st.download_button(
                    label="📥 Download Sent Email Log (Excel)",
                    data=output,
                    file_name="sent_emails_log.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Optional: Remove partial_log.xlsx
                if os.path.exists("partial_log.xlsx"):
                    os.remove("partial_log.xlsx")

            else:
                st.warning("Hiçbir e-posta gönderilmedi.")
        else:
            st.warning("Lütfen ürün adı girin.")

elif menu == "Müşteriye Mail Gönder":
    st.header("📧 Müşteriye Mail Gönder")

    query_product = st.text_input("Ürün Adı (örn. TİTANYUM)")
    test_email = st.text_input("Test E-mail (zorunlu değil)", value="")

    if st.button("Müşteri Maillerini Gönder"):
        if query_product:
            with st.spinner("Mailler gönderiliyor..."):
                progress_bar = st.progress(0)
                result_text, sent_log = send_email_to_customers(
                    query_product,
                    test_email or None,
                    progress_callback=progress_bar.progress
                )

            st.success(result_text)

            if sent_log:
                for entry in sent_log:
                    st.markdown(f"- ✅ Sent to: **{entry['Email']}** – _{entry['Product']}_")

                df_log = pd.DataFrame(sent_log)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_log.to_excel(writer, index=False, sheet_name="Customer Emails")
                output.seek(0)

                st.download_button(
                    label="📥 Download Sent Customer Log (Excel)",
                    data=output,
                    file_name="customer_emails_log.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

elif menu == "Gelen Mailleri Oku":
    st.header("📥 Gelen Mailler ve Bilgi Çıkartma")

    limit = st.slider("Kaç adet son e-posta kontrol edilsin?", 1, 10, 3)

    if st.button("📬 Mailleri Getir ve Yorumla"):
        with st.spinner("E-postalar alınıyor ve OpenAI ile işleniyor..."):
            df = get_emails_from_inbox(limit)

        if "Error" in df.columns:
            st.error(f"❌ Hata: {df['Error'].iloc[0]}")
        else:
            st.success("✅ E-postalar işlendi!")

            for i, row in df.iterrows():
                st.markdown(f"### ✉️ Mail {i+1}")
                st.markdown(f"- **From:** {row['From']}")
                st.markdown(f"- **Subject:** {row['Subject']}")
                st.markdown(f"- **Date:** {row['Date']}")
                st.markdown(f"**Body Preview:** `{row['Body']}`")
                st.markdown(f"**📊 Extracted Info:**\n```\n{row['Extracted Info']}\n```")

            # Excel Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="📥 Excel olarak indir",
                data=output,
                file_name="received_emails_extracted_info.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

elif menu == "AI Chat (Assistant)":
    st.header("💬 Chat with Assistant")

    user_input = st.chat_input("Type your question...")

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    if user_input:
        st.session_state.chat_history.append({"role": "user", "content": user_input})

        with st.spinner("Assistant is thinking..."):
            try:
                from openai import OpenAI

                # ✅ API key'i st.secrets içinden al
                client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
                assistant_id = "asst_8pJbOUPHAZ3NuMlw9Ona1TTW"  # 🔁 replace with your real ID
                thread = client.beta.threads.create()

                client.beta.threads.messages.create(
                    thread_id=thread.id,
                    role="user",
                    content=user_input
                )

                run = client.beta.threads.runs.create(
                    thread_id=thread.id,
                    assistant_id=assistant_id
                )

                while run.status != "completed":
                    run = client.beta.threads.runs.retrieve(
                        thread_id=thread.id,
                        run_id=run.id
                    )

                messages = client.beta.threads.messages.list(thread_id=thread.id)
                response = messages.data[0].content[0].text.value
                st.session_state.chat_history.append({"role": "assistant", "content": response})

            except Exception as e:
                st.error(f"❌ Error: {e}")

    for msg in st.session_state.chat_history:
        if msg["role"] == "user":
            st.chat_message("user").markdown(msg["content"])
        else:
            st.chat_message("assistant").markdown(msg["content"])


elif menu == "Power BI Raporu":
    st.header("📊 Power BI Raporu")
    powerbi_url = "https://app.powerbi.com/view?r=eyJrIjoiMzI1YzljZjYtYTc2Yy00MjI5LWI2MzctODQzNjBjOWY5NDg3IiwidCI6ImQxZjRlN2VjLTAwOWEtNDJhZC05MWQ4LWVlOTQ2NDYxNGMwNyIsImMiOjl9"


    components.html(
        f"""
        <div style="position: relative; width: 100vw; height: 1000px; left: -5vw;">
            <iframe title="Power BI Report"
                    width="100%"
                    height="100%"
                    src="{powerbi_url}"
                    style="border:none;">
            </iframe>
        </div>
        """,
        height=1000,
        scrolling=True
    )



