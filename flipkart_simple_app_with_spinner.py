
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import tempfile
import os

def extract_text_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text
    except Exception as e:
        return None, str(e)

def extract_invoice_data(text):
    orders = []
    matches = re.findall(r"Order Id:\s*(OD\d+).*?AW\s*B.*?(FM\w+).*?Name: (.*?),\n(.*?)\n.*?Total\s+PRICE: (\d+\.\d{2})", text, re.DOTALL)
    for order_id, awb_id, name, address, total in matches:
        pin_match = re.search(r"(\d{6})", address)
        pin_code = pin_match.group(1) if pin_match else ""
        state_match = re.search(r"IN-([A-Z]{2})", address)
        state = state_match.group(1) if state_match else ""
        payment_mode = "COD" if "COD" in text else "PREPAID"
        orders.append({
            "Order ID": order_id,
            "AWB ID": awb_id,
            "Customer Name": name.strip(),
            "Shipping Address": address.strip(),
            "Pin Code": pin_code,
            "State": state,
            "Total Amount": float(total),
            "Payment Mode": payment_mode,
            "Product": "Maxotech Affleck Ultra High Speed 24 Inch"
        })
    return pd.DataFrame(orders)

st.title("📦 Flipkart PDF to Excel (Fast & Simple)")

uploaded_file = st.file_uploader("Upload Flipkart Label PDF", type="pdf")

if uploaded_file is not None:
    with st.spinner("📄 Reading PDF and extracting data..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_file_path = tmp_file.name

        text = extract_text_from_pdf(tmp_file_path)

        if not text:
            st.error("❌ Failed to extract text. Please try another PDF.")
        else:
            df = extract_invoice_data(text)
            st.write("✅ Extracted Orders", df)

            if not df.empty:
                output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
                df.to_excel(output_path, index=False)
                with open(output_path, "rb") as f:
                    st.download_button("Download Excel", f, file_name="Flipkart_Orders.xlsx")

        os.remove(tmp_file_path)
