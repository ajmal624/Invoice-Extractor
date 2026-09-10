import os
import re
import io
import json
import base64
import zipfile
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from io import BytesIO
from PIL import Image
from dotenv import load_dotenv
from openai import OpenAI

# ========= CONFIG =========
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    st.error("❌ OPENAI_API_KEY not found in .env file")
    st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)

st.set_page_config(page_title="📄 Invoice Data Extractor", layout="wide")
st.title("📄 Invoice Data Extractor — GPT-4.1-mini")

# ========= SESSION STATE =========
for key in ["parsed_data", "df_summary", "items_df", "df_mapping", "df_custom"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ========= FILE UPLOAD =========
uploaded_pdf = st.file_uploader("📤 Upload Invoice PDF", type=["pdf"])
uploaded_template = st.file_uploader("📋 Upload Excel Template", type=["xlsx"])

# ========= HELPERS =========
def pdf_to_images(pdf_bytes, dpi=300):
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    imgs = []
    for i in range(len(pdf_doc)):
        pix = pdf_doc.load_page(i).get_pixmap(dpi=dpi)
        img = Image.open(BytesIO(pix.tobytes("png"))).convert("RGB")
        imgs.append(img)
    pdf_doc.close()
    return imgs

def encode_img_b64(pil_img):
    buf = BytesIO()
    pil_img.save(buf, format="PNG")
    return f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}"

def clean_json_output(raw_text):
    """Clean and safely parse GPT JSON output."""
    if not raw_text:
        return {"error": "Empty response"}
    text = re.sub(r"^```(?:json)?|```$", "", raw_text.strip(), flags=re.MULTILINE)
    text = text.replace("\n", " ").strip()

    try:
        parsed = json.loads(text)
        if isinstance(parsed, list):
            # Wrap list in dict to handle gracefully
            return {"data": parsed}
        return parsed
    except Exception:
        try:
            text_fixed = re.sub(r"([{,])\s*([A-Za-z0-9_]+):", r'\1 "\2":', text)
            parsed = json.loads(text_fixed)
            if isinstance(parsed, list):
                return {"data": parsed}
            return parsed
        except Exception as e2:
            return {"error": f"JSON parse failed: {e2}", "raw": text}

def flatten_dict(d, parent="", sep="_"):
    """Flatten nested dictionaries for summary view."""
    if not isinstance(d, dict):
        return {}
    items = []
    for k, v in d.items():
        new_k = f"{parent}{sep}{k}" if parent else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_k, sep).items())
        else:
            items.append((new_k, v))
    return dict(items)

def separate_summary_items(parsed):
    """Separate summary fields and line items."""
    flat = {}
    items_df = pd.DataFrame()

    if not isinstance(parsed, dict):
        return pd.DataFrame([{"Field": "data", "Value": str(parsed)}]), pd.DataFrame()

    for k, v in parsed.items():
        if isinstance(v, list) and all(isinstance(i, dict) for i in v):
            items_df = pd.concat([items_df, pd.DataFrame(v)], ignore_index=True)
        elif isinstance(v, dict):
            flat.update(flatten_dict(v, k))
        else:
            flat[k] = v

    df_summary = pd.DataFrame(list(flat.items()), columns=["Field", "Value"])
    return df_summary, items_df

def deep_get(data, path, default="Not Found"):
    keys = path.split(".")
    for k in keys:
        if isinstance(data, dict) and k in data:
            data = data[k]
        else:
            return default
    return data

# 🆕 JSON extraction helper
def extract_json_like_data(df):
    """Detect JSON strings in Value column and extract as key-value rows."""
    json_rows = []
    for _, row in df.iterrows():
        val = str(row["Value"]).strip()
        if val.startswith("{") and val.endswith("}"):
            try:
                parsed = json.loads(val)
                for k, v in parsed.items():
                    json_rows.append({"Parent_Field": row["Field"], "Key": k, "Value": v})
            except Exception:
                continue
    if json_rows:
        return pd.DataFrame(json_rows)
    return pd.DataFrame()

# 🆕 modified make_excel to include optional JSON extracted sheet
def make_excel(df_summary, df_items):
    buf = BytesIO()
    json_df = extract_json_like_data(df_summary)  # 🆕 extract JSON data
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Summary")
        if not df_items.empty:
            df_items.to_excel(writer, index=False, sheet_name="Items")
        # 🆕 Add JSON Extracted Sheet if exists
        if not json_df.empty:
            json_df.to_excel(writer, index=False, sheet_name="JSON_Extracted")
    buf.seek(0)
    return buf

def make_zip(pdf_bytes, excel_bytes, pdf_name, excel_name):
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(pdf_name, pdf_bytes)
        z.writestr(excel_name, excel_bytes.getvalue())
    zip_buf.seek(0)
    return zip_buf

# ========= PROMPT =========
default_prompt = """
SYSTEM:
You are a professional invoice data extractor. Return valid JSON only — no markdown, no text.
If a field is missing, write "Not Found".
Dates must always be ISO format (YYYY-MM-DD).

USER:
Extract all invoice-level, buyer, and seller fields, totals, taxes, and line items from the document images.
Ensure `invoice_total` refers to the actual INVOICE TOTAL (not Amount Due).
If no payment due date is found, just omit it.
If the invoice contains multiple site addresses or locations, include them under a `sites` array, where each site contains `address` and `date`.

Return JSON like:
{
  "billing_details": {"payable_to": "<vendor or 'Not Found'>"},
  "invoice": {
    "invoice_number": "<str>",
    "invoice_date": "<YYYY-MM-DD>",
    "invoice_total": "<invoice total amount>"
  },
  "seller": {"name": "<str>", "address": "<str>"},
  "buyer": {"name": "<str>", "address": "<str>"},
  "payment_due_by": "<YYYY-MM-DD>",
  "currency": "<str>",
  "sites": [{"address": "<str>", "date": "<YYYY-MM-DD>"}],
  "line_items": [{"description": "<str>", "quantity": "<num>", "unit_price": "<num>", "line_total": "<num>"}]
}
Return only valid JSON.
"""

# ========= MAIN LOGIC =========
if uploaded_pdf:
    pdf_bytes = uploaded_pdf.read()
    pdf_name = uploaded_pdf.name

    # PDF Preview
    st.markdown(
        f'<iframe src="data:application/pdf;base64,{base64.b64encode(pdf_bytes).decode()}" width="100%" height="600"></iframe>',
        unsafe_allow_html=True,
    )

    images = pdf_to_images(pdf_bytes)

    if st.button("⚙️ Generate Output"):
        try:
            img_entries = [{"type": "image_url", "image_url": {"url": encode_img_b64(i)}} for i in images]
            with st.spinner("🧠 Extracting data using GPT-4.1-mini..."):
                resp = client.chat.completions.create(
                    model="gpt-4.1-mini",
                    temperature=0,
                    messages=[
                        {"role": "system", "content": "Return valid JSON only."},
                        {"role": "user", "content": [{"type": "text", "text": default_prompt}, *img_entries]},
                    ],
                )

            raw_out = resp.choices[0].message.content
            parsed = clean_json_output(raw_out)
            st.session_state["parsed_data"] = parsed

            df_summary, df_items = separate_summary_items(parsed)
            st.session_state["df_summary"], st.session_state["items_df"] = df_summary, df_items

            # ========= DISPLAY =========
            st.subheader("📋 Extracted Summary & Items")
            st.dataframe(df_summary, height=250, use_container_width=True)
            if not df_items.empty:
                st.dataframe(df_items, height=250, use_container_width=True)

            excel_bytes = make_excel(df_summary, df_items)
            zip_buf = make_zip(pdf_bytes, excel_bytes, pdf_name, f"{pdf_name.split('.')[0]}_alldata.xlsx")

            st.download_button("⬇️ Download Excel (All Data)", excel_bytes,
                               f"{pdf_name.split('.')[0]}_alldata.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("🗜️ Download ZIP (PDF + All Data)", zip_buf,
                               f"{pdf_name.split('.')[0]}_bundle.zip",
                               mime="application/zip")

            # ========= TEMPLATE MAPPING =========
            st.markdown("---")
            st.subheader("🧩 Template Mapping")

            vendor = deep_get(parsed, "billing_details.payable_to")
            inv_no = deep_get(parsed, "invoice.invoice_number")
            inv_total = deep_get(parsed, "invoice.invoice_total")
            inv_date = deep_get(parsed, "invoice.invoice_date")
            due = parsed.get("payment_due_by", parsed.get("due_date", "Not Found"))
            sites = parsed.get("sites", [])

            # Build mapping records
            records = []
            if sites:
                for i, s in enumerate(sites, 1):
                    records.append({
                        "S.No": i,
                        "Memo #": "Not Found",
                        "Vendor Name": vendor,
                        "Address": s.get("address", "Not Found"),
                        "Inv #": inv_no,
                        "Inv Date": inv_date,
                        "Due Date": due,
                        "Amt": inv_total
                    })
            else:
                records.append({
                    "S.No": 1,
                    "Memo #": "Not Found",
                    "Vendor Name": vendor,
                    "Address": deep_get(parsed, "buyer.address"),
                    "Inv #": inv_no,
                    "Inv Date": inv_date,
                    "Due Date": due,
                    "Amt": inv_total
                })

            df_custom = pd.DataFrame(records)
            st.session_state["df_custom"] = df_custom

            # Template mapping logic
            if uploaded_template:
                try:
                    uploaded_template.seek(0)
                    df_template = pd.read_excel(uploaded_template)
                    template_fields = [c.strip() for c in df_template.columns]

                    filtered = []
                    for _, row in df_custom.iterrows():
                        row_dict = {col: row[col] if col in row else "Not Found" for col in template_fields}
                        filtered.append(row_dict)
                    df_mapping = pd.DataFrame(filtered)[template_fields]
                    st.success("✅ Template uploaded — columns mapped.")
                except Exception as e:
                    st.error(f"Template mapping failed: {e}")
                    df_mapping = df_custom
            else:
                st.warning("⚠️ No template uploaded — using default mapping columns.")
                df_mapping = df_custom[["S.No", "Memo #", "Vendor Name", "Address", "Inv #", "Inv Date", "Due Date", "Amt"]]

            st.dataframe(df_mapping, height=300, use_container_width=True)

            buf_map = BytesIO()
            with pd.ExcelWriter(buf_map, engine="openpyxl") as writer:
                df_mapping.to_excel(writer, index=False, sheet_name="Template_Mapping")
            buf_map.seek(0)

            zip_map = make_zip(pdf_bytes, buf_map, pdf_name, "template_mapping.xlsx")

            st.download_button("⬇️ Download Excel (Template Mapping)", buf_map,
                               f"{pdf_name.split('.')[0]}_template_mapping.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("🗜️ Download ZIP (PDF + Template Mapping Excel)", zip_map,
                               f"{pdf_name.split('.')[0]}_template_bundle.zip",
                               mime="application/zip")

        except Exception as e:
            st.error(f"❌ Error: {e}")