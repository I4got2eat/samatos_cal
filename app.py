import streamlit as st
import fitz
import base64
import json
import pandas as pd
import re
from anthropic import Anthropic

def extract_specific_pages_as_images(pdf_bytes, start_page, end_page, dpi=300):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    target_images = []
    actual_start = max(0, start_page - 1)
    actual_end = min(len(doc), end_page)

    for page_num in range(actual_start, actual_end):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=dpi)
        image_bytes = pix.tobytes("jpeg")
        base64_image = base64.b64encode(image_bytes).decode('utf-8')
        target_images.append({
            "page_num": page_num + 1,
            "base64_image": base64_image
        })

    doc.close()
    return target_images

def extract_table_data(base64_image):
    client = Anthropic(api_key="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")
    prompt = """
    Extract the table content from this image. 
    1. Analyze the document to identify hierarchical headers (Main Categories and Sub Categories).
    2. Flatten the table structure into a list of objects.
    3. For every row, identify the 'Current_Main_Section' and 'Current_Sub_Section' that applies to that data.
    4. Include all available columns in the row (e.g., Eil_Nr, Pavadinimas, Vnt, Kiekis, TS_skyrius, Pastabos).
    5. Return ONLY a JSON object with a single key 'table_data' containing the array of objects.
    
    Rules:
    - If a row is a section header, do not include it as a data row; use it to update the context for subsequent rows.
    - If a value is missing or spanned, infer it from the context of the identified section.
    - Do not use markdown, do not add conversational text.
    """
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        temperature=0,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/jpeg",
                            "data": base64_image
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt
                    }
                ]
            }
        ]
    )
    
    raw_text = response.content[0].text.strip()
    match = re.search(r'\{.*\}', raw_text, re.DOTALL)
    clean_json = match.group(0) if match else raw_text
        
    return json.loads(clean_json)

st.set_page_config(page_title="PDF Table Extractor", layout="wide")
st.title("PDF Table Extractor (Claude Powered)")

uploaded_file = st.file_uploader("Upload PDF Project Specs", type=["pdf"])

col1, col2 = st.columns(2)
with col1:
    start_page = st.number_input("Start Page", min_value=1, value=1)
with col2:
    end_page = st.number_input("End Page", min_value=1, value=1)

if st.button("Extract Tables"):
    if uploaded_file is not None and start_page <= end_page:
        with st.spinner("Processing..."):
            pdf_bytes = uploaded_file.read()
            images = extract_specific_pages_as_images(pdf_bytes, start_page, end_page)
            
            all_data = []
            for img_dict in images:
                try:
                    result = extract_table_data(img_dict['base64_image'])
                    if "table_data" in result:
                        all_data.extend(result["table_data"])
                except Exception as e:
                    st.error(f"Error on page {img_dict['page_num']}: {e}")

            if all_data:
                df = pd.DataFrame(all_data)
                st.dataframe(df, use_container_width=True)
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button("Download CSV", csv, "extracted_tables.csv", "text/csv")