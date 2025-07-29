import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from openai import OpenAI
import fitz  # PyMuPDF
import os
from dotenv import load_dotenv
from io import BytesIO

# Load environment variables
load_dotenv()

# ---------------- DEFAULT CONFIG ----------------
TARGET_SECTIONS = [
    "business overview", "corporate information", "md&a", "principal product",
    "segment report", "business operations", "company overview", "companies affair",
    "company affair", "introduction", "background", "overview of the company",
    "overview of the business", "background information", "principal service"
]
#OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

if not OPENAI_API_KEY:
    st.error("OpenAI_API_KEY not found. Please set it as an environment variable.")
    st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)


# ---------------- Helper Functions ----------------

def get_internal_links(base_url, html_content):
    soup = BeautifulSoup(html_content, "lxml")
    links = set()
    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"]
        joined_url = urljoin(base_url, href)
        parsed_base = urlparse(base_url)
        parsed_url = urlparse(joined_url)
        if parsed_base.netloc == parsed_url.netloc:
            links.add(joined_url)
    return links


def crawl_website(start_url, max_pages=5, text_limit=15000):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/115.0 Safari/537.36"
        )
    }
    visited = set()
    to_visit = set([start_url])
    combined_text = ""
    crawl_log = []

    while to_visit and len(visited) < max_pages:
        url = to_visit.pop()
        if url in visited:
            continue
        try:
            response = requests.get(url, timeout=10, headers=headers)
            if response.status_code != 200:
                crawl_log.append(f"❌ Failed to fetch {url} - Status code: {response.status_code}")
                continue

            visited.add(url)
            soup = BeautifulSoup(response.text, "lxml")
            text = soup.get_text(separator=" ", strip=True)

            chars_extracted = len(text)
            crawl_log.append(f"✅ Extracted {chars_extracted} characters from {url}")

            combined_text += text[:5000] + "\n\n"
            new_links = get_internal_links(start_url, response.text)
            to_visit.update(new_links - visited)

        except Exception as e:
            crawl_log.append(f"⚠️ Exception fetching {url}: {e}")
            continue

    return combined_text[:text_limit], crawl_log


def extract_pdf_sections(pdf_file, text_limit=10000):
    try:
        doc = fitz.open(stream=pdf_file.read() if hasattr(pdf_file, "read") else pdf_file, filetype="pdf")
    except Exception as e:
        st.warning(f"Failed to open PDF: {e}")
        return "", []

    sections_text = ""
    pages_used = []

    for page_num, page in enumerate(doc, start=1):
        text = page.get_text()
        lower_text = text.lower()
        for keyword in TARGET_SECTIONS:
            if keyword in lower_text:
                sections_text += text + "\n\n"
                pages_used.append(page_num)
                break
        if len(sections_text) >= text_limit:
            break

    doc.close()
    return sections_text[:text_limit], pages_used


def analyze_text(text, source_name, model_name):
    prompt = (
        f"You are a financial analyst. Provide a consolidated **functional analysis** for a company based on the following content from {source_name}. "
        f"Summarize clearly within 75 words, focusing on functions performed. Avoid unrelated content.\n\n{text.strip()}"
    )
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"GPT analysis failed: {str(e)}"


# ---------------- Streamlit UI ----------------

st.set_page_config(page_title="FAR Analysis Tool", layout="wide")
st.title("🧾 Functional Analysis (FAR) Automation")

with st.expander("⚙️ Configuration", expanded=True):
    gpt_model = st.selectbox("Select GPT Model:", options=["gpt-3.5-turbo", "gpt-4", "gpt-4o"], index=0)
    max_pages = st.number_input("Max Website Pages to Crawl:", min_value=1, max_value=50, value=5)
    website_char_limit = st.number_input("Website Text Limit (chars):", min_value=1000, max_value=50000, value=15000)
    pdf_char_limit = st.number_input("PDF Text Limit (chars):", min_value=1000, max_value=30000, value=10000)

uploaded_excel = st.file_uploader("Select Input Excel File (.xlsx or .xls)", type=['xlsx', 'xls'])

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    if df.shape[1] < 2:
        st.error("Input Excel must have at least two columns: one for Company Name and another for Website URL.")
        st.stop()

    default_company_col = df.columns[0]
    default_website_col = df.columns[1]

    with st.expander("🔧 Column Mapping (Optional: Override Auto-detected Columns)", expanded=False):
        cols = df.columns.tolist()
        selected_company_col = st.selectbox("Select column for Company Name:", cols, index=0)
        selected_website_col = st.selectbox("Select column for Website URL:", cols, index=1)

    company_col = selected_company_col or default_company_col
    website_col = selected_website_col or default_website_col

    st.success(
        f"Loaded {len(df)} companies from Excel. Using '{company_col}' as Company Name and '{website_col}' as Website URL.")

    #company_col = df.columns[0]
    #website_col = df.columns[1]

    #st.success(
    #    f"Loaded {len(df)} companies from Excel. Using columns: '{company_col}' for Company Name and '{website_col}' for Website URL.")

    #expected_cols = {"Company Name", "Website URL"}
    #if not expected_cols.issubset(df.columns):
    #    st.error(f"Input Excel must contain columns: {expected_cols}")
    #    st.stop()

    #st.success(f"Loaded {len(df)} companies from Excel.")
    #st.markdown("---")
    #st.header("Company-wise Analysis")

    if "analysis_results" not in st.session_state:
        st.session_state.analysis_results = {}

    for idx, row in df.iterrows():
        company_name = row[company_col]
        website_url = row[website_col]
        company_key = f"company_{idx}"

        if company_key not in st.session_state:
            st.session_state[company_key] = {}

        st.subheader(f"{idx + 1}. {company_name}")

        pdf_file = st.file_uploader(f"Upload Annual Report PDF for {company_name} (optional)", type="pdf",
                                    key=f"pdf_{idx}")

        col1, col2 = st.columns(2)

        with col1:
            if "website_analysis" not in st.session_state[company_key]:
                st.write(f"🌐 Crawling website: {website_url} (max {max_pages} pages, {website_char_limit} chars limit)")
                website_text, crawl_log = crawl_website(website_url, max_pages=max_pages, text_limit=website_char_limit)
                if website_text:
                    st.write("💬 Running GPT analysis for Website...")
                    website_analysis = analyze_text(website_text, f"{company_name} Website", gpt_model)
                else:
                    website_analysis = "Could not extract website content."

                st.session_state[company_key]["website_analysis"] = website_analysis
                st.session_state[company_key]["crawl_log"] = crawl_log
            else:
                website_analysis = st.session_state[company_key]["website_analysis"]
                crawl_log = st.session_state[company_key].get("crawl_log", [])

            # Show crawl log
            with st.expander("🔍 Crawl Log (Page-wise Details)"):
                for log in crawl_log:
                    st.markdown(f"- {log}")

            st.markdown(f"**Website Analysis:** {website_analysis}")

        with col2:
            if pdf_file is not None:
                if st.session_state[company_key].get("last_pdf_filename") != pdf_file.name:
                    st.write(f"📄 Extracting PDF sections (up to {pdf_char_limit} chars)...")
                    pdf_text, pages_used = extract_pdf_sections(pdf_file, text_limit=pdf_char_limit)
                    if pdf_text:
                        st.write("💬 Running GPT analysis for Annual Report...")
                        pdf_analysis = analyze_text(pdf_text, f"{company_name} Annual Report", gpt_model)
                    else:
                        pdf_analysis = "No relevant sections found in PDF."
                    st.session_state[company_key]["pdf_analysis"] = pdf_analysis
                    st.session_state[company_key]["pages_used"] = pages_used
                    st.session_state[company_key]["last_pdf_filename"] = pdf_file.name
                else:
                    pdf_analysis = st.session_state[company_key].get("pdf_analysis", "")
                    pages_used = st.session_state[company_key].get("pages_used", [])

                pages_used_str = ", ".join(str(p) for p in pages_used)
                st.markdown(f"**Annual Report Analysis (Pages: {pages_used_str}):** {pdf_analysis}")
            else:
                st.info("No PDF uploaded; skipping Annual Report analysis.")
                pdf_analysis = ""
                pages_used_str = ""

        # Store result per company for export
        st.session_state.analysis_results[company_name] = {
            "Company Name": company_name,
            "Analysis as per Annual Report": pdf_analysis,
            "Pages in Annual Report": pages_used_str,
            "Analysis as per Website": website_analysis
        }

    # Export Button
    if st.button("📥 Export all analyses to Excel"):
        results = list(st.session_state.analysis_results.values())
        if results:
            out_df = pd.DataFrame(results)
            output_buffer = BytesIO()
            out_df.to_excel(output_buffer, index=False)
            output_buffer.seek(0)

            st.success("Excel file is ready for download.")
            st.download_button(
                label="Download FAR_Analysis.xlsx",
                data=output_buffer,
                file_name="FAR_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No analysis data available to export.")
else:
    st.info("Please upload an input Excel file to begin analysis.")
