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
    "overview of the business", "background information"
]

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    st.error("OpenAI_API_KEY not found. Please set it as an environment variable.")
    st.stop()
# ------------------------------------------------

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

    while to_visit and len(visited) < max_pages:
        url = to_visit.pop()
        if url in visited:
            continue
        try:
            response = requests.get(url, timeout=10, headers=headers)
            if response.status_code != 200:
                st.warning(f"Failed to fetch {url} - Status code: {response.status_code}")
                continue

            visited.add(url)

            soup = BeautifulSoup(response.text, "lxml")

            # Uncomment below lines if you want to remove these tags while testing
            # for tag in soup(["script", "style", "nav", "footer", "header", "form", "aside"]):
            #     tag.decompose()

            text = soup.get_text(separator=" ", strip=True)

            st.info(f"Extracted {len(text)} characters from {url}")

            combined_text += text[:5000] + "\n\n"

            # Get internal links for crawling more pages
            new_links = get_internal_links(start_url, response.text)
            to_visit.update(new_links - visited)

        except Exception as e:
            st.warning(f"Exception fetching {url}: {e}")
            continue

    return combined_text[:text_limit]


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


def extract_pdf_sections(pdf_file, text_limit=10000):
    # pdf_file can be a filename or BytesIO object
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
        # Stop if we exceed text_limit early
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

    # Validate expected columns
    expected_cols = {"Company Name", "Website URL"}
    if not expected_cols.issubset(df.columns):
        st.error(f"Input Excel must contain columns: {expected_cols}")
        st.stop()

    st.success(f"Loaded {len(df)} companies from Excel.")

    # Prepare output list
    output_data = []

    st.markdown("---")
    st.header("Company-wise Analysis")

    # Use session state to store results to allow rerun preservation
    if "analysis_results" not in st.session_state:
        st.session_state.analysis_results = {}

    # For each company
    for idx, row in df.iterrows():
        company_name = row["Company Name"]
        website_url = row["Website URL"]

        st.subheader(f"{idx + 1}. {company_name}")

        pdf_file = st.file_uploader(f"Upload Annual Report PDF for {company_name} (optional)", type="pdf",
                                    key=f"pdf_{idx}")

        col1, col2 = st.columns(2)

        with col1:
            st.write(f"🌐 Crawling website: {website_url} (max {max_pages} pages, {website_char_limit} chars limit)")
            website_text = crawl_website(website_url, max_pages=max_pages, text_limit=website_char_limit)
            if website_text:
                st.write("💬 Running GPT analysis for Website...")
                website_analysis = analyze_text(website_text, f"{company_name} Website", gpt_model)
            else:
                website_analysis = "Could not extract website content."
            st.markdown(f"**Website Analysis:** {website_analysis}")

        with col2:
            if pdf_file is not None:
                st.write(f"📄 Extracting PDF sections (up to {pdf_char_limit} chars)...")
                pdf_text, pages_used = extract_pdf_sections(pdf_file, text_limit=pdf_char_limit)
                if pdf_text:
                    st.write("💬 Running GPT analysis for Annual Report...")
                    pdf_analysis = analyze_text(pdf_text, f"{company_name} Annual Report", gpt_model)
                else:
                    pdf_analysis = "No relevant sections found in PDF."
                pages_used_str = ", ".join(str(p) for p in pages_used)
                st.markdown(f"**Annual Report Analysis (Pages: {pages_used_str}):** {pdf_analysis}")
            else:
                st.info("No PDF uploaded; skipping Annual Report analysis.")
                pdf_analysis = ""
                pages_used_str = ""

        # Save current company's data in session state for export
        st.session_state.analysis_results[company_name] = {
            "Company Name": company_name,
            "Analysis as per Annual Report": pdf_analysis,
            "Pages in Annual Report": pages_used_str,
            "Analysis as per Website": website_analysis
        }

    # After process button
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
