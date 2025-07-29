import tkinter as tk
from tkinter import filedialog, scrolledtext
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from openai import OpenAI
import fitz  # PyMuPDF
import os
from dotenv import load_dotenv
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
    raise ValueError("OpenAI_API_KEY not found. Please set it as an environment variable.")
# ------------------------------------------------

client = OpenAI(api_key=OPENAI_API_KEY)

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
    visited = set()
    to_visit = set([start_url])
    combined_text = ""

    while to_visit and len(visited) < max_pages:
        url = to_visit.pop()
        if url in visited:
            continue
        try:
            response = requests.get(url, timeout=10)
            visited.add(url)

            soup = BeautifulSoup(response.text, "lxml")
            for tag in soup(["script", "style", "nav", "footer", "header", "form", "aside"]):
                tag.decompose()
            text = soup.get_text(separator=" ", strip=True)
            combined_text += text[:5000] + "\n\n"

            new_links = get_internal_links(start_url, response.text)
            to_visit.update(new_links - visited)
        except:
            continue

    return combined_text[:text_limit]

def extract_pdf_sections(pdf_path, text_limit=10000):
    doc = fitz.open(pdf_path)
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

# ---------------- UI & MAIN LOGIC ----------------
def run_analysis():
    status_text.delete(1.0, tk.END)

    gpt_model = gpt_model_var.get()
    max_pages = int(max_pages_entry.get())
    website_char_limit = int(website_char_limit_entry.get())
    pdf_char_limit = int(pdf_char_limit_entry.get())

    excel_path = filedialog.askopenfilename(title="Select Input Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not excel_path:
        status_text.insert(tk.END, "❌ No Excel file selected.\n")
        return

    output_folder = filedialog.askdirectory(title="Select Folder to Save Output File")
    if not output_folder:
        status_text.insert(tk.END, "❌ No output folder selected.\n")
        return

    df = pd.read_excel(excel_path)
    output_data = []

    for _, row in df.iterrows():
        company_name = row["Company Name"]
        website_url = row["Website URL"]

        status_text.insert(tk.END, f"\n🚀 Processing {company_name}...\n")
        status_text.see(tk.END)
        root.update()

        # Select PDF
        status_text.insert(tk.END, f"📄 Please select Annual Report PDF for {company_name}.\n")
        status_text.see(tk.END)
        root.update()

        pdf_file = filedialog.askopenfilename(title=f"Select PDF for {company_name}", filetypes=[("PDF files", "*.pdf")])
        if not pdf_file:
            status_text.insert(tk.END, f"⚠️ No PDF selected for {company_name}. Skipping PDF analysis.\n")
            pdf_analysis = ""
            pages_used_str = ""
        else:
            status_text.insert(tk.END, f"🔎 Extracting PDF sections...\n")
            status_text.see(tk.END)
            root.update()

            pdf_text, pages_used = extract_pdf_sections(pdf_file, text_limit=pdf_char_limit)
            if pdf_text:
                status_text.insert(tk.END, f"💬 Running GPT analysis for Annual Report...\n")
                status_text.see(tk.END)
                root.update()
                pdf_analysis = analyze_text(pdf_text, f"{company_name} Annual Report", gpt_model)
                pages_used_str = ", ".join(str(p) for p in pages_used)
            else:
                pdf_analysis = "No relevant sections found in PDF."
                pages_used_str = ""

        # Website
        status_text.insert(tk.END, f"🌐 Crawling website {website_url}...\n")
        status_text.see(tk.END)
        root.update()

        website_text = crawl_website(website_url, max_pages=max_pages, text_limit=website_char_limit)
        if website_text:
            status_text.insert(tk.END, f"💬 Running GPT analysis for Website...\n")
            status_text.see(tk.END)
            root.update()
            website_analysis = analyze_text(website_text, f"{company_name} Website", gpt_model)
        else:
            website_analysis = "Could not extract website content."

        output_data.append({
            "Company Name": company_name,
            "Analysis as per Annual Report": pdf_analysis,
            "Pages in Annual Report": pages_used_str,
            "Analysis as per Website": website_analysis
        })

        status_text.insert(tk.END, f"✅ Completed analysis for {company_name}.\n")
        status_text.see(tk.END)
        root.update()

    # Save Excel
    out_df = pd.DataFrame(output_data)
    output_file_path = os.path.join(output_folder, "FAR_Analysis.xlsx")
    out_df.to_excel(output_file_path, index=False)
    status_text.insert(tk.END, f"\n🏁 ✅ All analyses saved to {output_file_path}\n")
    status_text.see(tk.END)
    root.update()

# ---------------- TKINTER UI ----------------
# ---------------- TKINTER UI ----------------
root = tk.Tk()
root.title("FAR Analysis Tool")
root.geometry("900x700")
root.configure(bg="#f0f2f5")

# -------- Header Label --------
header = tk.Label(root, text="🧾 Functional Analysis (FAR) Automation", font=("Arial", 16, "bold"), bg="#f0f2f5", fg="#2c3e50")
header.pack(pady=(10, 5))

# -------- Configuration Frame --------
config_frame = tk.LabelFrame(root, text="Configuration", font=("Arial", 12, "bold"), padx=10, pady=10, bg="#ffffff")
config_frame.pack(padx=10, pady=10, fill="x")

# GPT Model dropdown
tk.Label(config_frame, text="Select GPT Model:", bg="#ffffff", anchor="w").grid(row=0, column=0, sticky="w", pady=5)
gpt_model_var = tk.StringVar(value="gpt-3.5-turbo")
gpt_dropdown = tk.OptionMenu(config_frame, gpt_model_var, "gpt-3.5-turbo", "gpt-4", "gpt-4o")
gpt_dropdown.config(width=20)
gpt_dropdown.grid(row=0, column=1, pady=5, padx=10, sticky="w")

# Max crawl pages
tk.Label(config_frame, text="Max Website Pages to Crawl:", bg="#ffffff", anchor="w").grid(row=1, column=0, sticky="w", pady=5)
max_pages_entry = tk.Entry(config_frame, width=25)
max_pages_entry.insert(0, "5")
max_pages_entry.grid(row=1, column=1, pady=5, padx=10, sticky="w")

# Website char limit
tk.Label(config_frame, text="Website Text Limit (chars):", bg="#ffffff", anchor="w").grid(row=2, column=0, sticky="w", pady=5)
website_char_limit_entry = tk.Entry(config_frame, width=25)
website_char_limit_entry.insert(0, "15000")
website_char_limit_entry.grid(row=2, column=1, pady=5, padx=10, sticky="w")

# PDF char limit
tk.Label(config_frame, text="PDF Text Limit (chars):", bg="#ffffff", anchor="w").grid(row=3, column=0, sticky="w", pady=5)
pdf_char_limit_entry = tk.Entry(config_frame, width=25)
pdf_char_limit_entry.insert(0, "10000")
pdf_char_limit_entry.grid(row=3, column=1, pady=5, padx=10, sticky="w")

# -------- Run Button --------
btn_run = tk.Button(root, text="▶ Start FAR Analysis", command=run_analysis, bg="#27ae60", fg="white",
                    font=("Arial", 12, "bold"), width=30, height=2)
btn_run.pack(pady=10)

# -------- Status Text --------
status_label = tk.Label(root, text="Status Log", font=("Arial", 12, "bold"), bg="#f0f2f5", anchor="w")
status_label.pack(anchor="w", padx=10)

status_text = scrolledtext.ScrolledText(root, width=110, height=25, wrap=tk.WORD, font=("Consolas", 10))
status_text.pack(padx=10, pady=(0, 10))


root.mainloop()