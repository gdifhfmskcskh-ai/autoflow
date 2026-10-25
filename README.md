# ⚡ AutoFlow Studio - Cloud Automation Suite

> **Rebranded, Modernized Python Automation Platform**  
> Optimized for 100% Free Hosting on **Streamlit Community Cloud** and **Vercel**.

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/)
[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## 🌟 Features Included

1. **📊 Excel Pivot & Chart Automator** (`openpyxl`, `pandas`):
   - Generates multi-dimensional pivot tables.
   - Calculates dynamic summary rows with `=SUM()` formulas.
   - Embeds native Excel BarCharts with professional corporate styling.
2. **📑 Web & Document Table Extractor** (`pandas`, `requests`, `lxml`):
   - Extracts tables from HTML files, web URLs, or raw HTML snippets.
   - Exports directly to CSV, Excel, or JSON.
3. **🧾 Smart Invoice & Document Generator** (`jinja2`, `html/css`):
   - Clean, zero-binary HTML/PDF invoice generation.
   - Works on any serverless or containerized environment without `wkhtmltopdf` errors.
4. **📁 Cloud File Organizer & Hash Deduplicator** (`hashlib`):
   - Categorizes files into Images, Documents, Datasets, Archives.
   - Detects duplicate files using MD5 cryptographic checksums.
5. **✉️ Bulk Email Dispatcher & Templating Engine**:
   - Environment-variable-based SMTP configuration (`SMTP_USER`, `SMTP_PASS`).
   - Dynamic tag substitution (`{{Name}}`, `{{Company}}`).
6. **🌐 Automated News Scraper**:
   - Headless, cloud-compatible HTTP scraping without requiring local Chrome GUI or chromedriver.

---

## 🚀 Free Hosting Instructions (ফ্রি হোস্টিং নির্দেশিকা)

### 🎈 Option 1: Streamlit Community Cloud (Recommended / সবচেয়ে সহজ)
Streamlit Cloud offers free 24/7 cloud hosting for Python apps directly connected to your GitHub repository.

1. **Step 1 - Fork or Push to GitHub**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit of AutoFlow Studio"
   git remote add origin https://github.com/YOUR_USERNAME/autoflow-studio.git
   git branch -M main
   git push -u origin main
   ```
2. **Step 2 - Connect to Streamlit**:
   - Go to [share.streamlit.io](https://share.streamlit.io/)
   - Log in with your GitHub account.
   - Click **"New app"**.
3. **Step 3 - Select Repository & Deploy**:
   - **Repository:** `YOUR_USERNAME/autoflow-studio`
   - **Branch:** `main`
   - **Main file path:** `streamlit_app.py`
   - Click **"Deploy!"**
4. Your application will be live with a permanent, secure HTTPS URL!

---

### ▲ Option 2: Vercel Serverless Hosting
Vercel provides instant global CDN and serverless Python/Node execution.

1. Install the Vercel CLI or connect via GitHub:
   ```bash
   npm i -g vercel
   vercel
   ```
2. Or in the [Vercel Dashboard](https://vercel.com/new):
   - Import your GitHub repository.
   - Vercel automatically reads `vercel.json` and provisions the endpoints.
   - Click **Deploy**.

---

### 🇧🇩 বাংলা নির্দেশিকা (Bengali Guide)

**স্ক্রিপ্টগুলো ফ্রি হোস্টিংয়ে কিভাবে রান করবেন?**
1. **স্ট্রিমলাইট (Streamlit Cloud):**
   - GitHub এ একটি নতুন পাবলিক রিপোজিটরি তৈরি করে কোডগুলো পুশ করুন।
   - [share.streamlit.io](https://share.streamlit.io) এ গিয়ে GitHub দিয়ে সাইন-ইন করুন।
   - "New App" সিলেক্ট করে আপনার রিপো এবং `streamlit_app.py` দিন। "Deploy" চাপলেই সম্পূর্ণ ফ্রি ক্লাউড হোস্টিং চালু হবে।
2. **ভারসেল (Vercel):**
   - Vercel.com এ ফ্রি একাউন্ট দিয়ে লগইন করুন এবং আপনার GitHub রিপোজিটরি ইমপোর্ট করে এক ক্লিকে ডিপ্লয় করুন।

---

## 🛠 Local Setup & Running

```bash
# 1. Clone repository
git clone https://github.com/YOUR_USERNAME/autoflow-studio.git
cd autoflow-studio

# 2. Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# 3. Install requirements
pip install -r requirements.txt

# 4. Run Streamlit app locally
streamlit run streamlit_app.py
```

---

## 📄 License
MIT License. Free for personal and commercial usage.
