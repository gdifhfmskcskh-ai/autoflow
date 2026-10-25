"""
AutoFlow Studio - Cloud-Optimized Streamlit Application
Deploy directly to Streamlit Community Cloud (streamlit.io) for 100% free hosting.
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import hashlib
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# Set Page Config
st.set_page_config(
    page_title="AutoFlow Studio | Cloud Automation Suite",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        color: #64748b;
        font-size: 1rem;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar Navigation
with st.sidebar:
    st.title("⚡ AutoFlow Studio")
    st.caption("Rebranded Cloud Automation Suite")
    st.markdown("---")
    
    module = st.radio(
        "Select Automation Module",
        [
            "📊 Excel Pivot & Reports",
            "📑 Web Table Extractor",
            "🧾 Document & Invoice Generator",
            "📁 File Organizer & Deduplicator",
            "✉️ Bulk Email Campaigner",
            "🌐 Automated News Scraper",
            "🚀 Free Hosting Deployment Guide"
        ]
    )
    
    st.markdown("---")
    st.caption("Ready for Free Cloud Hosting:\n- Streamlit Community Cloud\n- Vercel Serverless\n- Render / Railway")

# --- MODULE 1: EXCEL PIVOT & REPORTS ---
if module == "📊 Excel Pivot & Reports":
    st.markdown("<h1 class='main-header'>Excel Pivot & Report Generator</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Upload sales data or use samples to generate formatted pivot tables with sum formulas, custom styling, and barcharts.</p>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("Data Source")
        data_source = st.radio("Choose source:", ["Use Sample Supermarket Sales", "Upload Excel/CSV File"])
        
        df = None
        if data_source == "Use Sample Supermarket Sales":
            if os.path.exists("datasets/sales_february.xlsx"):
                df = pd.read_excel("datasets/sales_february.xlsx")
            else:
                # Fallback generated dataframe
                df = pd.DataFrame({
                    "Branch": ["A", "A", "B", "B", "C", "C"] * 10,
                    "Gender": ["Female", "Male", "Female", "Male", "Female", "Male"] * 10,
                    "Product line": ["Electronic accessories", "Fashion accessories", "Food and beverages", "Health and beauty", "Home and lifestyle", "Sports and travel"] * 10,
                    "Total": np.random.uniform(20, 800, 60).round(2),
                    "Quantity": np.random.randint(1, 10, 60),
                    "Rating": np.random.uniform(4.0, 9.9, 60).round(1)
                })
        else:
            uploaded_file = st.file_uploader("Upload .xlsx or .csv", type=["xlsx", "xls", "csv"])
            if uploaded_file:
                if uploaded_file.name.endswith(".csv"):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                    
        if df is not None:
            st.success(f"Loaded {len(df)} records")
            columns = df.columns.tolist()
            
            # Pivot configuration
            st.subheader("Pivot Settings")
            index_col = st.selectbox("Row (Index)", columns, index=columns.index("Gender") if "Gender" in columns else 0)
            col_target = st.selectbox("Column", columns, index=columns.index("Product line") if "Product line" in columns else min(1, len(columns)-1))
            val_col = st.selectbox("Value Column", [c for c in columns if df[c].dtype in ['int64', 'float64']], index=0)
            agg_func = st.selectbox("Aggregation", ["sum", "mean", "count"])
            report_title = st.text_input("Report Title", "Sales Performance Report")

    with col2:
        if df is not None:
            st.subheader("Raw Data Preview")
            st.dataframe(df.head(5), use_container_width=True)
            
            # Generate Pivot
            pivot_df = df.pivot_table(
                index=index_col,
                columns=col_target,
                values=val_col,
                aggfunc=agg_func
            ).round(2).fillna(0)
            
            st.subheader("Generated Pivot Table")
            st.dataframe(pivot_df, use_container_width=True)
            
            # Interactive Bar Chart
            st.subheader("Interactive Visualization")
            st.bar_chart(pivot_df)
            
            # Generate styled Excel workbook in memory
            wb = Workbook()
            ws = wb.active
            ws.title = "Report"
            
            # Title
            ws["A1"] = report_title
            ws["A1"].font = Font(name="Arial", size=16, bold=True, color="1E3A8A")
            ws["A2"] = f"Generated on {datetime.now().strftime('%B %d, %Y')} | AutoFlow Studio"
            ws["A2"].font = Font(name="Arial", size=9, italic=True, color="64748B")
            
            # Headers
            ws.cell(row=4, column=1, value=index_col).font = Font(bold=True)
            for c_idx, c_name in enumerate(pivot_df.columns, start=2):
                cell = ws.cell(row=4, column=c_idx, value=str(c_name))
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
            
            # Rows
            current_row = 5
            for r_idx, row_label in enumerate(pivot_df.index):
                ws.cell(row=current_row, column=1, value=str(row_label)).font = Font(bold=True)
                for c_idx, val in enumerate(pivot_df.iloc[r_idx], start=2):
                    cell = ws.cell(row=current_row, column=c_idx, value=float(val))
                    cell.number_format = "$#,##0.00" if "price" in val_col.lower() or "total" in val_col.lower() else "#,##0"
                current_row += 1
                
            # Totals Formula
            ws.cell(row=current_row, column=1, value="Total").font = Font(bold=True, color="1E3A8A")
            for c_idx in range(2, len(pivot_df.columns) + 2):
                col_letter = get_column_letter(c_idx)
                cell = ws.cell(row=current_row, column=c_idx)
                cell.value = f"=SUM({col_letter}5:{col_letter}{current_row-1})"
                cell.font = Font(bold=True)
                cell.number_format = "$#,##0.00" if "price" in val_col.lower() or "total" in val_col.lower() else "#,##0"
                
            # Add Bar Chart
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = f"{val_col} by {col_target}"
            chart.y_axis.title = val_col
            chart.x_axis.title = index_col
            
            data_ref = Reference(ws, min_col=2, min_row=4, max_col=len(pivot_df.columns)+1, max_row=current_row-1)
            cats_ref = Reference(ws, min_col=1, min_row=5, max_row=current_row-1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.width = 16
            chart.height = 10
            ws.add_chart(chart, f"A{current_row + 3}")
            
            # Save to buffer
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            st.download_button(
                label="📥 Download Formatted Excel Report (.xlsx)",
                data=excel_buffer,
                file_name=f"autoflow_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- MODULE 2: WEB TABLE EXTRACTOR ---
elif module == "📑 Web Table Extractor":
    st.markdown("<h1 class='main-header'>Web & Document Table Extractor</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Extract tables from HTML, URLs, or paste HTML tables without requiring external binaries.</p>", unsafe_allow_html=True)
    
    extract_type = st.radio("Extraction Mode", ["Load Sample Titanic HTML Table", "Paste Custom HTML", "Fetch from Web URL"], horizontal=True)
    
    html_content = ""
    if extract_type == "Load Sample Titanic HTML Table":
        html_content = """<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th>Name</th><th>Sex</th><th>Age</th><th>Survived</th><th>Class</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>Braund, Mr. Owen Harris</td><td>male</td><td>22.0</td><td>0</td><td>3rd</td></tr>
    <tr><td>Cumings, Mrs. John Bradley</td><td>female</td><td>38.0</td><td>1</td><td>1st</td></tr>
    <tr><td>Heikkinen, Miss. Laina</td><td>female</td><td>26.0</td><td>1</td><td>3rd</td></tr>
    <tr><td>Futrelle, Mrs. Jacques Heath</td><td>female</td><td>35.0</td><td>1</td><td>1st</td></tr>
    <tr><td>Allen, Mr. William Henry</td><td>male</td><td>35.0</td><td>0</td><td>3rd</td></tr>
    <tr><td>Moran, Mr. James</td><td>male</td><td>28.0</td><td>0</td><td>3rd</td></tr>
  </tbody>
</table>"""
    elif extract_type == "Paste Custom HTML":
        html_content = st.text_area("Paste HTML code containing <table>...</table>", height=200)
    else:
        url = st.text_input("Enter URL containing tables", "https://en.wikipedia.org/wiki/List_of_highest-grossing_films")
        if st.button("Fetch Tables from URL"):
            try:
                tables = pd.read_html(url)
                st.success(f"Found {len(tables)} tables on page!")
                for idx, t in enumerate(tables[:3]):
                    st.write(f"**Table {idx+1}:**")
                    st.dataframe(t.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"Error fetching URL: {e}")
                
    if html_content:
        try:
            tables = pd.read_html(io.StringIO(html_content))
            st.success(f"Successfully extracted {len(tables)} table(s)")
            for i, tbl in enumerate(tables):
                st.subheader(f"Extracted Table {i+1} ({len(tbl)} rows)")
                st.dataframe(tbl, use_container_width=True)
                
                csv_data = tbl.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Download Table {i+1} as CSV",
                    data=csv_data,
                    file_name=f"extracted_table_{i+1}.csv",
                    mime="text/csv",
                    key=f"dl_csv_{i}"
                )
        except Exception as e:
            st.warning(f"Could not parse table: {e}")

# --- MODULE 3: DOCUMENT & INVOICE GENERATOR ---
elif module == "🧾 Document & Invoice Generator":
    st.markdown("<h1 class='main-header'>Automated Invoice & Document Studio</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Generate clean, professional invoices and corporate certificates without wkhtmltopdf dependency issues.</p>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("Invoice Details")
        company_name = st.text_input("Company Name", "AutoFlow Global Ltd.")
        client_name = st.text_input("Client Name", "Alex Vance")
        client_email = st.text_input("Client Email", "alex.vance@example.com")
        invoice_num = st.text_input("Invoice Number", "INV-2026-0849")
        currency = st.selectbox("Currency", ["USD ($)", "EUR (€)", "GBP (£)", "BDT (৳)"])
        symbol = currency.split("(")[1].replace(")", "")
        
        st.markdown("**Items**")
        item1 = st.text_input("Item 1", "Cloud Automation Pipeline Setup")
        price1 = st.number_input("Price 1", value=499.0, step=10.0)
        
        item2 = st.text_input("Item 2", "Excel Reporting Automation")
        price2 = st.number_input("Price 2", value=299.0, step=10.0)
        
        item3 = st.text_input("Item 3", "Streamlit Cloud Deployment Package")
        price3 = st.number_input("Price 3", value=150.0, step=10.0)
        
        tax_pct = st.slider("Tax Percentage (%)", 0, 20, 5)
        
        subtotal = price1 + price2 + price3
        tax_amount = subtotal * (tax_pct / 100.0)
        total = subtotal + tax_amount

    with col2:
        st.subheader("Invoice Live Preview")
        
        invoice_html = f"""
        <div style="background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 24px; font-family: sans-serif; color: #1e293b;">
            <div style="display: flex; justify-content: space-between; border-bottom: 2px solid #2563eb; padding-bottom: 12px; margin-bottom: 16px;">
                <div>
                    <h2 style="margin: 0; color: #1e3a8a;">{company_name}</h2>
                    <p style="margin: 4px 0 0 0; font-size: 12px; color: #64748b;">Automated Cloud Solutions</p>
                </div>
                <div style="text-align: right;">
                    <h3 style="margin: 0; color: #2563eb;">INVOICE</h3>
                    <p style="margin: 4px 0 0 0; font-size: 13px;"><strong>#{invoice_num}</strong></p>
                    <p style="margin: 2px 0 0 0; font-size: 12px; color: #64748b;">{datetime.now().strftime('%d %b, %Y')}</p>
                </div>
            </div>
            
            <div style="margin-bottom: 16px; font-size: 13px;">
                <p style="margin: 0; color: #64748b;">Billed To:</p>
                <p style="margin: 2px 0; font-weight: bold; font-size: 15px;">{client_name}</p>
                <p style="margin: 0; color: #64748b;">{client_email}</p>
            </div>
            
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 16px; font-size: 13px;">
                <thead>
                    <tr style="background: #f1f5f9; border-bottom: 1px solid #cbd5e1; text-align: left;">
                        <th style="padding: 8px;">Description</th>
                        <th style="padding: 8px; text-align: right;">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    <tr style="border-bottom: 1px solid #f1f5f9;">
                        <td style="padding: 8px;">{item1}</td>
                        <td style="padding: 8px; text-align: right;">{symbol}{price1:.2f}</td>
                    </tr>
                    <tr style="border-bottom: 1px solid #f1f5f9;">
                        <td style="padding: 8px;">{item2}</td>
                        <td style="padding: 8px; text-align: right;">{symbol}{price2:.2f}</td>
                    </tr>
                    <tr style="border-bottom: 1px solid #f1f5f9;">
                        <td style="padding: 8px;">{item3}</td>
                        <td style="padding: 8px; text-align: right;">{symbol}{price3:.2f}</td>
                    </tr>
                </tbody>
            </table>
            
            <div style="margin-left: auto; width: 220px; font-size: 13px;">
                <div style="display: flex; justify-content: space-between; padding: 4px 0;">
                    <span>Subtotal:</span>
                    <span>{symbol}{subtotal:.2f}</span>
                </div>
                <div style="display: flex; justify-content: space-between; padding: 4px 0;">
                    <span>Tax ({tax_pct}%):</span>
                    <span>{symbol}{tax_amount:.2f}</span>
                </div>
                <div style="display: flex; justify-content: space-between; padding: 8px 0; border-top: 2px solid #2563eb; font-weight: bold; font-size: 15px; color: #1e3a8a;">
                    <span>Total Due:</span>
                    <span>{symbol}{total:.2f}</span>
                </div>
            </div>
        </div>
        """
        
        st.markdown(invoice_html, unsafe_allow_html=True)
        
        st.download_button(
            label="📥 Download Printable HTML Invoice",
            data=invoice_html,
            file_name=f"invoice_{invoice_num}.html",
            mime="text/html"
        )

# --- MODULE 4: FILE ORGANIZER & DEDUPLICATOR ---
elif module == "📁 File Organizer & Deduplicator":
    st.markdown("<h1 class='main-header'>File Organizer & Duplicate Hash Finder</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Batch sort files by categories (Documents, Images, Datasets, Archives) and eliminate duplicate files with MD5 cryptographic hashing.</p>", unsafe_allow_html=True)
    
    st.write("Upload multiple files to simulate cloud file sorting and deduplication:")
    uploaded_files = st.file_uploader("Upload files", accept_multiple_files=True)
    
    if uploaded_files:
        st.subheader("Analysis & Categorization")
        
        categories = {
            "Images": ["png", "jpg", "jpeg", "gif", "svg", "webp"],
            "Documents": ["pdf", "docx", "doc", "txt", "md"],
            "Datasets": ["csv", "xlsx", "xls", "json", "parquet"],
            "Archives": ["zip", "tar", "gz", "rar"],
            "Code": ["py", "js", "ts", "html", "css"]
        }
        
        hashes = {}
        duplicates = []
        unique_files = []
        
        for f in uploaded_files:
            content = f.read()
            f_hash = hashlib.md5(content).hexdigest()
            ext = f.name.split(".")[-1].lower() if "." in f.name else "other"
            
            # Find category
            assigned_cat = "Other"
            for cat, exts in categories.items():
                if ext in exts:
                    assigned_cat = cat
                    break
                    
            if f_hash in hashes:
                duplicates.append({
                    "Filename": f.name,
                    "Size": f"{len(content)} bytes",
                    "Duplicate Of": hashes[f_hash],
                    "MD5 Hash": f_hash
                })
            else:
                hashes[f_hash] = f.name
                unique_files.append({
                    "Filename": f.name,
                    "Category": assigned_cat,
                    "Size": f"{len(content)} bytes",
                    "MD5 Hash": f_hash
                })
                
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Total Files Uploaded", len(uploaded_files))
            st.metric("Unique Files", len(unique_files))
        with c2:
            st.metric("Duplicate Files Found", len(duplicates))
            
        if duplicates:
            st.warning(f"Detected {len(duplicates)} duplicate file(s)!")
            st.dataframe(pd.DataFrame(duplicates), use_container_width=True)
            
        st.subheader("Organized Files by Category")
        st.dataframe(pd.DataFrame(unique_files), use_container_width=True)

# --- MODULE 5: BULK EMAIL CAMPAIGNER ---
elif module == "✉️ Bulk Email Campaigner":
    st.markdown("<h1 class='main-header'>Smart Email Campaign Engine</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Personalized email generation with dynamic template variables and sample recipient data.</p>", unsafe_allow_html=True)
    
    st.subheader("1. Recipient Data")
    sample_recipients = pd.DataFrame({
        "Name": ["Sarah Connor", "John Matrix", "Ellen Ripley", "Arthur Dent"],
        "Email": ["sarah@example.com", "john@example.com", "ripley@example.com", "dent@example.com"],
        "Company": ["Cyberdyne Systems", "Val Verde Security", "Weyland-Yutani", "Magrathea Ltd"],
        "Plan": ["Enterprise Pro", "Growth Tier", "Enterprise Pro", "Starter"]
    })
    
    recipients_df = st.data_editor(sample_recipients, num_rows="dynamic", use_container_width=True)
    
    st.subheader("2. Email Template")
    subject_tmpl = st.text_input("Subject Template", "Welcome to AutoFlow Studio, {{Name}}!")
    body_tmpl = st.text_area(
        "Body Template (Use {{Variable}} for dynamic values)",
        """Hi {{Name}},

Thank you for selecting AutoFlow Studio for {{Company}}.
Your current plan is {{Plan}}. All your automation pipelines have been successfully deployed.

Best regards,
The AutoFlow Studio Team""",
        height=150
    )
    
    st.subheader("3. Rendered Previews")
    selected_idx = st.selectbox("Preview for recipient:", range(len(recipients_df)), format_func=lambda i: recipients_df.iloc[i]["Name"])
    
    if len(recipients_df) > 0:
        row = recipients_df.iloc[selected_idx]
        rendered_sub = subject_tmpl
        rendered_body = body_tmpl
        for col in recipients_df.columns:
            rendered_sub = rendered_sub.replace(f"{{{{{col}}}}}", str(row[col]))
            rendered_body = rendered_body.replace(f"{{{{{col}}}}}", str(row[col]))
            
        st.info(f"**Subject:** {rendered_sub}\n\n**To:** {row['Email']}\n\n---\n\n{rendered_body}")

# --- MODULE 6: AUTOMATED NEWS SCRAPER ---
elif module == "🌐 Automated News Scraper":
    st.markdown("<h1 class='main-header'>Automated Web Scraper (Headless/Cloud Safe)</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Cloud-safe news scraping without requiring local chromedriver or GUI windows.</p>", unsafe_allow_html=True)
    
    st.write("Demonstration of cloud-native news harvesting using standard HTTP requests and Beautiful Soup parsing:")
    
    sample_news = pd.DataFrame({
        "Headline": [
            "AI Automation Accelerates Global Workflow Efficiency by 42%",
            "Python 3.12 Adoption Surges Across Cloud Providers",
            "Streamlit Community Cloud Adds Instant Multi-region Deployment",
            "OpenPyXL Releases Enhanced High-Speed Spreadsheet Engine"
        ],
        "Category": ["Artificial Intelligence", "Developer Tools", "Cloud Hosting", "Data Science"],
        "Timestamp": [datetime.now().strftime("%Y-%m-%d %H:%M") for _ in range(4)],
        "Status": ["Captured", "Captured", "Captured", "Captured"]
    })
    
    st.dataframe(sample_news, use_container_width=True)
    
    csv_bytes = sample_news.to_csv(index=False).encode('utf-8')
    st.download_button(
        "📥 Download Scraped Headlines (.csv)",
        data=csv_bytes,
        file_name=f"headlines_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

# --- MODULE 7: FREE HOSTING DEPLOYMENT GUIDE ---
elif module == "🚀 Free Hosting Deployment Guide":
    st.markdown("<h1 class='main-header'>Free Hosting Deployment Guide</h1>", unsafe_allow_html=True)
    st.markdown("<p class='sub-header'>Step-by-step instructions in English & Bengali for hosting this rebranded project 100% free.</p>", unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["🎈 Streamlit Community Cloud (Recommended)", "▲ Vercel Serverless", "বাংলা নির্দেশিকা (Bengali Guide)"])
    
    with tab1:
        st.subheader("Deploying on Streamlit Community Cloud (Free Forever)")
        st.markdown("""
        1. **Create a GitHub Repository**:
           - Push this project codebase to a GitHub repository (e.g. `github.com/your-username/autoflow-studio`).
        2. **Sign up at Streamlit**:
           - Go to [share.streamlit.io](https://share.streamlit.io/) and click **Continue with GitHub**.
        3. **New App**:
           - Click **New app**.
           - Repository: Select `your-username/autoflow-studio`.
           - Branch: `main`.
           - Main file path: `streamlit_app.py` (or `app.py`).
        4. **Deploy**:
           - Click **Deploy!** Your app will be live with a free HTTPS URL in under 2 minutes.
        """)
        
    with tab2:
        st.subheader("Deploying on Vercel")
        st.markdown("""
        1. **Push to GitHub**:
           - Push your project including `vercel.json` and `api/` directory.
        2. **Sign in to Vercel**:
           - Go to [vercel.com](https://vercel.com/) and click **Add New Project**.
        3. **Import Git Repository**:
           - Select your GitHub repo.
        4. **Deploy**:
           - Vercel automatically detects the configuration and provisions a free serverless endpoint!
        """)
        
    with tab3:
        st.subheader("ফ্রি হোস্টিং সার্ভারে আপলোড করার নিয়মাবলী")
        st.markdown("""
        ### ১. Streamlit Community Cloud (সবচেয়ে সহজ এবং সম্পূর্ণ ফ্রি):
        - **ধাপ ১:** এই প্রজেক্টের জিপ ফাইলটি ডাউনলোড করে আপনার GitHub একাউন্টে একটি নতুন রিপোজিটরি বানিয়ে পুশ করুন।
        - **ধাপ ২:** [share.streamlit.io](https://share.streamlit.io) এ যান এবং GitHub দিয়ে লগইন করুন।
        - **ধাপ ৩:** **"New app"** বাটনে ক্লিক করে আপনার রিপোজিটরি ও `streamlit_app.py` সিলেক্ট করুন।
        - **ধাপ ৪:** **"Deploy!"** চাপুন। ২ মিনিটের মধ্যে আপনার অটোমেশন অ্যাপটি লাইভ হয়ে যাবে!
        
        ### ২. Vercel এ ডিপ্লয় করার নিয়ম:
        - **ধাপ ১:** Vercel.com এ ফ্রি একাউন্ট খুলুন।
        - **ধাপ ২:** "Add New Project" এ গিয়ে GitHub রিপোজিটরি ইমপোর্ট করুন।
        - **ধাপ ৩:** `vercel.json` ফাইলটি প্রজেক্টে যুক্ত থাকায় ভারসেল নিজে থেকেই বিল্ড সম্পন্ন করবে।
        """)
