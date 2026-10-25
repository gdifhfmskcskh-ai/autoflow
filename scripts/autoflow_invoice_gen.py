"""
AutoFlow Studio - Zero-Binary Invoice Generator
Works on any Linux, Vercel, or Streamlit container without wkhtmltopdf binary dependencies.
"""

from jinja2 import Template
from datetime import datetime
import os

INVOICE_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Invoice #{{ invoice_number }}</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 40px; color: #1e293b; }
        .header { display: flex; justify-content: space-between; border-bottom: 2px solid #2563eb; padding-bottom: 20px; }
        .brand { font-size: 24px; font-weight: bold; color: #1e3a8a; }
        .meta { text-align: right; }
        .client-box { margin: 30px 0; padding: 15px; background: #f8fafc; border-radius: 6px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background: #1e3a8a; color: white; padding: 10px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #e2e8f0; }
        .total-box { margin-top: 25px; margin-left: auto; width: 280px; }
        .row { display: flex; justify-content: space-between; padding: 6px 0; }
        .grand-total { border-top: 2px solid #2563eb; font-size: 18px; font-weight: bold; color: #1e3a8a; }
    </style>
</head>
<body>
    <div class="header">
        <div>
            <div class="brand">{{ company_name }}</div>
            <div style="color: #64748b; font-size: 13px;">Cloud Automation Solutions</div>
        </div>
        <div class="meta">
            <h2 style="margin: 0; color: #2563eb;">INVOICE</h2>
            <div><strong>#{{ invoice_number }}</strong></div>
            <div style="color: #64748b; font-size: 13px;">Date: {{ issue_date }}</div>
        </div>
    </div>
    
    <div class="client-box">
        <div style="color: #64748b; font-size: 12px;">BILLED TO:</div>
        <div style="font-size: 16px; font-weight: bold;">{{ client_name }}</div>
        <div>{{ client_email }}</div>
    </div>
    
    <table>
        <thead>
            <tr>
                <th>Description</th>
                <th style="text-align: right;">Amount (USD)</th>
            </tr>
        </thead>
        <tbody>
            {% for item in items %}
            <tr>
                <td>{{ item.name }}</td>
                <td style="text-align: right;">${{ "%.2f"|format(item.price) }}</td>
            </tr>
            {% endfor %}
        </tbody>
    </table>
    
    <div class="total-box">
        <div class="row"><span>Subtotal:</span><span>${{ "%.2f"|format(subtotal) }}</span></div>
        <div class="row"><span>Tax ({{ tax_rate }}%):</span><span>${{ "%.2f"|format(tax_amount) }}</span></div>
        <div class="row grand-total"><span>Total Due:</span><span>${{ "%.2f"|format(total) }}</span></div>
    </div>
</body>
</html>
"""

def generate_invoice(client_name: str, client_email: str, items: list, output_path: str = "invoice.html"):
    subtotal = sum(i["price"] for i in items)
    tax_rate = 5.0
    tax_amount = subtotal * (tax_rate / 100.0)
    total = subtotal + tax_amount
    
    template = Template(INVOICE_TEMPLATE)
    rendered = template.render(
        company_name="AutoFlow Global Studio",
        invoice_number=f"INV-{datetime.now().strftime('%Y%m%d%H%M')}",
        issue_date=datetime.now().strftime("%B %d, %Y"),
        client_name=client_name,
        client_email=client_email,
        items=items,
        subtotal=subtotal,
        tax_rate=tax_rate,
        tax_amount=tax_amount,
        total=total
    )
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(rendered)
    print(f"Invoice generated: {output_path}")
    return output_path

if __name__ == "__main__":
    sample_items = [
        {"name": "Streamlit Cloud Automation Pipeline", "price": 499.0},
        {"name": "Excel Pivot Reporting Suite", "price": 299.0},
        {"name": "Vercel Serverless Integration", "price": 199.0}
    ]
    generate_invoice("Jane Doe", "jane@example.com", sample_items)
