"""
AutoFlow Studio - Cloud Web Table Extractor
Scrapes and converts HTML tables into clean CSV and Excel files.
"""

import pandas as pd
import requests
import io

def extract_tables_from_url(url: str, output_prefix: str = "extracted_table"):
    """
    Fetches web page and extracts all HTML <table> tags into Pandas DataFrames.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    response = requests.get(url, headers=headers, timeout=15)
    response.raise_for_status()
    
    tables = pd.read_html(io.StringIO(response.text))
    print(f"Discovered {len(tables)} tables on page.")
    
    exported_files = []
    for idx, table in enumerate(tables):
        filename = f"{output_prefix}_{idx + 1}.csv"
        table.to_csv(filename, index=False)
        exported_files.append(filename)
        print(f"Table {idx + 1} saved ({len(table)} rows) -> {filename}")
        
    return exported_files

def extract_tables_from_html_file(file_path: str):
    with open(file_path, "r", encoding="utf-8") as f:
        html = f.read()
    tables = pd.read_html(io.StringIO(html))
    return tables

if __name__ == "__main__":
    # Example: Wikipedia highest grossing films
    test_url = "https://en.wikipedia.org/wiki/List_of_highest-grossing_films"
    try:
        extract_tables_from_url(test_url)
    except Exception as err:
        print("Scraping note:", err)
