"""
AutoFlow Studio - Cloud-Native News Scraper
Runs on any serverless or Streamlit container with zero browser binary requirements.
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os

def scrape_headlines(target_url: str = "https://news.ycombinator.com/", output_file: str = None):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AutoFlowStudio/1.0"
    }
    
    response = requests.get(target_url, headers=headers, timeout=10)
    response.raise_for_status()
    
    soup = BeautifulSoup(response.text, "html.parser")
    headlines = []
    
    # Generic article/heading extractor
    for tag in soup.find_all(['h1', 'h2', 'h3', 'a']):
        text = tag.get_text().strip()
        link = tag.get('href', '')
        if len(text) > 25 and not text.startswith("http"):
            headlines.append({
                "headline": text,
                "url": link,
                "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            
    # Deduplicate and limit to top 25
    df = pd.DataFrame(headlines).drop_duplicates(subset=['headline']).head(25)
    
    if output_file is None:
        today_str = datetime.now().strftime("%Y%m%d")
        output_file = f"news_headlines_{today_str}.csv"
        
    df.to_csv(output_file, index=False)
    print(f"Scraped {len(df)} headlines into '{output_file}'")
    return df

if __name__ == "__main__":
    scrape_headlines()
