#Goes through all HTML pages of the provided website, and extracts their text in TXT and HTML format. 
#!/usr/bin/env python
# coding: utf-8

# In[1]:


import aiohttp
import asyncio
import aiofiles
from bs4 import BeautifulSoup
import pandas as pd
import time
from urllib.parse import urljoin, urlparse
from tqdm import tqdm
import re
import nest_asyncio


# In[2]:


# Fix async issues
nest_asyncio.apply()


# In[3]:


class AsyncWebScraper:
    def __init__(self, start_url):
        self.start_url = start_url
        self.visited = set()
        self.to_visit = [start_url]
        self.data = []
        self.headers = {"User-Agent": "Mozilla/5.0"}

    async def fetch(self, session, url):
        """Fetch webpage async"""
        try:
            async with session.get(url, headers=self.headers, timeout=10) as response:
                if response.status == 200:
                    return await response.text()
        except Exception as e:
            print(f"Failed to scrape {url}: {e}")
        return None

    def clean_text(self, text):
        """Improve readability"""
        text = text.strip()
        text = re.sub(r'\s*\|\s*', '\n', text)  # Split on " | " into new lines
        text = re.sub(r'(?<=[a-zα-ω])\s+(?=[A-ZΑ-Ω])', '\n', text)  # Add newline before capitalized words
        text = re.sub(r'(?<=\.)\s+', '\n', text)  # Add new line after a period
        text = re.sub(r'\s{2,}', '\n', text)  # Remove excess spaces
        return text.strip()

    async def scrape_page(self, session, url):
        """Scrape text from a page."""
        if url in self.visited:
            return
        self.visited.add(url)

        print(f"Scraping: {url}")
        html = await self.fetch(session, url)
        if not html:
            return

        soup = BeautifulSoup(html, "html.parser")
        for element in soup(["script", "style", "noscript"]):
            element.extract()

        # Extract text
        texts = [self.clean_text(text) for text in soup.stripped_strings if len(text) > 3]
        self.data.append({"url": url, "text": "\n".join(texts)})

        # Extract internal links and add them to queue
        for link in soup.find_all("a", href=True):
            absolute_url = urljoin(url, link["href"])
            if self.is_internal_url(absolute_url) and absolute_url not in self.visited:
                self.to_visit.append(absolute_url)

    def is_internal_url(self, url):
        """Check if a URL belongs to the same domain and is not external."""
        parsed_url = urlparse(url)
        return parsed_url.netloc == urlparse(self.start_url).netloc and parsed_url.scheme in ["http", "https"]

    async def crawl(self):
        """Run async crawler"""
        start_time = time.time()  # Track start time

        async with aiohttp.ClientSession() as session:
            while self.to_visit:
                tasks = []
                for _ in range(min(len(self.to_visit), 10)):  # Process up to 10 pages in parallel
                    url = self.to_visit.pop(0)
                    tasks.append(self.scrape_page(session, url))

                await asyncio.gather(*tasks)  # Run tasks in parallel

        end_time = time.time()  # Track end time
        total_time = round(end_time - start_time, 2)
        print(f"\n✅ Scraping completed in {total_time} seconds.\n")

        return pd.DataFrame(self.data)


# In[ ]:


start_url = "https://www.XXX.com/" #YOUR WEBSITE HERE
scraper = AsyncWebScraper(start_url)

df = await scraper.crawl()  # Run scraper
display(df)


# In[ ]:


def save_to_html(df, filename="website_text.html"):
    """Save extracted website text as a structured HTML file."""
    with open(filename, "w", encoding="utf-8-sig") as file:
        file.write("<html><head><title>Extracted Website Text</title></head><body>\n")
        file.write("<h1>Extracted Website Content</h1>\n")
        
        for index, row in df.iterrows():
            formatted_text = row['text'].replace("\n", "<br>")
            file.write(f"<h2><a href='{row['url']}'>{row['url']}</a></h2>\n")
            file.write("<hr>\n")
            file.write(f"<p>{formatted_text}</p>\n")
            file.write("<br><hr><br>\n")
        
        file.write("</body></html>")
    print(f"✅ Extracted text saved as {filename} (HTML Format)")

save_to_html(df, "website_text.html")


# In[ ]:


def save_to_text(df, filename="website_text.txt"):
    """Save extracted website text in a structured TXT file."""
    with open(filename, "w", encoding="utf-8-sig") as file:
        for index, row in df.iterrows():
            file.write(f"URL: {row['url']}\n")
            file.write("="*80 + "\n")
            formatted_text = row["text"]  # Text is already cleaned
            file.write(formatted_text + "\n\n")
            file.write("-"*80 + "\n\n")
    print(f"✅ Extracted text saved as {filename} (Readable TXT Format)")

save_to_text(df, "website_text.txt")

