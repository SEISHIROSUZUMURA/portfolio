import requests
from xml.etree import ElementTree
from datetime import datetime, timedelta
import time
from openpyxl import Workbook
import random

class RateLimiter:
    def __init__(self, calls, period):
        self.calls = calls
        self.period = period
        self.last_reset = datetime.now()
        self.num_calls = 0

    def __call__(self, func):
        def wrapper(*args, **kwargs):
            now = datetime.now()
            if now - self.last_reset > timedelta(seconds=self.period):
                self.num_calls = 0
                self.last_reset = now
            
            if self.num_calls >= self.calls:
                time_to_wait = self.period - (now - self.last_reset).total_seconds()
                if time_to_wait > 0:
                    time.sleep(time_to_wait)
                self.num_calls = 0
                self.last_reset = datetime.now()
            
            self.num_calls += 1
            return func(*args, **kwargs)
        return wrapper

@RateLimiter(calls=1, period=3)
def call_api(url, params):
    """Make an API call with rate limiting."""
    response = requests.get(url, params=params)
    if response.status_code != 200:
        raise Exception(f"API request failed with status code: {response.status_code}")
    return response

def parse_pubdate(pub_date):
    """Parse publication date from PubMed XML element."""
    date_parts = [
        elem.text for elem in [pub_date.find(tag) for tag in ['Year', 'Month', 'Day']]
        if elem is not None
    ]
    return "-".join(date_parts) if date_parts else "Date not available"

def extract_text_or_default(element, xpath, default="Not available"):
    """Extract text from XML element or return default value."""
    found = element.find(xpath)
    return found.text if found is not None else default

def parse_authors_and_affiliations(article):
    """Extract authors and their affiliations from article XML."""
    authors = []
    affiliations = set()
    for author in article.findall(".//Author"):
        lastname = extract_text_or_default(author, "LastName")
        forename = extract_text_or_default(author, "ForeName")
        if lastname != "Not available":
            authors.append(f"{lastname}, {forename}")
        
        affiliation = extract_text_or_default(author, "AffiliationInfo/Affiliation")
        if affiliation != "Not available":
            affiliations.add(affiliation)

    return (
        "; ".join(authors) if authors else "Authors not available",
        "; ".join(affiliations) if affiliations else "Affiliations not available"
    )

def fetch_pubmed_data(query, max_results=1000):
    """Fetch data from PubMed API with rate limiting."""
    base_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
    search_url = f"{base_url}esearch.fcgi"
    fetch_url = f"{base_url}efetch.fcgi"

    # Search for article IDs
    search_params = {
        "db": "pubmed",
        "term": query,
        "retmax": max_results,
        "retmode": "json",
        "sort": "relevance"
    }
    response = call_api(search_url, search_params)
    search_results = response.json()

    if "esearchresult" not in search_results:
        print("No results found")
        return []

    id_list = search_results["esearchresult"]["idlist"]
    return fetch_articles(fetch_url, id_list)

def fetch_articles(fetch_url, id_list, batch_size=200):
    """Fetch article details in batches with rate limiting."""
    articles = []
    for i in range(0, len(id_list), batch_size):
        batch_ids = id_list[i:i+batch_size]
        
        fetch_params = {
            "db": "pubmed",
            "id": ",".join(batch_ids),
            "retmode": "xml"
        }
        response = call_api(fetch_url, fetch_params)
        root = ElementTree.fromstring(response.content)

        for article in root.findall(".//PubmedArticle"):
            articles.append(parse_article(article))

        print(f"Fetched {len(articles)} articles so far...")
        
        # ランダムな待機時間を追加（2〜3秒）
        time.sleep(random.uniform(2, 3))

    return articles

def parse_article(article):
    """Parse individual article XML and extract relevant information."""
    pmid = extract_text_or_default(article, ".//PMID")
    title = extract_text_or_default(article, ".//ArticleTitle")
    abstract = extract_text_or_default(article, ".//Abstract/AbstractText")
    
    pub_date = article.find(".//PubDate")
    date = parse_pubdate(pub_date) if pub_date is not None else "Date not available"

    journal = extract_text_or_default(article, ".//Journal/Title")
    authors, affiliations = parse_authors_and_affiliations(article)

    keywords = article.findall(".//Keyword")
    keywords = "; ".join([k.text for k in keywords]) if keywords else "Keywords not available"

    doi = extract_text_or_default(article, ".//ELocationID[@EIdType='doi']")

    return {
        "pmid": pmid,
        "title": title,
        "abstract": abstract,
        "date": date,
        "journal": journal,
        "authors": authors,
        "affiliations": affiliations,
        "keywords": keywords,
        "doi": doi
    }

def save_results_to_excel(results, filename):
    """Save parsed articles data to an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthetic Biology Articles"

    headers = ["pmid", "title", "date", "journal", "authors", "affiliations", "keywords", "doi", "abstract"]
    ws.append(headers)

    for article in results:
        row = [article.get(key, "Not available") for key in headers]
        ws.append(row)

    for column in ws.columns:
        max_length = max(len(str(cell.value)) for cell in column)
        adjusted_width = min((max_length + 2) * 1.2, 100)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

    wb.save(filename)
    print(f"Results saved to {filename}")

def main():
    query = "synthetic biology"
    results = fetch_pubmed_data(query, max_results=1000)
    
    print(f"\nTotal articles found for '{query}': {len(results)}")

    save_results_to_excel(results, "synthetic_biology_articles_comprehensive.xlsx")

if __name__ == "__main__":
    main()
