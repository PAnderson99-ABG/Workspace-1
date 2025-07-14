import requests
from bs4 import BeautifulSoup
import urllib.parse

def search_menards_sku(sku):
    search_url = f"https://www.menards.com/main/search.html?search={urllib.parse.quote(sku)}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }

    response = requests.get(search_url, headers=headers)
    if response.status_code != 200:
        print(f"Failed to fetch results for SKU {sku} (status {response.status_code})")
        return None

    soup = BeautifulSoup(response.text, 'html.parser')
    product_links = soup.select('a.productGrid__itemLink')

    if not product_links:
        print(f"No product found for SKU {sku}")
        return None

    # Get the first product link
    partial_url = product_links[0].get('href')
    full_url = urllib.parse.urljoin("https://www.menards.com", partial_url)
    return full_url

# === Example usage ===
if __name__ == "__main__":
    sku = input("Enter SKU number: ").strip()
    url = search_menards_sku(sku)
    if url:
        print(f"Product URL for SKU {sku}:\n{url}")
