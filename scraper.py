import requests
import re
import json
import csv
import os
from datetime import datetime
from bs4 import BeautifulSoup
import openpyxl

# ─── Fetch Page ───────────────────────────────────────────

def fetch_page(url):
    print("\n[*] Fetching: " + url)
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
        }
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            print("[+] Page fetched successfully.")
            return response.text
        else:
            print("[-] Failed. Status code: " + str(response.status_code))
            return None
    except requests.exceptions.ConnectionError:
        print("[-] Could not connect to: " + url)
        return None
    except requests.exceptions.Timeout:
        print("[-] Connection timed out.")
        return None

# ─── Extractors ───────────────────────────────────────────

def extract_headlines(soup):
    print("\n[*] Extracting headlines...")
    headlines = []
    for tag in ['h1', 'h2', 'h3', 'h4']:
        for item in soup.find_all(tag):
            text = item.get_text().strip()
            if text and len(text) > 3:
                headlines.append({
                    "type": tag.upper(),
                    "text": text
                })
    print("[+] Found " + str(len(headlines)) + " headlines.")
    return headlines

def extract_links(soup, base_url):
    print("\n[*] Extracting links...")
    links = []
    for item in soup.find_all('a', href=True):
        href = item['href'].strip()
        text = item.get_text().strip()
        if href.startswith('http'):
            full_url = href
        elif href.startswith('/'):
            full_url = base_url.rstrip('/') + href
        elif href.startswith('#') or href.startswith('mailto'):
            continue
        else:
            full_url = base_url.rstrip('/') + '/' + href
        if full_url not in [l['url'] for l in links]:
            links.append({
                "text": text if text else "No text",
                "url": full_url
            })
    print("[+] Found " + str(len(links)) + " links.")
    return links

def extract_images(soup, base_url):
    print("\n[*] Extracting images...")
    images = []
    for item in soup.find_all('img'):
        src = item.get('src', '').strip()
        alt = item.get('alt', '').strip()
        if not src:
            continue
        if src.startswith('http'):
            full_url = src
        elif src.startswith('/'):
            full_url = base_url.rstrip('/') + src
        else:
            full_url = base_url.rstrip('/') + '/' + src
        images.append({
            "alt": alt if alt else "No description",
            "url": full_url
        })
    print("[+] Found " + str(len(images)) + " images.")
    return images

def extract_emails(soup, html):
    print("\n[*] Extracting emails...")
    email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
    emails = list(set(email_pattern.findall(html)))
    print("[+] Found " + str(len(emails)) + " emails.")
    return emails

def extract_phone_numbers(html):
    print("\n[*] Extracting phone numbers...")
    phone_pattern = re.compile(
        r'(\+?[\d\s\-\(\)]{7,15})'
    )
    raw = phone_pattern.findall(html)
    phones = []
    for p in raw:
        cleaned = p.strip()
        if len(cleaned) >= 7:
            if cleaned not in phones:
                phones.append(cleaned)
    phones = phones[:50]
    print("[+] Found " + str(len(phones)) + " phone numbers.")
    return phones

def extract_tables(soup):
    print("\n[*] Extracting tables...")
    tables = []
    for table in soup.find_all('table'):
        rows = []
        for row in table.find_all('tr'):
            cells = []
            for cell in row.find_all(['td', 'th']):
                cells.append(cell.get_text().strip())
            if cells:
                rows.append(cells)
        if rows:
            tables.append(rows)
    print("[+] Found " + str(len(tables)) + " tables.")
    return tables

def extract_paragraphs(soup):
    print("\n[*] Extracting text paragraphs...")
    paragraphs = []
    for p in soup.find_all('p'):
        text = p.get_text().strip()
        if text and len(text) > 5:
            paragraphs.append(text)
    print("[+] Found " + str(len(paragraphs)) + " paragraphs.")
    return paragraphs

# ─── Save Functions ───────────────────────────────────────

def save_json(data, filename):
    with open(filename + ".json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print("[+] JSON saved: " + filename + ".json")

def save_csv(data, filename):
    if not data:
        return
    with open(filename + ".csv", "w", newline="", encoding="utf-8") as f:
        if isinstance(data[0], dict):
            writer = csv.DictWriter(f, fieldnames=data[0].keys())
            writer.writeheader()
            writer.writerows(data)
        else:
            writer = csv.writer(f)
            for row in data:
                writer.writerow([row])
    print("[+] CSV saved: " + filename + ".csv")

def save_excel(results, filename):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sections = {
        "Headlines": results.get("headlines", []),
        "Links": results.get("links", []),
        "Images": results.get("images", []),
        "Emails": results.get("emails", []),
        "Phones": results.get("phones", []),
        "Paragraphs": results.get("paragraphs", []),
    }

    for sheet_name, data in sections.items():
        if not data:
            continue
        ws = wb.create_sheet(title=sheet_name)
        if isinstance(data[0], dict):
            headers = list(data[0].keys())
            ws.append(headers)
            for row in data:
                ws.append(list(row.values()))
        else:
            ws.append([sheet_name])
            for item in data:
                ws.append([item])

    if results.get("tables"):
        for i, table in enumerate(results["tables"]):
            ws = wb.create_sheet(title="Table_" + str(i + 1))
            for row in table:
                ws.append(row)

    excel_filename = filename + ".xlsx"
    wb.save(excel_filename)
    print("[+] Excel saved: " + excel_filename)

# ─── Main ─────────────────────────────────────────────────

def main():
    print("=" * 50)
    print("   Universal Web Scraper")
    print("   by ashardrach")
    print("=" * 50)

    url = input("\nEnter URL to scrape: ").strip()
    if not url.startswith("http"):
        url = "http://" + url

    base_url = "/".join(url.split("/")[:3])

    html = fetch_page(url)
    if not html:
        return

    soup = BeautifulSoup(html, 'html.parser')

    print("\nWhat do you want to extract?")
    print("1. Headlines")
    print("2. Links")
    print("3. Images")
    print("4. Emails")
    print("5. Phone numbers")
    print("6. Tables")
    print("7. Paragraphs")
    print("8. Everything")
    choice = input("\nEnter choice (1-8): ").strip()

    results = {}

    if choice == "1" or choice == "8":
        results["headlines"] = extract_headlines(soup)
    if choice == "2" or choice == "8":
        results["links"] = extract_links(soup, base_url)
    if choice == "3" or choice == "8":
        results["images"] = extract_images(soup, base_url)
    if choice == "4" or choice == "8":
        results["emails"] = extract_emails(soup, html)
    if choice == "5" or choice == "8":
        results["phones"] = extract_phone_numbers(html)
    if choice == "6" or choice == "8":
        results["tables"] = extract_tables(soup)
    if choice == "7" or choice == "8":
        results["paragraphs"] = extract_paragraphs(soup)

    if not results:
        print("[-] Nothing extracted.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    domain = base_url.replace("https://", "").replace("http://", "").replace("/", "_")
    filename = "scrape_" + domain + "_" + timestamp

    print("\n[*] Saving results...")
    save_json(results, filename)
    save_csv(results.get("headlines", results.get("links",
             results.get("emails", []))), filename)
    save_excel(results, filename)

    print("\n" + "=" * 50)
    print("SCRAPE COMPLETE")
    print("=" * 50)
    for key, value in results.items():
        if isinstance(value, list):
            print("[+] " + key.capitalize() + ": " + str(len(value)) + " found")
    print("=" * 50)

main()