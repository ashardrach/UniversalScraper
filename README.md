# Universal Web Scraper 🕷️

A Python tool that extracts any data type from any website
and saves results in JSON, CSV, and Excel format.

## What It Extracts
- Headlines (H1, H2, H3, H4)
- All links with anchor text
- Images with descriptions
- Email addresses
- Phone numbers
- Tables and structured data
- Text paragraphs

## Output Formats
- JSON — for developers and APIs
- CSV — for data analysis
- Excel (.xlsx) — for clients and business use
  (each data type on its own sheet)

## How to Run
1. Install required libraries:
   pip install requests beautifulsoup4 openpyxl
2. Run the scraper:
   python scraper.py
3. Enter target URL
4. Choose what to extract
5. Find your results in the project folder

## Legal Notice
Only scrape websites you have permission to scrape.
Always check robots.txt before scraping.
Do not use for collecting personal data without consent.

## Use Cases
- Competitor price monitoring
- Lead generation from public directories
- News and content aggregation
- Market research and data collection

## Built By
ashardrach — cybersecurity student and Python developer.
GitHub: github.com/ashardrach