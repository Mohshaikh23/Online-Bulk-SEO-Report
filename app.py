import requests
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Google Sheets API credentials and settings
SHEET_SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
sheet_id = "1e0Zsq7Iyoe3KZbJ4aowBVMLkOJ_veMnyLsssagUtWqE"
page_url = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
api_key = "Your API-key"  # Replace with your actual API key

def authenticate():
    """Authenticate and authorize using Google Sheets API."""
    flow = InstalledAppFlow.from_client_secrets_file(
        'client_secret.json', scopes=SHEET_SCOPES
    )
    credentials = flow.run_local_server(port=0)
    service = build('sheets', 'v4', credentials=credentials)
    return service.spreadsheets()

def fetch_website_data(sheet, range_name):
    """Fetch website URLs from Google Sheets."""
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get('values', [])
    if not values:
        raise ValueError(f"No data found in {range_name}.")
    return [url[0] for url in values]

def fetch_metrics(url):
    """Fetch metrics for a given URL from PageSpeed Insights API."""
    params = {
        "url": f'https://{url}',
        "strategy": "desktop",
        "key": api_key,
        "category": ["accessibility", "best-practices", "performance", "pwa", "seo"]
    }

    try:
        res = requests.get(page_url, params=params)
        res.raise_for_status()  # Check if the request was successful

        cat = res.json().get("lighthouseResult", {}).get("categories", {})
        audits = res.json().get("lighthouseResult", {}).get("audits", {})

        performance = f"{cat.get('performance', {}).get('score', 0) * 100} %"
        seo = f"{cat.get('seo', {}).get('score', 0) * 100} %"
        best_practices = f"{cat.get('best-practices', {}).get('score', 0) * 100} %"
        accessibility = f"{cat.get('accessibility', {}).get('score', 0) * 100} %"

        # Extract page speed metrics in seconds
        fcp = audits.get("first-contentful-paint", {}).get("numericValue", 0) / 1000
        fmp = audits.get("first-meaningful-paint", {}).get("numericValue", 0) / 1000
        lcp = audits.get("largest-contentful-paint", {}).get("numericValue", 0) / 1000
        tbt = audits.get("total-blocking-time", {}).get("numericValue", 0) / 1000

        return {
            "Website": url,
            "Performance": performance,
            "SEO": seo,
            "Best Practices": best_practices,
            "Accessibility": accessibility,
            "First Contentful Paint (s)": fcp,
            "First Meaningful Paint (s)": fmp,
            "Largest Contentful Paint (s)": lcp,
            "Total Blocking Time (s)": tbt
        }
    except Exception as e:
        print(f"Error fetching data for {url}: {e}")
        return None

def main():
    try:
        # Authenticate and get spreadsheet service
        sheet_service = authenticate()

        # Fetch website URLs from Google Sheets
        websites = fetch_website_data(sheet_service,
                                      "Australia_Personalized!F1:F150")

        # Fetch metrics for each website
        metrics_data = []
        for url in websites:
            metrics = fetch_metrics(url)
            if metrics:
                metrics_data.append(metrics)

        # Create DataFrame from metrics data
        df = pd.DataFrame(metrics_data)

        # Save DataFrame to Excel
        df.to_excel("website_metrics.xlsx", index=False)
        print("Metrics saved to website_metrics.xlsx.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
