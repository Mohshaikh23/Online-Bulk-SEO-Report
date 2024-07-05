import requests
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

page_url = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
api_key = "AIzaSyDasl6GVbKprqBSWjS43e269-mMg86K5uU"  # Replace with your actual API key
SHEET_SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
sheet_id = "1e0Zsq7Iyoe3KZbJ4aowBVMLkOJ_veMnyLsssagUtWqE"


def get_websites():
    # Authenticate and authorize
    flow = InstalledAppFlow.from_client_secrets_file(
        'client_secret.json', scopes=SHEET_SCOPES
    )
    credentials = flow.run_local_server(port=0)
    service = build('sheets', 'v4', credentials=credentials)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id,
                                range="Australia_Personalized!F1:F150").execute()
    values = result.get('values', [])

    # Convert the values to a Pandas dataframe, considering the first row as headers
    if values:
        df1 = pd.DataFrame(values[1:], columns=values[0])
    else:
        df1 = pd.DataFrame()

    df1.to_csv('websites_data.csv')

def metrics():
    data = []
    df1 = pd.read_csv('websites_data.csv')
    for url in df1['Website']:
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

            # Extract page speed metrics in milliseconds and convert to seconds
            fcp = audits.get("first-contentful-paint", {}).get("numericValue", 0) / 1000
            fmp = audits.get("first-meaningful-paint", {}).get("numericValue", 0) / 1000
            lcp = audits.get("largest-contentful-paint", {}).get("numericValue", 0) / 1000
            tbt = audits.get("total-blocking-time", {}).get("numericValue", 0) / 1000

            # Append the fetched data to the list
            data.append({
                "Website": url,
                "Performance": performance,
                "SEO": seo,
                "Best Practices": best_practices,
                "Accessibility": accessibility,
                "First Contentful Paint (s)": fcp,
                "First Meaningful Paint (s)": fmp,
                "Largest Contentful Paint (s)": lcp,
                "Total Blocking Time (s)": tbt
            })

            # Create a DataFrame from the list and print it
            df = pd.DataFrame(data)
        
        except Exception as e:
            print(f"Error fetching data for {url}: {e}")

    # Save the final DataFrame to an Excel file
    df.to_excel("website_metrics.xlsx", index=False)

def execution():
    try:
        get_websites()
        metrics()
    except Exception as e:
        print(e)