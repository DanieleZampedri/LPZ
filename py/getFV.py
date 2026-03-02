import random
import time
import requests
from fxconverter import CurrencyConverter
import pandas as pd
from openpyxl import load_workbook
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from dotenv import load_dotenv
import os   

EXCEL_FILE_PATH = '../excel/Valuation Model LPZ Investing.xlsx'
load_dotenv()
COOKIE = os.getenv("COOKIE")
converter = CurrencyConverter()

URL = "https://it.investing.com/pro/_/api/query?raw=true"
HEADERS = {
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0",
    "Referer": "https://it.investing.com/",
    "Cookie": COOKIE,
}
QUERY = """
query loadModel ($slug: String!, $ticker: String!) {
  model: create_model (slug: $slug, ticker: $ticker) {
    workbook
  }
}
"""

def create_session():
    session = requests.Session()
    retry_strategy = Retry(
        total=2,
        backoff_factor=1,
        status_forcelist=[500, 502, 503, 504],
        allowed_methods=["POST"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

def extract_data(ticker,session):
    payload = {
        "query": QUERY,
        "variables": {
            "slug": "dcf-growth-exit-5yr",
            "ticker": ticker
        }
    }
    
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            response = session.post(URL, json=payload, headers=HEADERS, timeout=30)
            if response.status_code == 429:
                wait_time = (2 ** attempt) * 30
                print(f"Rate limit raggiunto. Attesa di {wait_time} secondi prima del nuovo tentativo...")
                time.sleep(wait_time)
                continue
            if response.status_code != 200:
                print(f"Errore nella richiesta API: {response.status_code}")
                return None
            
            data = response.json()
            if data.get("data") is None or data["data"].get("model") is None:
                print(f"Errore nella risposta API: {data}")
                return None
            named_values = data["data"]["model"]["workbook"]["named_values"]
            currency = data["data"]["model"]["workbook"]["properties"]["trading_currency"]
            raw = {
                "fv": named_values.get("fv_mid"),
                "ebitda": named_values.get("ebitda"),
                "ebitda_gp": named_values.get("ebitda_gp"),
                "ebitda_gnp": named_values.get("ebitda_gnp"),
                "fcf": named_values.get("fcf"),
                "net_debt": named_values.get("net_debt"),
            }

            result = {}
            for key, value in raw.items():
                if value is not None and "value" in value:
                    result[key] = to_usd(float(value["value"]), currency)
                else:
                    print(f"{key} non trovato per {ticker}")
                    result[key] = None
            return result
        
        
        except KeyError as e:
            print(f"Key error per {ticker}: {e}")
            return None
        except requests.exceptions.RequestException as e:
            print(f"Errore nella richiesta API: {e}")
            if attempt < max_attempts - 1:
                time.sleep(2 ** attempt * 5)
            else:
                return None
    return None
        

def to_usd(amount:float, currency:str) -> float:
    if currency == "USD":
        return amount
    else:
        try:
            return converter.convert(amount, currency.upper(), "USD")
        except Exception as e:
            print(f"Errore nella conversione della valuta: {e}")
            return amount

def process_sheets():
    wb = load_workbook(EXCEL_FILE_PATH)
    session = create_session()
    batch_size = 40
    sheets_to_skip = ["TOTAL SCORES", "Set up","Company names", "Database"]
    for sheet_name in wb.sheetnames:
        if sheet_name in sheets_to_skip:
            print(f"Skipping sheet: {sheet_name}")
            continue
        print(f"Processing sheet: {sheet_name}")
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=sheet_name, header=3)
        tickers = df.iloc[:,2].dropna()
        if len(tickers) == 0:
            print(f"Nessun ticker trovato nel foglio {sheet_name}")
            continue
        ws = wb[sheet_name]
        for i in range(0, len(tickers), batch_size):
            batch_tickers = tickers[i:i+batch_size]
            for df_idx, ticker in batch_tickers.items():
                row_number = df_idx + 4 + 1
                result = extract_data(ticker, session)
                time.sleep(random.uniform(0.8, 1.2))
                if sheet_name == "Real Estate":
                    offset = 1
                else:
                    offset = 0
                if result is not None:
                    print("Dati estratti per {ticker}: {result}")
                    ws.cell(row=row_number, column=8+offset, value=result.get("fv"))
                    ws.cell(row=row_number, column=12+offset, value=result.get("ebitda"))
                    ws.cell(row=row_number, column=26+offset, value=result.get("ebitda_gp"))
                    ws.cell(row=row_number, column=29+offset, value=result.get("ebitda_gnp"))
                    ws.cell(row=row_number, column=32+offset, value=result.get("fcf"))
                    ws.cell(row=row_number, column=35+offset, value=result.get("net_debt"))

                else:
                    print(f"Non sono riuscito a estrarre i dati per {ticker}")
            if i + batch_size < len(tickers):
                wait_time = random.uniform(10, 30)
                print(f"Attesa di {wait_time:.2f} secondi prima del prossimo batch...")
                time.sleep(wait_time)
        wb.save(EXCEL_FILE_PATH)
        print(f"Tutti i dati sono stati salvati nel file {sheet_name}.")

def main():
    try:
        process_sheets()
    except FileNotFoundError:
        print(f"Errore: Il file {EXCEL_FILE_PATH} non è stato trovato.")
    except Exception as e:
        print(f"Si è verificato un errore: {e}")
        raise


if __name__ == "__main__":
    main()