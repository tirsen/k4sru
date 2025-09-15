import argparse
import requests
import openpyxl
from datetime import datetime
from dataclasses import dataclass
from typing import Dict, List

SALES_PER_PAGE = 7
EXCLUDE_ORDER_TYPES = ["RS STC"]

# Riksbanken API constants
RIKSBANK_API_BASE = "https://api.riksbank.se/swea/v1"
USD_SEK_SERIES_ID = "SEKUSDPMI"

# Export as yearly/daily USD/SEK as xlsx from
# https://www.riksbank.se/sv/statistik/rantor-och-valutakurser/sok-rantor-och-valutakurser

# It is allowed to use either the daily exchange rate or the yearly exchange rate,
# as long as it is consistent for all trades.

parser = argparse.ArgumentParser(prog="k4", description='Generate K4 SRU files.')
parser.add_argument('--org-nummer', required=True,
                    help='org eller personnummer')
parser.add_argument('--fullt-namn', required=True,
                    help='för och efternamn')
parser.add_argument('--adress', required=True,
                    help='adress')
parser.add_argument('--postnummer', required=True,
                    help='postnummer')
parser.add_argument('--postort', required=True,
                    help='postort')
parser.add_argument('--epost', required=True,
                    help='epost')
parser.add_argument('--trades', required=True,
                    help='affärer i xlsx')
parser.add_argument('--rate-type', choices=['year', 'day'], required=True,
                    help='Specify if using yearly or daily exchange rate')


@dataclass
class Trade:
  antal: int
  beteckning: str
  forsaljsningspris: float
  omkostnadsbelopp: float
  gainloss: float
  ordertype: str
  date: datetime
  ordinary_income: float = 0.0

  def vinst(self):
    return self.gainloss if self.gainloss > 0 else 0

  def forlust(self):
    return -self.gainloss if self.gainloss < 0 else 0


def parse_trades(file):
  trades = openpyxl.load_workbook(file)
  for row in trades.active.rows:
    if row[0].value == 'Sell':
      yield Trade(antal=row[3].value,
                  beteckning=row[1].value,
                  omkostnadsbelopp=row[10].value,
                  forsaljsningspris=row[13].value,
                  gainloss=row[18].value,
                  ordertype=row[28].value if len(row) > 28 else None,
                  date=datetime.strptime(row[12].value, "%m/%d/%Y"),
                  ordinary_income=row[8].value if len(row) > 8 and row[8].value else 0.0)


def fetch_exchange_rates(year: str) -> Dict[datetime, float]:
    """Fetch all exchange rates for a given year"""
    # Validate year
    current_year = datetime.now().year
    year_int = int(year)
    if year_int > current_year:
        raise ValueError(f"Cannot fetch exchange rates for future year {year}. Current year is {current_year}.")
    if year_int < 2000:  # Reasonable minimum year
        raise ValueError(f"Year {year} is too far in the past. Please use a year from 2000 onwards.")
    
    # Use January 2nd and December 30th to avoid holidays
    start_date = f"{year}-01-02"
    end_date = f"{year}-12-30"
    
    url = f"{RIKSBANK_API_BASE}/Observations/{USD_SEK_SERIES_ID}/{start_date}/{end_date}"
    response = requests.get(url)
    
    # Handle 204 No Content response
    if response.status_code == 204:
        raise ValueError(f"No exchange rate data available for year {year}. This could be because:\n"
                        f"1. The year is too recent (data might not be published yet)\n"
                        f"2. The year is too old (data might not be available)\n"
                        f"3. There might be an issue with the API\n"
                        f"Please try using a different year or check the Riksbanken website for available data.")
    
    # Print response details for debugging
    print(f"API Response Status: {response.status_code}")
    print(f"API Response Headers: {response.headers}")
    print(f"API Response Text: {response.text[:500]}...")  # Print first 500 chars
    
    response.raise_for_status()
    
    try:
        data = response.json()
    except requests.exceptions.JSONDecodeError as e:
        raise ValueError(f"Failed to parse API response. Status: {response.status_code}, Response: {response.text[:500]}...") from e
    
    
    rates = {datetime.strptime(obs['date'], "%Y-%m-%d"): float(obs['value']) for obs in data}
    
    if not rates:
        raise ValueError(f"No exchange rates found for year {year}. This could be because:\n"
                        f"1. The year is too recent (data might not be published yet)\n"
                        f"2. The year is too old (data might not be available)\n"
                        f"3. There might be an issue with the API\n"
                        f"Please try using a different year or check the Riksbanken website for available data.")
    
    return rates

def get_yearly_rate(rates: Dict[datetime, float]) -> float:
    """Calculate the average yearly rate from daily rates"""
    if not rates:
        raise ValueError("No exchange rates found for the year")
    return sum(rates.values()) / len(rates)

def get_daily_rate(rates: Dict[datetime, float], date: datetime) -> float:
    """Get the exchange rate for a specific date, or the closest available date within 7 days"""
    if not rates:
        raise ValueError("No exchange rates found for the year")
    
    # Try to find exact date
    if date in rates:
        return rates[date]
    
    # Find closest date within 7 days
    closest_date = min(rates.keys(), key=lambda d: abs(d - date))
    if abs((closest_date - date).days) > 7:
        raise ValueError(f"No exchange rate found within seven days of {date}")
    return rates[closest_date]

def get_year_from_trades(trades: List[Trade]) -> str:
    """Determine the year from trades and validate all trades are from the same year"""
    if not trades:
        raise ValueError("No trades found in the input file")
    
    # Get the year from the first trade
    year = trades[0].date.year
    
    # Validate all trades are from the same year
    for trade in trades:
        if trade.date.year != year:
            raise ValueError(f"All trades must be from the same year. Found trades from {year} and {trade.date.year}")
    
    return str(year)

def main():
  # https://www.skatteverket.se/foretag/etjansterochblanketter/blanketterbroschyrer/broschyrer/info/269.4.39f16f103821c58f680007305.html
  args = parser.parse_args()

  use_yearly_rate = args.rate_type == 'year'

  # Parse all trades first
  trades = [trade for trade in list(parse_trades(args.trades)) if trade.ordertype not in EXCLUDE_ORDER_TYPES]
  
  # Determine the year from trades
  year = get_year_from_trades(trades)
  print(f"Processing trades from year: {year}")

  # Pre-fetch all exchange rates for the year
  rates = fetch_exchange_rates(year)

  with open("INFO.SRU", mode="w") as outfile:
    outfile.write("#DATABESKRIVNING_START\n")
    outfile.write("#PRODUKT SRU\n")
    outfile.write("#FILNAMN BLANKETTER.SRU\n")
    outfile.write("#DATABESKRIVNING_SLUT\n")
    outfile.write("#MEDIELEV_START\n")
    outfile.write("#ORGNR " + args.org_nummer + "\n")
    outfile.write("#NAMN " + args.fullt_namn + "\n")
    outfile.write("#ADRESS " + args.adress + "\n")
    outfile.write("#POSTNR " + args.postnummer + "\n")
    outfile.write("#POSTORT " + args.postort + "\n")
    outfile.write("#EMAIL " + args.epost + "\n")
    outfile.write("#MEDIELEV_SLUT\n")

  now = datetime.now()

  total_vinst_sek = 0
  total_forlust_sek = 0
  total_ordinary_income_sek = 0

  with open("BLANKETTER.SRU", mode="w") as outfile:
    chunk_counter = 1
    for chunk in chunk_list(trades, SALES_PER_PAGE):
      counter = 10
      outfile.write("#BLANKETT K4-" + year + "P4\n")
      outfile.write("#IDENTITET " + args.org_nummer + " " + now.strftime('%Y%m%d %H%M%S') + "\n")
      outfile.write("#NAMN " + args.fullt_namn + "\n")
      for trade in chunk:
        if counter - 31 > SALES_PER_PAGE:
          raise ValueError("Can only have %d sales per page" % SALES_PER_PAGE)

        if use_yearly_rate:
            exchange_rate = get_yearly_rate(rates)
        else:
            exchange_rate = get_daily_rate(rates, trade.date)

        vinst_sek = trade.vinst() * exchange_rate
        forlust_sek = trade.forlust() * exchange_rate
        ordinary_income_sek = trade.ordinary_income * exchange_rate
        total_vinst_sek += vinst_sek
        total_forlust_sek += forlust_sek
        total_ordinary_income_sek += ordinary_income_sek

        outfile.write("#UPPGIFT 3" + str(counter) + "0 " + str(trade.antal) + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "1 " + trade.beteckning + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "2 " + str(round(trade.forsaljsningspris * exchange_rate)) + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "3 " + str(round(trade.omkostnadsbelopp * exchange_rate)) + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "4 " + str(round(vinst_sek)) + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "5 " + str(round(forlust_sek)) + "\n")
        counter += 1
      outfile.write("#UPPGIFT 7014 " + str(chunk_counter) + "\n")
      outfile.write("#BLANKETTSLUT\n")
      chunk_counter += 1
    outfile.write("#FIL_SLUT\n")

  print("K4 SRU files generated. Year: {0}, Rate: {1}".format(year, "yearly" if use_yearly_rate else "daily"))
  print("Summa vinst SEK: " + swedish_float(total_vinst_sek))
  print("Summa förlust SEK: " + swedish_float(total_forlust_sek))
  print("Nettovinst/Nettoförlust SEK: " + swedish_float(total_vinst_sek - total_forlust_sek))
  print("Summa ordinarie inkomst SEK: " + swedish_float(total_ordinary_income_sek))



def chunk_list(lst, size):
  return [lst[i:i + size] for i in range(0, len(lst), size)]


def swedish_float(f):
  # Format with thousands separators and Swedish decimal comma
  formatted = f"{f:,.2f}".replace(',', ' ').replace('.', ',')
  return formatted


if __name__ == "__main__":
  main()
