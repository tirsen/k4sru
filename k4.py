import argparse

import openpyxl
from datetime import datetime
from dataclasses import dataclass

SALES_PER_PAGE = 7

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
parser.add_argument('--year', required=True,
                    help='vilket räkenskapsår')
parser.add_argument('--valutakurs', default=1, type=float,
                    help='om indatat är i annan valutakurs, se https://www.riksbank.se/sv/statistik/rantor-och-valutakurser/valutakurser-till-deklarationen/')
parser.add_argument('--trades', required=True,
                    help='affärer i xlsx')


@dataclass
class Trade:
  antal: int
  beteckning: str
  forsaljsningspris: float
  omkostnadsbelopp: float
  gainloss: float

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
                  gainloss=row[18].value)


def main():
  # https://www.skatteverket.se/foretag/etjansterochblanketter/blanketterbroschyrer/broschyrer/info/269.4.39f16f103821c58f680007305.html
  args = parser.parse_args()
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

  trades = list(parse_trades(args.trades))
  now = datetime.now()

  with open("BLANKETTER.SRU", mode="w") as outfile:
    chunk_counter = 1
    for chunk in chunk_list(trades, SALES_PER_PAGE):
      counter = 31
      outfile.write("#BLANKETT K4-" + args.year + "P4\n")
      outfile.write("#IDENTITET " + args.org_nummer + " " + now.strftime('%Y%m%d %H%M%S') + "\n")
      outfile.write("#NAMN " + args.fullt_namn + "\n")
      for trade in chunk:
        if counter - 31 > SALES_PER_PAGE:
          raise ValueError("Can only have %d sales per page" % SALES_PER_PAGE)
        outfile.write("#UPPGIFT 3" + str(counter) + "0 " + str(trade.antal) + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "1 " + trade.beteckning + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "2 " + str(round(trade.forsaljsningspris * args.valutakurs)) + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "3 " + str(round(trade.omkostnadsbelopp * args.valutakurs)) + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "4 " + str(round(trade.vinst() * args.valutakurs)) + "\n")
        outfile.write("#UPPGIFT 3" + str(counter) + "5 " + str(round(trade.forlust() * args.valutakurs)) + "\n")
        counter += 1
      outfile.write("#UPPGIFT 7014 " + str(chunk_counter) + "\n")
      outfile.write("#BLANKETTSLUT\n")
      chunk_counter += 1
    outfile.write("#FIL_SLUT\n")


def chunk_list(lst, size):
  return [lst[i:i + size] for i in range(0, len(lst), size)]


def swedish_float(f):
  return str(f).replace('.', ',')


if __name__ == "__main__":
  main()
