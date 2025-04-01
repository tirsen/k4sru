Genererar K4 SRU från XLSX nedladdad från Etrade:
https://us.etrade.com/etx/sp/stockplan#/myAccount/gainsLosses

Hämta växelkurs från:
https://www.riksbank.se/sv/statistik/rantor-och-valutakurser/valutakurser-till-deklarationen/

Detta projekt använder [Hermit](https://github.com/cashapp/hermit). Sätt upp så här:
```
source .hermit/bin/activate
pip install -r requirements.txt
```

Kör så här:
```
python k4.py \
  --org-nummer <person nummer> \
  --fullt-namn '...' \
  --year 2023 \
  --adress '...' \
  --postnummer '...' \
  --postort '...' \
  --epost ... \
  --valutakurs '<växelkurs>' \
  --trades gains_loss_2023.xlsx
```