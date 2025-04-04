Genererar K4 SRU från XLSX nedladdad från Etrade:
https://us.etrade.com/etx/sp/stockplan#/myAccount/gainsLosses

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
  --trades gains_loss_2023.xlsx \
  --rate-type year  # Använd detta för årliga kurser
  # ELLER
  --rate-type day   # Använd detta för dagliga kurser
```

Notera: Scriptet hämtar valutakurser direkt från Riksbankens API. Se till att du har internetanslutning när du kör scriptet.