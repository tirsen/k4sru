Genererar K4 SRU från XLSX nedladdad från Etrade:
https://us.etrade.com/etx/sp/stockplan#/myAccount/gainsLosses

Detta projekt använder [Hermit](https://github.com/cashapp/hermit) och [uv](https://github.com/astral-sh/uv). Sätt upp så här:
```
bin/hermit install
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