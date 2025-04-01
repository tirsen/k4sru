Genererar K4 SRU från XLSX nedladdad från Etrade:
https://us.etrade.com/etx/sp/stockplan#/myAccount/gainsLosses

Ladda ner årlig/daglig USD/SEK som xlsx från
https://www.riksbank.se/sv/statistik/rantor-och-valutakurser/sok-rantor-och-valutakurser

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
  --yearly-rate-file yearly_rates.xlsx  # Använd detta för årliga kurser
  # ELLER
  --daily-rate-file daily_rates.xlsx    # Använd detta för dagliga kurser
```