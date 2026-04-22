# Export via GitHub Actions (utan fysisk klientdator)

Lösningen kör exporten i GitHub Actions och:
- laddar upp resultat som artifacts
- publicerar senaste `.xls` i `docs/latest/` för nedladdning via Pages

## 1) Filer i repot

Säkerställ att följande finns:
- `order_export.py`
- `.github/workflows/data-export.yml`
- `docs/index.html`

## 2) Sätt secrets i repot

GitHub:
`Settings -> Secrets and variables -> Actions -> New repository secret`

Skapa:
- `EXPORT_BASE_URL`
- `EXPORT_USER`
- `EXPORT_PASS`

Valfritt för e-post (utan SMTP, via Resend API):
- `RESEND_API_KEY`
- `RESEND_FROM`
- `RESEND_TO`

## 3) Första körning

Gå till:
`Actions -> Data Export -> Run workflow`

Valfritt:
- `view` (t.ex. `all` eller `inprocess`)
- `max_orders` (`0` = alla)

## 4) Resultatfiler

Artifacts per körning:
- `report_latest.xlsx`
- `report_latest.xls`
- `orders_latest.csv`
- `order_items_latest.csv`

Publik fil (Pages):
- `docs/latest/report_latest.xls`

XLS/XLSX innehåller en flik `Report` med kolumner:
- `description` (från item-rader)
- `quantity` (från item-rader)
- `created` (från order)

Urval:
- poster där `created` ligger inom senaste två månaderna

## 5) Automatik

Workflowen körs:
- manuellt (`workflow_dispatch`)
- automatiskt sista tisdagen varje månad kl. 08 (Europe/Stockholm)
  - tekniskt triggas den tisdagar på två UTC-tider och har en inbyggd guard som bara kör i exakt månadfönstret

## 6) E-post utan SMTP (valfritt)

Om `RESEND_API_KEY`, `RESEND_FROM` och `RESEND_TO` finns:
- workflowen skickar mejl när `report_latest.xls` har uppdaterats
- mejlet innehåller länk till senaste filen och länk till körningen

Tips:
- `RESEND_TO` kan vara flera adresser separerade med kommatecken
- exempel: `anna@bolag.se, bo@bolag.se`

## Viktigt

- Om källsystemet kräver IP-whitelist kan GitHub-hosted runners blockeras.
- Om repot är publikt blir även `docs/latest/report_latest.xls` publik.
