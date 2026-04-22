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

Valfritt för e-post (Gmail SMTP):
- `SMTP_USER` (din Gmail-adress)
- `SMTP_PASS` (Google App Password, inte vanliga lösenordet)
- `MAIL_FROM` (avsändare, normalt samma som `SMTP_USER`)
- `MAIL_TO` (mottagare; flera kan separeras med kommatecken)

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
- `report_latest.csv`
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

## 6) E-post via Gmail SMTP (valfritt)

Om `SMTP_USER`, `SMTP_PASS`, `MAIL_FROM`, `MAIL_TO` finns:
- workflowen skickar mejl vid körning
- ämnesrad: `Månadsrapport - Astice`
- mejlet beskriver datumfiltreringen (inkluderat/bortfiltrerat)
- mejlet innehåller inte tabellen
- `report_latest.xls` bifogas i mejlet

Google-krav:
- aktivera 2-stegsverifiering på Google-kontot
- skapa ett App Password
- använd App Password i `SMTP_PASS`

Tips:
- `MAIL_TO` kan vara flera adresser separerade med kommatecken
- exempel: `anna@bolag.se, bo@bolag.se`

## Viktigt

- Om källsystemet kräver IP-whitelist kan GitHub-hosted runners blockeras.
- Om repot är publikt blir även `docs/latest/report_latest.xls` publik.
