# FlexPrint Export via GitHub Actions (utan din fysiska dator)

Den här lösningen kör exporten i GitHub Actions och:
- laddar upp resultatet som artifacts
- publicerar senaste `.xls` till `docs/latest/` (för GitHub Pages)

## 1) Lägg detta i ett gemensamt repo (organisation/team)

Lägg in:
- `flexprint_order_export.py`
- `.github/workflows/flexprint-export.yml`

Tips: använd ett org-repo istället för privat repo i ditt personliga konto.

## 2) Sätt secrets i repot

I GitHub:
`Settings -> Secrets and variables -> Actions -> New repository secret`

Skapa:
- `FLEXPRINT_USER`
- `FLEXPRINT_PASS`

## 3) Kör workflow manuellt första gången

Gå till:
`Actions -> FlexPrint Export -> Run workflow`

Valfritt:
- `view` (t.ex. `all` eller `inprocess`)
- `max_orders` (`0` = alla)

## 4) Hämta filerna

När run är klar:
`Actions -> run -> Artifacts`

Filer:
- `flexprint_export_full.xlsx`
- `flexprint_export_full.xls`
- `flexprint_orders_full_retry.csv`
- `flexprint_order_items_full_retry.csv`

## 5) Schemalagd körning

Workflowen kör vardagar enligt cron:
- `04:15 UTC` (i workflow-filen)

Ändra vid behov i `.github/workflows/flexprint-export.yml`.

## Viktigt att känna till

- Körningen är inte beroende av din dator.
- Körningen är inte bunden till att du personligen är inloggad lokalt.
- Om FlexPrint begränsar inloggning via IP/brandvägg kan GitHub-hosted runners blockeras.
  I så fall behövs self-hosted runner i ert nät.
- Om repot är publikt blir även den publicerade `docs/latest/flexprint_export_full.xls` publik.
