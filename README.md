# FlexPrint Order Export

Exporterar orderhistorik från FlexPrint till:
- CSV på ordernivå
- CSV på item-/detaljnivå
- XLSX med två flikar (`Orders`, `Items`)
- XLS med två flikar (`Orders`, `Items`)

## Lokalt

```bash
python flexprint_order_export.py \
  --output flexprint_orders_full_retry.csv \
  --items-output flexprint_order_items_full_retry.csv \
  --xlsx-output flexprint_export_full.xlsx \
  --xls-output flexprint_export_full.xls
```

Miljövariabler:
- `FLEXPRINT_USER`
- `FLEXPRINT_PASS`

## GitHub Actions

Workflow finns i:
- `.github/workflows/flexprint-export.yml`

Setup-guide:
- `GITHUB_ACTIONS_SETUP.md`

## HTML-sida för körning + nedladdning

En enkel webbsida finns i:
- `docs/index.html`

Aktivera GitHub Pages:
1. `Settings -> Pages`
2. Source: `Deploy from a branch`
3. Branch: `main` och folder: `/docs`

Sidan används för att:
- ladda ner senaste publicerade `docs/latest/flexprint_export_full.xls`
- öppna workflow-sidan för att starta ny export i GitHub Actions

Workflowen publicerar senaste `.xls` till:
- `docs/latest/flexprint_export_full.xls`
- `docs/latest/last_updated_utc.txt`

Viktigt: eftersom repot är publikt blir den publicerade `.xls`-filen också publik via Pages.
