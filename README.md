# FlexPrint Order Export

Exporterar orderhistorik från FlexPrint till:
- CSV på ordernivå
- CSV på item-/detaljnivå
- XLSX med två flikar (`Orders`, `Items`)

## Lokalt

```bash
python flexprint_order_export.py \
  --output flexprint_orders_full_retry.csv \
  --items-output flexprint_order_items_full_retry.csv \
  --xlsx-output flexprint_export_full.xlsx
```

Miljövariabler:
- `FLEXPRINT_USER`
- `FLEXPRINT_PASS`

## GitHub Actions

Workflow finns i:
- `.github/workflows/flexprint-export.yml`

Setup-guide:
- `GITHUB_ACTIONS_SETUP.md`

