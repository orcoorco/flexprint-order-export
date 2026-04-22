# Order Export

Genererar:
- CSV på ordernivå
- CSV på item-/detaljnivå
- XLSX med en flik (`Report`)
- XLS med en flik (`Report`)

`Report` innehåller kolumnerna:
- `description` (från item-detaljer)
- `quantity` (från item-detaljer)
- `created` (från orderraden)

Datat begränsas till poster med `created` inom senaste två månaderna.

## Lokalt

```bash
python order_export.py \
  --base-url "https://example.com/espressi/" \
  --username "$EXPORT_USER" \
  --password "$EXPORT_PASS" \
  --output orders_latest.csv \
  --items-output order_items_latest.csv \
  --xlsx-output report_latest.xlsx \
  --xls-output report_latest.xls
```

Miljövariabler:
- `EXPORT_BASE_URL`
- `EXPORT_USER`
- `EXPORT_PASS`

## GitHub Actions

Workflow:
- `.github/workflows/data-export.yml`

Setup-guide:
- `GITHUB_ACTIONS_SETUP.md`

Valfritt:
- e-postnotis via Gmail SMTP (`SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `MAIL_FROM`, `MAIL_TO`)

## Webbsida

Sidan finns i:
- `docs/index.html`

Den:
- kräver ingen PAT
- visar status
- laddar ner senaste publicerade `.xls`

Publicerad fil:
- `docs/latest/report_latest.xls`
- `docs/latest/last_updated_utc.txt`

## Sökmotorer

Sidan är markerad för att inte indexeras med:
- `meta robots noindex,nofollow`
- `docs/robots.txt` med `Disallow: /`

Obs: det är en stark signal men ingen absolut garanti för att all indexering stoppas i alla motorer.
