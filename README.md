# LG Buying Guides Dashboard 2.0

Improved static HTML rebuild of the live dashboard at [hilldongukyim.github.io/lg-buying-guides-dashboard](https://hilldongukyim.github.io/lg-buying-guides-dashboard/).

## What changed

- Split the original single-file app into `index.html`, `styles.css`, and modular JS under [`src/`](/Users/vbond/dashboard/src).
- Added a normalization layer in [`src/data-model.js`](/Users/vbond/dashboard/src/data-model.js) to fix duplicate country code variants such as `CA-EN` and `CA_EN`.
- Removed header rows from `pages` when they are mistakenly included as data.
- Kept the same high-level UX:
  - intelligence brief
  - country cards
  - category compare
  - raw data tabs
  - page links drawer
  - theme/lang toggles

## Run locally

This project should be served through a local web server because it fetches `./data/data.json`.

Example:

```bash
cd /Users/vbond/dashboard
python3 scripts/dashboard_server.py
```

Then open <http://localhost:8000>.

## Regenerate from Excel

The dashboard now supports an offline preprocessing step that converts the real workbook into the JSON snapshot the UI reads.

When running through `scripts/dashboard_server.py`, the `Refresh Data` button calls `/api/refresh` and rebuilds `data.json` from:
- [`data/monthly_ga_lineupguide.xlsx`](/Users/vbond/dashboard/data/monthly_ga_lineupguide.xlsx)
- [`data/monthly_ga_featurelibrary.xlsx`](/Users/vbond/dashboard/data/monthly_ga_featurelibrary.xlsx)
- [`data/page-links.csv`](/Users/vbond/dashboard/data/page-links.csv)

Command:

```bash
cd /Users/vbond/dashboard
python3 scripts/build_dashboard_data.py
```

Custom paths:

```bash
python3 scripts/build_dashboard_data.py \
  --lineup-workbook /Users/vbond/dashboard/data/monthly_ga_lineupguide.xlsx \
  --feature-workbook /Users/vbond/dashboard/data/monthly_ga_featurelibrary.xlsx \
  --pages-csv /Users/vbond/dashboard/data/page-links.csv \
  --live-snapshot /Users/vbond/dashboard/data/live-data.raw.json \
  --output /Users/vbond/dashboard/data/data.json
```

## Data notes

- Local snapshot is stored at [`data/data.json`](/Users/vbond/dashboard/data/data.json).
- Page Links data can be edited in [`data/page-links.csv`](/Users/vbond/dashboard/data/page-links.csv).
- The full default dashboard snapshot with both `feature_library` and `lineup` is preserved from the live site and normalized into [`data/data.json`](/Users/vbond/dashboard/data/data.json).
- The unmodified live snapshot backup is stored at [`data/live-data.raw.json`](/Users/vbond/dashboard/data/live-data.raw.json).
- Source workbook has also been copied into the project at [`data/ga_montly_lineup_feature_guide.xlsx`](/Users/vbond/dashboard/data/ga_montly_lineup_feature_guide.xlsx).
- Country normalization is centralized in [`src/config.js`](/Users/vbond/dashboard/src/config.js) and applied in [`src/data-model.js`](/Users/vbond/dashboard/src/data-model.js).
- Canonical Canada locale codes are:
  - `CA-EN`
  - `CA-FR`

## Real workbook hookup

- Current UI reads the normalized JSON snapshot at [`data/data.json`](/Users/vbond/dashboard/data/data.json).
- `Page Links` are loaded from [`data/page-links.csv`](/Users/vbond/dashboard/data/page-links.csv) first. If the CSV is missing, the build falls back to [`data/live-data.raw.json`](/Users/vbond/dashboard/data/live-data.raw.json).
- Real workbook file is stored beside it at [`data/ga_montly_lineup_feature_guide.xlsx`](/Users/vbond/dashboard/data/ga_montly_lineup_feature_guide.xlsx).
- Workbook parsing is handled by [`scripts/generate_data.py`](/Users/vbond/dashboard/scripts/generate_data.py).
- Final merged dashboard snapshot generation is handled by [`scripts/build_dashboard_data.py`](/Users/vbond/dashboard/scripts/build_dashboard_data.py).
- Live snapshot normalization for the original page structure is handled by [`scripts/normalize_live_snapshot.py`](/Users/vbond/dashboard/scripts/normalize_live_snapshot.py).
- Parsing rules and normalization config live in [`scripts/workbook_config.py`](/Users/vbond/dashboard/scripts/workbook_config.py).
- `CA_EN` and `CA_FR` are normalized to canonical dashboard codes:
  - `CA-EN`
  - `CA-FR`
