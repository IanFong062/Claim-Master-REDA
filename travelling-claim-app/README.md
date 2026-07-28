# Travelling Claim Web App

Mobile-first travelling claim app for iPhone Safari, iPad, Android, Windows and Mac. Claim data stays in the browser. No login, no server upload, no claim data leaves the device.

## What It Does

- Create, edit, delete, duplicate and reorder trips.
- Auto-save draft in local browser storage.
- Save generated claims to local History for reopening and re-export.
- Remember common From, To, Client and Project Code suggestions.
- Preview all claim details before Excel generation.
- Block export clearly if trips exceed the template capacity.
- Generate `.xlsx` by copying the original company template package and patching only mapped cells.
- PWA support: `manifest.json`, service worker, app icon placeholder and offline app/template cache.

## Template Analysis

Template file: `templates/Travelling claim form - updated version.xlsx`

SHA256: `a3d78782754588db9223c0e751d8e892f9fc8cc282fa1cc8027c9d2f022e23e9`

Sheets:

| Sheet | Used Range | Merged Cells | Formula Cells |
|---|---:|---:|---:|
| Sheet1 | A1:N79 | 57 | 0 |
| Sheet2 | A1:A1 | 0 | 0 |
| Sheet3 | A1:A1 | 0 | 0 |

Trip input rows:

- Page 1: rows 7 to 27
- Page 2: rows 34 to 56
- Maximum trips: 44

Field mapping is centralized in `config/templateMapping.js`.

Important mapped cells:

| Field | Cell |
|---|---|
| Employee Name / Number | I1 |
| Address | B2 |
| Week Start / Week Ending | B1 |
| Claim Date / Remarks | B3 |
| Date | B row |
| From + To | C row, because C:D is merged |
| Time From | E row |
| Time To | F row |
| Client | G row |
| Project Code | H row |
| MTR | I row |
| Bus | J row |
| Van | K row |
| Taxi | L row |
| Sub Total | M row |
| Deduction | N row |
| Page totals | N29, N58 |

## Known Template Limitations

- The supplied template has no separate From and To cells. Each trip row merges C:D, so export writes one combined `From -> To` value into C.
- The supplied template has no Other transport column. Other trips are annotated in the location text and included in Sub Total.
- The supplied template has no row-level Remarks export cell. Trip remarks remain in draft/history/preview.
- No worksheet image/logo object was detected in the supplied workbook, but all existing package parts are preserved.
- Automated tests verify Excel workbook integrity with `openpyxl`. They do not open Apple Numbers or Microsoft 365 directly.

## Windows: Run Locally

From PowerShell:

```powershell
cd C:\Users\user\Documents\Work\travelling-claim-app
python -m http.server 5173
```

Open:

```text
http://localhost:5173/
```

If Python is not installed, install current Python or use any static web server.

## Mac: Run Locally

```bash
cd /path/to/travelling-claim-app
python3 -m http.server 5173
```

Open:

```text
http://localhost:5173/
```

## iPhone Safari

1. Start the local server on your computer.
2. Make sure iPhone and computer are on the same Wi-Fi.
3. Find your computer LAN IP.
4. Open Safari on iPhone:

```text
http://YOUR_COMPUTER_IP:5173/
```

5. Use Share -> Add to Home Screen.
6. After first load, the app shell and template are cached for offline use.
7. When Excel is generated, Safari can save the `.xlsx` to Files.

## Deploy

GitHub Pages:

1. Push this folder to a GitHub repository.
2. In repository Settings -> Pages, deploy from the main branch.
3. Set the source folder to the repository root.
4. Open the Pages URL.

Netlify:

1. Drag and drop this folder into Netlify, or connect the Git repo.
2. Build command can be blank.
3. Publish directory should be the project root.

Local LAN:

```powershell
python -m http.server 5173 --bind 0.0.0.0
```

Then open `http://YOUR_COMPUTER_IP:5173/` from phone/tablet.

## Test

Using the bundled Codex Python runtime in this workspace:

```powershell
$env:PYTHONIOENCODING='utf-8'
$env:PYTHONPATH='C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\python'
& 'C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' tests\run_export_tests.py
```

With normal Python and `openpyxl` installed:

```bash
python tests/run_export_tests.py
```

The test output is written to `TEST_REPORT.md`. Test `.xlsx` files are written to `outputs/tests/`.

## Replace Company Template

1. Replace both files with the new template:

```text
templates/Travelling claim form - updated version.xlsx
public/templates/Travelling claim form - updated version.xlsx
```

2. Run:

```bash
python scripts/analyze_template.py
```

3. Update `config/templateMapping.js`:

- `templateSha256`
- sheet names
- used ranges
- merged cells
- trip row sections
- field cells
- transport columns
- total cells
- max trip count

4. Run tests:

```bash
python tests/run_export_tests.py
```

Do not edit the source template during export. The app copies and patches an exported workbook only.

## Build Notes

This project is intentionally dependency-light and runs as a static PWA. `package.json` includes Vite/React-friendly metadata for future migration, but V1 does not require npm install to run. The Excel export code uses browser ZIP/XML patching to preserve the original workbook package as much as possible.

If you later migrate to React + Vite:

```bash
npm install
npm run dev
```

Keep `config/templateMapping.js` as the single source for cell addresses.
