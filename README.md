# Excel & CSV Compare — browser-only version

The same comparison app, running **entirely inside the web browser**. You host it free on GitHub Pages and open it from any computer (office, home, tour) with a link. **Files are never uploaded**: they are read, compared and turned into the two Excel outputs inside the browser tab.

It gives the same results as the desktop (Python) version. It has the same seven steps, the same rules and the same two files, `File_B_Comparison_Highlighted.xlsx` and `Detailed_Before_After_Comparison.xlsx`, and every result is validated as PASS or FAIL.

---

## Deploy on GitHub Pages: step by step (browser only, no software to install)

You need a free GitHub account (github.com → *Sign up*).

1. **Unzip** `Excel_CSV_Compare_Web.zip` on your computer. You will see `index.html`, `package.json`, the folders `src`, `tests`, `sample_data`, and a folder named **`.github`**.
2. **Create a repository.** On github.com click **+** (top right) → **New repository**.
   - Name: `excel-csv-compare`
   - Choose **Public**. GitHub Pages is free only for public repositories. Only the app's code goes there, never your data.
   - Leave "Add a README" unticked → **Create repository**.
3. **Upload the files.** On the new empty repository page, click the link **"uploading an existing file"**. Select **everything inside** the unzipped folder, including the `.github` folder, and drag it onto the page. Wait for the list to finish, then click **Commit changes**.
   - Check that `.github/workflows/deploy.yml` appears in the file list. If it doesn't (some systems hide folders starting with a dot): click **Add file → Create new file**, type `.github/workflows/deploy.yml` as the name, paste in the contents of that file from the zip, and click **Commit changes**.
4. **Switch on Pages.** Go to **Settings → Pages**. Under *Build and deployment → Source*, choose **GitHub Actions**.
5. **Let it build.** Open the **Actions** tab. A run called *Deploy to GitHub Pages* starts automatically. If nothing is running, click it, then **Run workflow**. It takes about 2 minutes; wait for the green tick ✓.
6. **Open your app:** `https://<your-github-username>.github.io/excel-csv-compare/`. The link is also shown under Settings → Pages. Bookmark it and share it with colleagues.

**To update the app later,** upload the changed files the same way. GitHub rebuilds and republishes it automatically.

> Why GitHub builds it instead of you uploading a finished site: the comparison engine (DuckDB) is a 36 MB file, and GitHub's upload page accepts files only up to 25 MB. The workflow downloads the engine and builds the site on GitHub's side.

### Other ways to host it
`Excel_CSV_Compare_Web_built.zip` contains the finished site. Unzip it onto any web server: an office intranet server, IIS, Netlify (drag and drop), or a shared web folder. The site must be served over `http(s)://`, so double-clicking `index.html` from disk will not work. The server must send `.wasm` files as `application/wasm`; most servers already do.

---

## Using it
The steps are **Files → Sheets → Match → Rules → Compare → Results → Download**, the same as the desktop version.

- Add the two files by clicking or by drag-and-drop. Excel `.xlsx .xlsm .xls .xlsb` and CSV `.csv .txt .tsv` are supported, in any combination. Two sheets of one workbook also work.
- **Saved setups** are kept in that browser. On the Download step you can also click **Download as file** to get a `.json` setup, and open it on any computer with *Sheets → Saved setups → Open a setup file…*. A sample setup is in `sample_data`.
- The first visit downloads the comparison engine once (about 9 MB when the server compresses it, up to 36 MB if not). After that the browser keeps it.

## Privacy: how you can be sure files stay on your computer
- The page has a built-in security rule (`connect-src 'self'` in `index.html`). The browser itself refuses any attempt by the page to contact another website. There is no server code at all; GitHub only hosts the static files.
- To see it yourself: press F12 → **Network** tab, then run a comparison. No upload request appears.
- Once the page has loaded you can disconnect from the internet, and comparisons still work.

## Limits compared with the desktop version
| | Browser version | Desktop version |
|---|---|---|
| Install | none, just a link | Python on each PC, or one shared host PC |
| Works from | any computer with Chrome, Edge or Firefox | office PCs |
| Comfortable size | up to about 3–5 lakh rows per file (depends on the PC's memory) | 10+ lakh rows (disk-backed) |
| 1 lakh rows, CSV vs Excel | about 25–30 s | about 10–15 s |
| Typing a file path | not possible in a browser | yes |

For very large files (hundreds of MB, or more than about 5 lakh rows), use the desktop version.

Everything else is the same:
- exact sheet, row, column and cell references, including CSV blank lines and multi-line cells
- duplicate-key policies, blank keys, and every comparison rule
- highlights that match the report exactly, with File B's values never changed
- `.xlsx` copies kept part by part with styles and formulas; `.csv`, `.xls` and `.xlsb` converted to `.xlsx` with a clear note
- formula comparison, splitting across sheets past 1,048,576 rows, and the same validation checks

---

## For developers
```bash
npm ci
npm run dev                         # http://localhost:5173
npm run build                       # site in dist/
python -m pytest tests              # 50 tests (needs Python 3 + openpyxl; Chrome + playwright for the browser test)
```
- `src/lib/`: engine code that runs the same in the browser and in Node:
  - `loader.js` reads files (SheetJS 0.20.3 for Excel, its own CSV reader for exact row numbers)
  - `engine.js` does the matching in DuckDB SQL
  - `highlight.js` edits only the cell styles inside the File B package
  - `xlsxWriter.js` writes the outputs as a stream
  - `validate.js` runs the checks
  - `report.js` writes the report
- `src/ui/`, `src/main.jsx`: the screens (React).
- `tests/run_case.mjs` runs the browser engine in Node. The pytest files build real Excel/CSV files and check every output independently with openpyxl. `test_e2e.py` drives the built site in Chrome and also confirms the page made no requests to other sites.
