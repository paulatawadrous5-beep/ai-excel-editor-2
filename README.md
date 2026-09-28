# AI Excel Editor

Describe spreadsheet changes in plain English and get back a real `.xlsx`. Two modes:

- **Edit a file** – upload an `.xlsx`, describe changes, review the planned changes, apply, download.
- **Create from scratch** – describe a workbook, review the plan, generate, download.

The AI never runs code. It only returns a JSON list of **structured operations** from a fixed whitelist; the server validates them, applies them with ExcelJS, then **re-opens the result and verifies** the changes actually landed.

```
Instruction → AI (JSON operations only) → schema validation → ExcelJS engine
            → native chart injection → verification (+1 retry) → .xlsx download
```

## Features
- Sheets: create / rename / delete / reorder (reorder is best-effort, see limitations)
- Cells: set, ranges, clear, copy, move, find & replace
- Rows/columns: insert, delete, delete empty rows, heights, widths, auto-fit
- Formatting: bold/italic/underline, font size/family/color, fill, borders, alignment, wrap, number formats (currency, %, date, decimal)
- Structure: merge/unmerge, freeze panes, tables, sort
- Formulas: any formula body (SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, IF, SUMIF, COUNTIF…), Total column, Total row
- Charts: **native, editable Excel charts** – column, bar, line, pie
- Hyperlinks, data validation, conditional formatting, page setup
- Change preview before applying; progress messages; success/error states; responsive UI
- Security: `.xlsx` extension + zip magic-byte check, size limit, memory uploads (client filename never touches disk), UUID temp names, path-traversal guard, Helmet, CORS allow-list, rate limiting, no stack traces to clients, API key server-side only, temp files auto-deleted

## Tech stack
Node.js ≥ 18, Express, ExcelJS, JSZip (native chart XML), Multer, Helmet, express-rate-limit. Frontend: plain HTML/CSS/JS (no build step). Tests: Jest + Supertest.

## Project structure
```
server.js                  entry point
src/config.js              env config
src/routes/                edit, create, download endpoints
src/services/aiService.js  AI provider (isolated – swap here)
src/services/excelEngine.js  operation executor
src/services/chartEngine.js  native chart injection
src/services/verifier.js   post-apply verification
src/services/workbookAnalyzer.js  workbook summary sent to the AI
src/operations/            schema + validator (the security boundary)
src/middleware/            upload, rate limit, errors
src/utils/                 file + cell-reference helpers
public/                    frontend
tests/                     Jest suites
```

## Setup
```bash
npm install                # also generates package-lock.json - commit it
cp .env.example .env      # then set ANTHROPIC_API_KEY
npm start                 # http://localhost:3000
```

### Environment variables
See `.env.example`. **Required:** `ANTHROPIC_API_KEY` (an external AI API is needed; get a key at console.anthropic.com). Optional: `AI_MODEL`, `PORT`, `ALLOWED_ORIGINS`, `MAX_UPLOAD_MB`, `TEMP_DIR`, `FILE_TTL_MINUTES`, rate-limit settings. `AI_MODEL` defaults to `claude-sonnet-5`; change it if your account uses another model. To use another provider, replace `callAnthropic()` in `src/services/aiService.js`.

## Verification & reporting
After applying, the file is re-opened and each change is checked. A change only counts as **verified** if the re-opened file proves it. Applied-but-unconfirmed changes are **failed** (after one retry for idempotent operations). Change types with no simple invariant (e.g. sort, insert rows) are reported as **unable to verify**. The UI and API report all three counts and a status of `complete`, `complete_with_unverified`, `partial` or `failed`; if nothing verified or was unverifiable, no file is delivered.

## Delivery status (honest note)
This package was built in a sandbox with no npm access. Verified there: syntax of every file, module imports, the native chart XML (opened in openpyxl and LibreOffice), the validator, status logic and retry-safety logic (with stubs). **Not yet run:** `npm install` (so no `package-lock.json` is included - run it once and commit the file), the Jest suites, `npm start`, and the live AI call. Run `npm install && npm test -- --runInBand && npm start` first and fix anything that surfaces.

## Testing
```bash
npm test
```
Covers: cell refs, validator, Excel engine (sheets, cells, rows/cols, formatting, formulas, merges, freeze), chart XML for all four chart types, verifier, AI response parsing (mocked provider), and the full HTTP workflows for both modes (upload validation, plan, validate, apply, verify, download, path traversal).

## Deployment
Any Node host (Render, Railway, Fly, a VPS): set env vars, run `npm install && npm start`. Sessions and temp files are in-process/on local disk, so run a single instance (or move `sessionStore` to Redis and files to object storage to scale out). Serve behind HTTPS and set `ALLOWED_ORIGINS` to your domain.

GitHub: `git init && git add . && git commit -m "AI Excel Editor" && git push` (`.env` and `tmp/` are git-ignored).

## Troubleshooting
- *"AI is not configured"* → set `ANTHROPIC_API_KEY` in `.env`.
- *AI provider rejected the request* → check the key and `AI_MODEL`.
- *File rejected* → must be a genuine `.xlsx` under `MAX_UPLOAD_MB`.
- *"Session expired"* → sessions last `FILE_TTL_MINUTES`; re-upload.

## Known limitations
- The AI needs an external API and network access; results depend on the model's plan. The preview step lets you check the plan first.
- Sorting rewrites values only (per-cell styles/formulas in the sorted block are not moved).
- Sheet reorder is best-effort (ExcelJS has no public reorder API).
- ExcelJS may not preserve everything in complex files (e.g. existing pivot tables, form controls, some chart formatting). Always keep your original – it is never modified, but review the output.
- Charts assume the range's first row is headers and first column is categories; pie charts use one series.
- Formulas are written without cached results; Excel calculates them when the file opens.
- No login/multi-user isolation beyond unguessable session ids.
- Verification checks representative cells per operation, not every cell.
