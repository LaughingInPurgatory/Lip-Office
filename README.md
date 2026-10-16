# Lip Office

A small office suite that runs in a browser, served by Node.js.

The interface follows the system light/dark preference and includes a theme toggle;
your choice is remembered in this browser.

Two components:

- **Word processor** (`public/word.js`) — a Word-style editing surface with headings,
  bold/italic/underline/strikethrough, super/subscript, colour, bullet and numbered
  lists, indentation, alignment, hyperlinks, tables, rules, and undo/redo. Opens
  `.doc`, `.dot`, `.docx`, `.odt`, `.rtf`, `.txt` and `.html`; saves `.docx`, `.odt`,
  `.rtf`, `.txt` and `.html`.
- **Spreadsheet** (`public/sheets.js`) — a grid with a formula bar, keyboard navigation,
  multiple sheets, and formulas with cell/range references. Opens `.xls`, `.xlt`,
  `.xlsx`, `.ods` and `.csv`; saves `.xlsx`, `.ods` and `.csv`.

There are **no npm dependencies**. Node's built-in `zlib` is enough to read and write
the ZIP containers that `.docx`, `.xlsx`, `.odt` and `.ods` really are, and the OOXML
and OpenDocument XML inside them is parsed and generated directly. Nothing is uploaded
anywhere; documents live in `documents/` next to the server, or in a configurable host
directory when using Docker Compose.

## Running it

```sh
node server.js          # http://127.0.0.1:3000
PORT=8080 node server.js
```

`npm start` does the same thing. Requires Node 20 or newer (developed on 26).

```sh
npm test                # run the automated test suite
```

### Run with Docker Compose

The Compose setup publishes the app only on the host's loopback address, so it remains
local-only without login or HTTPS. By default, files persist in `./documents` and the
app is available at `http://127.0.0.1:3000`:

```sh
docker compose up -d --build
```

To put documents somewhere else on the server, set `DOCUMENTS_DIR` in a `.env` file next
to `compose.yaml` (use an absolute path), for example:

```dotenv
DOCUMENTS_DIR=/srv/lip-office/documents
LIP_OFFICE_PORT=3000
PUID=1000
PGID=1000
```

Make sure that directory exists and is writable by the configured `PUID` and `PGID`.
Then run `docker compose up -d --build`. The container runs as that unprivileged user,
uses a read-only root filesystem, and drops Linux capabilities.

The GitHub Actions workflow runs the test suite and publishes `ghcr.io/<owner>/<repo>`
on pushes to `main` and `v*` tags. To use the published image, set
`LIP_OFFICE_IMAGE=ghcr.io/<owner>/<repo>:latest`, run `docker compose pull`, then
`docker compose up -d`. GitHub Container Registry packages are private by default when
first published; change the package visibility in GitHub if anonymous pulls are wanted.

### Install on ZimaOS

Use [docker-compose.zimaos.yml](docker-compose.zimaos.yml) as a customized app Compose
file in ZimaOS. Before installing, replace `YOUR_ZIMAOS_LAN_IP:3000` with the NAS's LAN
address and port, create `/DATA/AppData/lip-office/documents`, and make it writable by
UID/GID `1000:1000`. The file publishes port 3000 to the LAN. Lip Office has no login or
HTTPS, so only use this on a trusted network. The GHCR image is private; authenticate
the ZimaOS Docker host to `ghcr.io` with a GitHub token that has `read:packages` before
pulling it. Do not put that token in the Compose file.

## Using it

Pick **Word processor** or **Spreadsheet**, then:

- **New** — start an empty file.
- **Import…** — open a file from your computer. It is converted in the browser session
  and is not stored until you save.
- **Save** — write it to `documents/` in the format chosen in the dropdown. The
  extension follows the format, so naming it `report` and choosing `.ods` saves
  `report.ods`.
- **Download** — fetch the stored file as real `.docx`/`.xlsx`/… bytes.
- `Cmd`/`Ctrl`+`S` saves. Clicking a name in the sidebar reopens it; `×` deletes it.

In the word processor, select text and use the toolbar. In the spreadsheet, arrows move,
typing starts an edit, `Enter` moves down, `Tab` moves right, `F2` edits in place,
`Escape` cancels, `Delete` clears a cell, and double-clicking a sheet tab renames it.

### Supported formulas

The engine covers common aggregate and conditional functions (`SUM`, `AVERAGE`, `COUNT`,
`SUMIF(S)`, `COUNTIF(S)`, `AVERAGEIF(S)`, `SUMPRODUCT`), lookups (`VLOOKUP`, `HLOOKUP`,
`XLOOKUP`, `INDEX`, `MATCH`), logical functions, common text functions, date/time
functions, and common math/statistics functions. See `functionNames()` in
`shared/formula.js` for the current list. Arithmetic, text concatenation, percent, and
comparison operators are supported too.

This is not full Excel compatibility. Cross-sheet references, named ranges, structured
table references, dynamic array formulas, and specialized statistical/financial functions
aren't implemented. Formulas using those features may display `#NAME?` or `#ERROR!`.

Errors surface as `#DIV/0!`, `#VALUE!`, `#NAME?`, `#REF!`, `#NUM!`, `#N/A`, `#ERROR!`
and `#CYCLE!`.

## How it is put together

```
server.js            HTTP server: static files, JSON API, no framework
lib/
  zip.js             ZIP reader/writer (zlib + own CRC32)
  ooxml.js           .docx and .xlsx read/write
  odf.js             .odt and .ods read/write
  rtf.js             .rtf read/write
  ole.js             OLE2 compound-file reader for legacy Office files
  legacy.js          text and cell extraction from Word/Excel binary streams
  convert.js         format dispatch: which reader/writer for which extension
  store.js           document storage, and the path-traversal guard
shared/              used by both the browser and Node
  html.js            tolerant markup parser (also reads OOXML/ODF, which is XML)
  sanitize.js        tag/attribute whitelist
  formula.js         formula engine
  csv.js             CSV
  model.js           document shapes, validation, file-type tables
public/              index.html, app.css, app.js, word.js, sheets.js
documents/           your files
```

The word processor's document format is HTML, and the spreadsheet's is
`{sheets: [{name, cells}]}`. Every file format is converted to and from those two
shapes, which is why all the converters are small and share one markup parser.

The API is five routes:

| Route | Purpose |
| --- | --- |
| `GET /api/documents` | list stored files |
| `GET /api/documents/:name` | open one, converted to the internal shape |
| `GET /api/documents/:name?download=1` | raw file bytes, as an attachment |
| `PUT /api/documents/:name` | save `{kind, format, doc}` |
| `POST /api/parse` | convert an uploaded file in memory, store nothing |
| `DELETE /api/documents/:name` | delete |

## Security notes

- **Local use only.** The app has no login or HTTPS, so native mode binds to `127.0.0.1`
  and refuses non-local bindings. The Docker image listens inside its container, while
  the standard Compose file publishes only on host loopback. The ZimaOS-specific Compose
  file publishes to the LAN and restricts accepted hostnames to the configured NAS
  address; only use it on a trusted network. Do not expose it to the public internet.
- Responses include a restrictive Content Security Policy and browser security headers.
  Mutating API requests must have a same-origin `Origin` header; cross-site requests are
  rejected. Request headers and upload time are bounded.
- Document names from the network are reduced to a basename and checked against a strict
  whitelist before touching the filesystem.
- Imported HTML is sanitised against a tag/attribute whitelist, both on the client when
  pasting and on the server before storage, because stored documents are later rendered
  with `innerHTML`. Scripts, event handlers, `javascript:` URLs and `url()`/`expression()`
  CSS are stripped.
- Uploads are capped at 32 MB.
- ZIP-based office files are limited to 128 MB per expanded entry and 256 MB total
  expanded content, protecting the server from highly compressed uploads.

## Known limits

- **Legacy binary `.doc`, `.dot`, `.xls` and `.xlt` import uses built-in readers.**
  Word import extracts document text and paragraph breaks. Excel import reads common
  BIFF5/BIFF8 cell values, shared strings and cached formula results. Legacy formatting,
  images, macros and unsupported BIFF records are not preserved; save to `.docx`, `.odt`,
  `.xlsx` or `.ods` to keep editing in a modern format.
- **Fidelity is good, not pixel-perfect.** The OOXML and ODF writers cover text,
  headings, bold/italic/underline/strikethrough, super/subscript, colour, font size,
  alignment, bullet and numbered lists, tables, hyperlinks, headers, footers, one to
  three text columns, Letter/A4 paper, portrait/landscape orientation, and common margins.
  `.docx` and `.odt` preserve embedded PNG/JPEG/GIF images; `.xlsx`
  and `.ods` preserve images anchored to cells. RTF, TXT, CSV, and legacy binary imports
  do not preserve images or page layout. Footnotes, comments, tracked changes, conditional
  formatting, charts, and cell number formats are not supported. Reading is broader than
  writing.
- **Images are limited to PNG, JPEG, and GIF**, up to 8 MB each. SVG and externally
  linked images are rejected for safety.
- **Font size and indentation are not settable from the toolbar** — headings cover the
  size case. Read-side font sizes and indents are preserved.
- **Dates read as text.** Excel stores dates as serial numbers, so on import they are
  converted to `YYYY-MM-DD` for display and become text cells; date arithmetic on them
  will not work. The 1904 date system is not handled.
- **ODS repeated-row/column runs are expanded up to the grid limit**, not expanded in
  place, and the sheet cap is 5000 rows by 256 columns.
- **CSV is single-sheet.** Saving a multi-sheet workbook as `.csv` writes the first sheet;
  the others and any images are not in the file.
- **The grid recomputes the whole visible sheet on every edit.** That is imperceptible at
  the default 100×26 and fine into a few thousand cells; `public/sheets.js` has a
  `shortcut:` note marking where incremental recalculation would go.

## Credits

Written as a single Node.js application. Apple's `textutil` and QuickLook were used
while developing to confirm the generated `.docx`/`.xlsx`/`.odt`/`.ods` files open in
readers other than this one.
