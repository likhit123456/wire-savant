# WireSavant

WireSavant is a small web tool that takes a URL or an uploaded document (PDF, Word, or PowerPoint) and gives you a quick, readable brief on it: what it is, whether it looks safe, what tech it's built on, and (for certain sites) live structured data pulled through Anakin's Wire layer.

It's a single FastAPI backend (`main.py`) serving one HTML/JS frontend (`index.html`). No build step, no database — everything runs from these two files plus a `.env` for your API keys.

## What it does

- **Single URL scan** — paste a URL, get a safety verdict, a short summary, an elevator pitch, detected tech stack, social links, and outbound links.
- **Batch scan** — scan up to 5 URLs at once and export the results as CSV or Markdown.
- **Document scan** — upload a PDF, DOCX, or PPTX (up to 15MB) and get the same kind of summary and safety check.
- **Compare mode** — pick npm or pypi, type in two package names, and see their live registry data side by side, pulled through Wire.
- **Wire enrichment** — for sites Wire has a catalog entry for (currently npm and pypi), the tool pulls structured data instead of relying only on scraped text. This shows up as a badge on the result and a small data panel.
- **Report export** — download any single result as Markdown, JSON, or a formatted Word document (.docx).
- **History** — your last 12 scans are remembered across page reloads (saved in your browser, not on the server).

## How it works, roughly

1. You give it a URL or a file.
2. For URLs: it tries to scrape the page (using Anakin's scraper, with a plain HTTP fallback if that fails), then runs a quick AI safety check, then asks an AI model for a summary. If the URL matches a known Wire service, it also pulls structured data for that service in parallel.
3. For documents: it extracts the text directly from the file, then runs the same safety check and summary steps.
4. Everything gets cached in memory by URL or by file content, so re-scanning the same thing twice is instant the second time.

The safety check and the summary both run on Groq using a small, fast model. Scraping and Wire data both run through Anakin's API.

## Setup

### Requirements

Python 3.10 or newer, plus the packages in `Requirements.txt`:

```
fastapi
uvicorn
httpx
groq
python-dotenv
python-multipart
python-docx
python-pptx
pypdf
```

Install them with:

```bash
pip install -r Requirements.txt
```

### API keys

You need two keys. Create a `.env` file in the project root:

```
ANAKIN_API_KEY=your_anakin_key_here
GROQ_API_KEY=your_groq_key_here
```

The app will refuse to start if either key is missing.

### File layout

Put `index.html` inside a `templates/` folder next to `main.py`:

```
project/
  main.py
  Requirements.txt
  .env
  templates/
    index.html
```

### Running it

```bash
python main.py
```

This starts the server on `http://0.0.0.0:8000` with auto-reload enabled. Open `http://localhost:8000` in a browser.

## Using each feature

### Single URL

Type or paste a URL into the main input and click Scan. You can also press Cmd/Ctrl+K to jump to the input, or just paste a URL anywhere on the page and it will start scanning automatically.

### Batch

Switch to the Batch tab, put one URL per line (up to 5), and click Scan All.

### Document

Switch to the Document tab and either drag a file onto the drop zone or click to browse. Supported types: PDF, DOCX, PPTX. Max size is 15MB.

### Compare

Switch to the Compare tab, choose npm or pypi, type a package name into box A and another into box B, then click Compare via Wire. Each package gets its own card with whatever fields Wire's catalog returns for that service (things like download counts, version, and license, when available). If a package isn't found, that card will say so instead of showing stats — the other card still renders normally.

This currently only works for npm and pypi, since those are the only two Wire services that have been confirmed to have live, working actions. A few other services (reddit, stackoverflow) have the parameter-handling code already written but haven't been verified against the live Wire catalog — see the comments near `WIRE_SERVICES` in `main.py` for the exact commands to check and turn one on.

### Exporting a report

After any single URL or document scan, use the action buttons under the result to:

- Copy the summary to your clipboard
- Export as Markdown
- Export as JSON
- Download a formatted Word document (.docx)
- Share a link that reloads the same scan (URL scans only)

### History

Your recent scans show up at the bottom of the page. Click any entry to re-run that scan (for documents, you'll need to re-upload the file since the file itself isn't stored). History is saved in your browser's local storage, so it survives a page refresh but won't follow you to a different browser or device. Use Clear to wipe it.

## API endpoints

If you want to call the backend directly instead of using the page:

| Method | Path | Purpose |
|---|---|---|
| GET/POST | `/analyze?url=...` | Scan a single URL |
| POST | `/batch` | Scan up to 5 URLs at once |
| POST | `/scan-file` | Scan an uploaded document |
| POST | `/compare` | Compare two npm or pypi packages via Wire |
| POST | `/export-docx` | Generate a Word document from a result |
| GET | `/stats` | Total counts of URLs and files scanned so far |
| GET | `/result?url=...` | Shareable link that reloads a past scan |

## Known limitations

- History is stored in the browser, not the server, so it's per-browser and will be lost if you clear site data.
- Wire enrichment only covers npm and pypi right now. Other services need to be verified against the live Wire catalog before they're added.
- The AI safety check and summary are both generated by a language model, so they should be treated as a fast first read, not a guarantee. Always use your own judgment on anything flagged as risky, and don't treat the summary as a substitute for reading the actual source.
- Caching is in memory, so it resets every time the server restarts.

## Credits

Scraping: Anakin
Inference: Groq, running llama-3.1-8b-instant
