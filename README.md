# File Converter (Vercel-ready)

Client-side file converter: xlsx to pdf, pdf to xlsx, xlsx to image, image to xlsx, pdf to image, image to pdf, plus csv/txt/png/jpg.

## Run locally
Open `index.html` in a browser (needs internet once for CDN libs).

## Deploy on Vercel
1. Push this folder to a GitHub repo (`index.html`, `src/app.js`, `vercel.json`).
2. Vercel → Add New → Project → Import from GitHub.
3. Framework: Other. Build command: empty. Output: `.`
4. Deploy.

## Notes
- All conversion runs in the browser; files never upload anywhere.
- Libs via CDN: SheetJS, jsPDF + autotable, pdf.js, JSZip. ExcelJS + Tesseract load lazily only for image→xlsx with OCR.
- 80MB per-file cap, PDF capped at 50 pages per conversion.
