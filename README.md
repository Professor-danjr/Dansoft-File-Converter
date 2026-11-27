# Dansoft File Converter (Vercel-ready, single file)

## Deploy
1. Upload `index.html` + `vercel.json` to your GitHub repo root (delete the old stray `app.js` — it is no longer needed, everything is inside index.html).
2. Vercel auto-redeploys from GitHub.

## Why single file?
GitHub uploads often flatten the `src/app.js` path, which broke the old build (`src/app.js` 404 → dead buttons). This build inlines all logic so path mistakes are impossible.
