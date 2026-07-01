# Seazen Work Hub

Your multi-tab marketing hub, packaged to run as a real website. Outside the
Claude sandbox the trend photos load normally, the live F&B feed works, and the
assistant works, all securely.

## What works once deployed
- Real trend photos load (no sandbox blocking images)
- Live "Refresh" pulls current F&B trends from the web
- The chat assistant and "Polish with AI" use your own Anthropic key
- Your data (work log, ideas, pins, etc.) is saved in the browser

## What you need
- A free GitHub account
- A free Vercel account (sign in with GitHub)
- An Anthropic API key from https://console.anthropic.com (Settings > API Keys)

---

## Deploy in about 10 minutes

### 1. Put the code on GitHub
- Create a new repository on GitHub (name it anything, e.g. `seazen-work-hub`).
- Upload the contents of this folder to it. Easiest on the web: open your new
  repo, click "uploading an existing file", drag in everything from this folder
  (keep the folder structure: `src/`, `api/`, and the config files), commit.

### 2. Import to Vercel
- Go to https://vercel.com and click "Add New > Project".
- Pick the GitHub repo you just created and click "Import".
- Vercel auto-detects Vite. Leave the build settings as they are.
- Before clicking Deploy, open "Environment Variables" and add:
  - Name: `ANTHROPIC_API_KEY`
  - Value: your key from the Anthropic console
- Click "Deploy". Wait for it to finish, then open the live URL.

That's it. The photos will be there immediately. The live feed and assistant
work because the key is set.

> If you deploy without the key, the app still runs and the curated trend photos
> still show. The live Refresh and the assistant will ask you to add the key.

---

## Run it locally first (optional)
You need Node.js 18 or newer.

```bash
npm install
cp .env.example .env.local      # then paste your key into .env.local
npm run dev
```

Open the address it prints (usually http://localhost:5173).

Note: `npm run dev` serves the app but not the `/api/claude` function, so the
live feed and assistant only work once deployed to Vercel (or run with the
Vercel CLI: `npm i -g vercel` then `vercel dev`).

---

## Notes
- Data is stored in your browser (localStorage), so it stays on the device/
  browser you use. Clearing site data clears it. The in-app Backup button
  exports a JSON copy.
- The live feed uses Anthropic web search. If Refresh reports it could not load,
  your curated set still shows; web search may need enabling on your API plan.
- To change the starter trends or their photos, edit `SEED_TRENDS` near the top
  of `src/App.jsx`.
- Costs: Vercel's hobby tier is free. Anthropic API usage is billed per use on
  your own account; the trend refresh and assistant are the only things that
  call it.

## Project layout
```
index.html            app shell
src/main.jsx          entry point
src/App.jsx           the whole app
src/storage.js        browser storage (replaces the sandbox storage)
src/index.css         Tailwind entry
api/claude.js         secure serverless proxy to Anthropic
```
