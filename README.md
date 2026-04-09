# Footer-Batch

Mass footer changer for PDFs and Word docs — upload, add footer, download instantly.

## Live Deployment

This app is deployed on **Vercel** and accessible globally at your Vercel URL.

## How to Deploy (make it live for everyone)

### Step 1: Push to GitHub

```bash
git add .
git commit -m "ready for deployment"
git push origin main
```

### Step 2: Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) and sign in with your GitHub account.
2. Click **"Add New Project"** → Import your **Footer-Batch** repository.
3. Vercel will auto-detect the settings from `vercel.json`. **No configuration needed.**
4. Click **"Deploy"** — your app will be live in ~60 seconds.
5. You'll get a URL like `https://footer-batch.vercel.app` — share it with anyone on the globe!

### How It Works

| Layer      | Technology            | Location                    |
|------------|-----------------------|-----------------------------|
| Frontend   | React + Vite          | `frontend/` → built to `frontend/dist` |
| Backend    | Vercel Serverless Fn  | `api/process-document.js`   |
| Hosting    | Vercel Edge Network   | 30+ global regions          |

- Vercel builds the frontend from `frontend/` and serves the static files.
- The `/api/process-document` endpoint runs as a serverless function.
- Files are processed in-memory (no storage needed) — scales automatically.

## Local Development

```bash
# Terminal 1 — Backend
cd backend
npm install
npm run dev

# Terminal 2 — Frontend
cd frontend
npm install
npm run dev
```

The frontend dev server runs at `http://localhost:5173` and proxies API calls to the backend at `http://localhost:5000`.
