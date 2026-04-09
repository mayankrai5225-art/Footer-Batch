# Footer-Batch
Mass footer changer for PDFs and Word docs - upload, add footer, download instantly.

## Deployment

- Local development uses `http://localhost:5000` for the backend.
- Production uses `VITE_API_BASE_URL` when set, otherwise it falls back to `/_/backend`.
- If you deploy the backend on another host, set `VITE_API_BASE_URL` to that public URL before building the frontend.
