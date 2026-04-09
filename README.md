# Footer-Batch
Mass footer changer for PDFs and Word docs - upload, add footer, download instantly.

## Deployment

- Deploy the repository root on Vercel.
- Vercel builds the `frontend` workspace and serves the app from `frontend/dist`.
- The production upload endpoint is the same origin `/api/process-document` route.
- Keep `backend/` for local development only.
