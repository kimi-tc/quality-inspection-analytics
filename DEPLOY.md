# Deployment Notes

## Current recommended path

Use Render to deploy this project as a single Node web service.

Why:
- The app already has a React frontend and an Express backend.
- Shared data is currently stored in a local JSON file.
- Render supports Node web services and persistent disks for filesystem data.

Official docs used:
- Render web services: https://render.com/docs/web-services/
- Render persistent disks: https://render.com/docs/disks

## What is already prepared

- `server/index.ts`
  - Serves `/api/*` endpoints
  - Serves built frontend files from `dist/` in production
  - Supports `DATA_DIR` so shared data can live on a persistent disk

- `render.yaml`
  - Defines one Render web service
  - Builds the frontend
  - Starts the shared backend
  - Mounts a persistent disk at `/opt/render/project/src/data`

## Local development

```bash
npm run dev
```

This starts:
- Vite frontend
- Express shared-data backend

## Render deployment steps

1. Push this project to GitHub.
2. In Render, create a new Blueprint or Web Service from the repo.
3. If using Blueprint, Render will read `render.yaml`.
4. After deploy completes, open the Render URL.
5. Import weekly Excel files from the web UI.
6. Bosses can open the same URL and view the same shared data.

## Shortest go-live checklist

1. Create a GitHub repo.
2. Push this folder to the repo.
3. Go to Render and click `New +` -> `Blueprint`.
4. Select the GitHub repo.
5. Confirm the service created from `render.yaml`.
6. Wait for the first deploy to finish.
7. Open the generated `onrender.com` URL.
8. Import your weekly Excel file once.
9. Send the same URL to your boss.

## Files that should be committed

- `src/*`
- `server/index.ts`
- `package.json`
- `package-lock.json`
- `vite.config.ts`
- `render.yaml`
- `DEPLOY.md`
- `.env.example`
- `.gitignore`

## Files that should not be committed

- `node_modules/`
- `dist/`
- `data/shared-dataset.json`
- real `.env` files

## Important limitation right now

Shared data is stored in a JSON file on the mounted disk.

This is fine for:
- one team
- low-frequency weekly imports
- read-heavy viewing

It is not ideal long-term for:
- many concurrent editors
- audit history
- row-level permissions

If needed later, the next upgrade path is:
- move dataset storage from JSON file to Postgres
