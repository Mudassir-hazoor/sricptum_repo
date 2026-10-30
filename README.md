# Scriptorium

Scriptorium is a privacy-first document-conversion website starter. The frontend is a static Astro site that can be hosted on GitHub Pages. The optional FastAPI backend currently provides a health endpoint and is a scaffold for future server-side converters.

> **Current scope:** The site includes example conversion catalogue pages, but file conversion is not implemented yet. Setting `PUBLIC_API_URL` does not add conversion endpoints; the backend currently only exposes `/` and `/health`.

## Requirements

- Node.js LTS and npm for the frontend
- Docker Desktop for the optional backend
- A GitHub account and a public repository for GitHub Pages hosting

## Run locally on Windows

Open PowerShell in the project folder and run:

```powershell
Set-Location .\frontend
npm.cmd install
npm.cmd run dev
```

Open the local URL printed by Astro, usually <http://localhost:4321>. Keep the terminal open while using the site. Use `Ctrl+C` to stop the development server.

To build and test the frontend:

```powershell
Set-Location .\frontend
npm.cmd run build
npm.cmd test
```

The static output is written to `frontend\dist`.

## Run the optional backend

From the project root:

```powershell
docker compose up --build
```

The backend health check is at <http://localhost:8000/health>. Stop it with `Ctrl+C`; run `docker compose down` to stop and remove its container.

## Deploy the frontend to GitHub Pages

1. Create a **public** GitHub repository named `scriptorium`.
2. Upload the contents of this project folder to the repository root, including the `.github` directory.
3. Edit `frontend\astro.config.mjs`:
   - Set `site` to `https://YOUR-USERNAME.github.io`.
   - Keep `base: '/scriptorium'` if the repository is named `scriptorium`.
   - If using a custom domain, set `site` to that domain and remove the `base` setting.
4. Push the changes to the repository's `main` branch.
5. In GitHub, open **Settings → Pages** and select **GitHub Actions** as the build and deployment source.
6. Open the **Actions** tab and check that the “Deploy to GitHub Pages” workflow succeeds.
7. For a repository named `scriptorium`, the site URL will be `https://YOUR-USERNAME.github.io/scriptorium/`.

The workflow in `.github\workflows\deploy-pages.yml` builds the `frontend` directory and publishes its static output.

## Optional API URL

The frontend reads the `PUBLIC_API_URL` environment variable at build time. The deployment workflow supplies the GitHub Actions repository variable of the same name. The current frontend/backend scaffold does not yet implement document conversion APIs, so leave this variable unset unless you are developing the backend integration.

## Project layout

```text
frontend/                 Astro static site
backend/                  FastAPI backend scaffold
.github/workflows/        CI and GitHub Pages deployment
docker-compose.yml        Local backend container
```

## Privacy and security

Do not upload sensitive documents to a server unless you have deployed and verified the intended processing and file-retention behavior. The current backend is not a document converter and does not implement file upload, conversion, or deletion workflows.
