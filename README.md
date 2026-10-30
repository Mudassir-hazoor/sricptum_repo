# Scriptorium

Scriptorium is a small, static preview of a calmer document-tools workspace.
The current site is a landing page only: file uploads, document conversions,
and server-side processing are not implemented.

## View the site

After GitHub Pages deployment is enabled, the public site will be available at:

<https://mudassir-hazoor.github.io/sricptum_repo/>

The GitHub Actions workflow deploys the site whenever changes are pushed to
`main`. In the repository settings, set **Pages → Build and deployment → Source**
to **GitHub Actions**.

## Run locally

Requires Node.js 22 or later.

```powershell
npm install
npm run dev
```

Open the local address printed by Astro, usually <http://localhost:4321>.

To check the production build:

```powershell
npm run build
```

The static site is generated in `dist`.

## Privacy

This preview is a static site. It does not accept, upload, or process files.
Do not use it to submit sensitive documents.
