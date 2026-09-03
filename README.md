# Skolar — Option A

Deployable React/Vite version of Skolar. The app no longer depends on a hosted artifact storage API; it uses browser localStorage so it can run as a normal static website.

## Run locally

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

The production files are created in `dist/`.

## Publish with GitHub + Vercel/Netlify

1. Create a GitHub repository named `skolar`.
2. Upload all files in this folder, keeping the `src` folder and config files.
3. Import the repository into Vercel or Netlify.
4. Build command: `npm run build`.
5. Output directory: `dist`.

## Important data note

This version stores data in each browser's localStorage. It is suitable for a demo/static deployment, but it is **not** a shared school database. For real multi-user accounts and shared records across phones/computers, connect the app to Supabase (authentication + database) in the next step.
