# Sideloading Guide

This project supports three sideloading lanes:

1. **Local desktop dev (fastest)** via `localhost`
2. **Local shared folder catalog** (Windows testing flow)
3. **Staging manifest** that points to **GitHub Pages**

## Important Model

A shared folder catalog distributes the **manifest only**.

The add-in web app (task pane HTML/JS/CSS) must be hosted at the HTTPS URLs in the manifest (`SourceLocation`, icon URLs).

- For one-machine local dev, `https://localhost:3000` is fine.
- For testing from other machines, use `manifests/manifest.staging.xml` that points to GitHub Pages.

## Lane 1: Local Desktop Dev

```bash
npm run start:desktop
```

When done:

```bash
npm run stop
```

## Lane 2: Local Shared Folder Catalog (Windows)

### 1) Create local share

```bash
npm run sideload:share:setup
```

Default local folder: `%USERPROFILE%\OfficeAddinCatalog`  
Default share name: `OfficeAddinCatalog`

### 2) Trust catalog in Office

```bash
npm run sideload:share:trust
```

Restart Excel after trust registration.

### 3) Publish staging manifest into share

```bash
npm run manifest:staging
npm run sideload:share:publish
```

In Excel: **Home > Add-ins > More Add-ins > Shared Folder**, then add `manifest.staging.xml`.

### 4) Cleanup

```bash
npm run sideload:share:cleanup
```

## Lane 3: Staging on GitHub Pages

GitHub Pages deploys from `main` via `.github/workflows/pages.yml`.

Staging manifest target base URL:

- `https://sbroenne.github.io/office-coding-agent`

Generated file:

- `manifests/manifest.staging.xml`

Regenerate anytime:

```bash
npm run manifest:staging
```

## Troubleshooting

- **Add-in not visible in Shared Folder**
  - Ensure Excel was restarted after `sideload:share:trust`.
  - Confirm the file exists in `%USERPROFILE%\OfficeAddinCatalog`.
- **Task pane doesn’t load on another machine**
  - The manifest probably points to `localhost`. Use `manifest.staging.xml`.
- **Share setup fails**
  - `New-SmbShare` may require elevated PowerShell (Run as Administrator).
