# docs/ — GitHub Pages (static only)

This directory is the bilingual architecture site published at
<https://unstoppablecurry.github.io/word-print-layout-pipeline/>.

It is **HTML/CSS/SVG documentation**. It does not run LibreOffice, Flask, or
the Windows XPS node. Do not treat it as a live converter.

| Path | Role |
| --- | --- |
| `index.html` | Chinese (primary) |
| `en/index.html` | English |
| `assets/` | Shared CSS, JS, SVG diagrams |
| `.nojekyll` | Skip Jekyll processing on Pages |

Enable **Settings → Pages → Source: GitHub Actions** so `.github/workflows/pages.yml` can deploy `docs/`.
