#!/usr/bin/env python3
"""Verify local relative links and required static-docs invariants."""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1] / "docs"
PAGES = [ROOT / "index.html", ROOT / "en" / "index.html", ROOT / "404.html"]
HREF = re.compile(r"""(?:href|src)=["']([^"'#]+)""", re.I)
SKIP_PREFIX = ("http://", "https://", "mailto:", "data:")


def resolve(page: Path, url: str) -> Path | None:
    if url.startswith(SKIP_PREFIX):
        return None
    if url.startswith("/word-print-layout-pipeline/"):
        rel = url[len("/word-print-layout-pipeline/") :]
        return (ROOT / rel).resolve()
    if url.startswith("/"):
        return None
    return (page.parent / url).resolve()


def main() -> int:
    errors: list[str] = []
    checked = 0
    for page in PAGES:
        if not page.exists():
            errors.append(f"missing page {page}")
            continue
        text = page.read_text(encoding="utf-8")
        for raw in HREF.findall(text):
            target = resolve(page, raw)
            if target is None:
                continue
            checked += 1
            if not str(target).startswith(str(ROOT.resolve())):
                errors.append(f"{page.relative_to(ROOT)}: escaped root {raw}")
                continue
            if not target.exists():
                errors.append(f"{page.relative_to(ROOT)}: broken {raw} -> {target}")
        lower = text.lower()
        if page.name != "404.html":
            if "libreoffice" not in lower or "xps" not in lower:
                errors.append(f"{page.relative_to(ROOT)}: expected LibreOffice and XPS mentions")
            if "静态" not in text and "static" not in lower:
                errors.append(f"{page.relative_to(ROOT)}: missing static-docs disclaimer")
            if "upload" in lower and "type=\"file\"" in lower:
                errors.append(f"{page.relative_to(ROOT)}: must not include a Word upload form")

    for svg in (ROOT / "assets/img/architecture.svg", ROOT / "assets/img/wrap-problem.svg"):
        raw = svg.read_bytes()
        try:
            raw.decode("utf-8")
        except UnicodeDecodeError as exc:
            errors.append(f"{svg.name}: not UTF-8 ({exc})")
        if not raw.startswith(b"<?xml"):
            errors.append(f"{svg.name}: missing XML declaration")

    for extra in (
        ROOT / "assets/css/site.css",
        ROOT / "assets/js/site.js",
        ROOT / "assets/img/architecture.svg",
        ROOT / "assets/img/wrap-problem.svg",
        ROOT / "assets/img/favicon.svg",
        ROOT / "robots.txt",
        ROOT / "sitemap.xml",
        ROOT / ".nojekyll",
    ):
        if not extra.exists():
            errors.append(f"missing {extra.relative_to(ROOT)}")

    sitemap = (ROOT / "sitemap.xml").read_text(encoding="utf-8")
    if "unstoppablecurry.github.io/word-print-layout-pipeline" not in sitemap:
        errors.append("sitemap missing intended host path")

    workflow = Path(__file__).resolve().parents[1] / ".github" / "workflows" / "pages.yml"
    wf = workflow.read_text(encoding="utf-8")
    for needle in ("actions/configure-pages", "actions/upload-pages-artifact", "actions/deploy-pages", "path: docs"):
        if needle not in wf:
            errors.append(f"workflow missing {needle}")

    print(f"checked {checked} relative asset links")
    if errors:
        print("FAIL")
        for e in errors:
            print(" -", e)
        return 1
    print("OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
