#!/usr/bin/env python3
"""
Convertit un fichier Meteo_YYYYMMDD-hhmm.html en PDF A4 paysage.

La conversion utilise Chromium via Playwright afin de conserver le rendu HTML/CSS
d'origine: couleurs, grille, SVG, ombres et typographie.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


FILE_RE = re.compile(r"^Meteo_(\d{8}-\d{4})\.html$", re.IGNORECASE)

PRINT_CSS = """
@page {
  size: A4 landscape;
  margin: 8mm;
}

html,
body {
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
"""


def find_latest_html(base_dir: Path) -> Path:
    candidates: list[tuple[str, Path]] = []
    for path in base_dir.glob("Meteo_*.html"):
        match = FILE_RE.match(path.name)
        if match:
            candidates.append((match.group(1), path))

    if not candidates:
        raise FileNotFoundError(
            f"Aucun fichier Meteo_YYYYMMDD-hhmm.html trouve dans {base_dir}"
        )

    candidates.sort(key=lambda item: item[0])
    return candidates[-1][1]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convertit un bulletin Meteo HTML en PDF A4 paysage."
    )
    parser.add_argument(
        "--input",
        type=Path,
        help=(
            "Fichier HTML source. Si omis, prend le plus recent "
            "Meteo_*.html du dossier courant."
        ),
    )
    parser.add_argument(
        "--output",
        type=Path,
        help=(
            "Chemin du PDF de sortie. Par defaut, meme nom que le HTML "
            "avec extension .pdf."
        ),
    )
    parser.add_argument(
        "--scale",
        type=float,
        default=0.9,
        help="Echelle d'impression entre 0.1 et 2.0. Defaut: 0.9.",
    )
    parser.add_argument(
        "--margin",
        default="8mm",
        help="Marge PDF sur les 4 cotes. Defaut: 8mm.",
    )
    return parser.parse_args()


def resolve_paths(args: argparse.Namespace) -> tuple[Path, Path]:
    base_dir = Path.cwd()
    input_html = args.input or find_latest_html(base_dir)
    if not input_html.is_absolute():
        input_html = (base_dir / input_html).resolve()

    if not input_html.exists():
        raise FileNotFoundError(f"Fichier introuvable: {input_html}")

    output_pdf = args.output or input_html.with_suffix(".pdf")
    if not output_pdf.is_absolute():
        output_pdf = (base_dir / output_pdf).resolve()

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    return input_html, output_pdf


def render_with_playwright(input_html: Path, output_pdf: Path, scale: float, margin: str) -> None:
    try:
        from playwright.sync_api import Error as PlaywrightError
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "Playwright n'est pas installe. Installer avec: pip install playwright "
            "puis playwright install chromium"
        ) from exc

    scale = max(0.1, min(scale, 2.0))
    html_uri = input_html.as_uri()

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            page = browser.new_page(
                viewport={"width": 1600, "height": 1200},
                device_scale_factor=1,
            )
            page.goto(html_uri, wait_until="networkidle")
            page.emulate_media(media="print")
            page.add_style_tag(content=PRINT_CSS.replace("8mm", margin))
            page.pdf(
                path=str(output_pdf),
                format="A4",
                landscape=True,
                print_background=True,
                scale=scale,
                margin={
                    "top": margin,
                    "right": margin,
                    "bottom": margin,
                    "left": margin,
                },
                prefer_css_page_size=True,
            )
            browser.close()
    except PlaywrightError as exc:
        raise RuntimeError(f"Echec du rendu Chromium/Playwright: {exc}") from exc


def main() -> int:
    args = parse_args()

    try:
        input_html, output_pdf = resolve_paths(args)
        print(f"Source : {input_html}")
        print(f"Sortie : {output_pdf}")
        print(f"Format : A4 paysage, echelle {args.scale:.2f}, marge {args.margin}")
        render_with_playwright(input_html, output_pdf, args.scale, args.margin)
    except Exception as exc:
        print(f"Erreur: {exc}", file=sys.stderr)
        return 1

    print("PDF genere avec succes.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
