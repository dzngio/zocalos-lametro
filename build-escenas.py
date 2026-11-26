"""Genera inicio.html y fin.html desde escena.src.html (logo y QR embebidos: las escenas quedan autocontenidas).
Uso: python build-escenas.py
"""
import base64, os, re

here = os.path.dirname(os.path.abspath(__file__))
src = open(os.path.join(here, "escena.src.html"), encoding="utf-8").read()

logo_svg = open(os.path.join(here, "assets", "LA METRO LOGO SOLO.svg"), encoding="utf-8").read()
shapes = "".join(s.replace(' class="cls-1"', "") for s in re.findall(r"<(?:polygon|path)[^>]*/>", logo_svg))
logo = '<svg viewBox="0 0 287.57 62.28" aria-label="la metro">' + shapes + "</svg>"

qr = "data:image/svg+xml;base64," + base64.b64encode(open(os.path.join(here, "assets", "qr-inscripcion.svg"), "rb").read()).decode()

for escena, title, out in [("inicio", "Ya comenzamos", "inicio.html"), ("fin", "Gracias por acompañarnos", "fin.html")]:
    html = src.replace("__ESCENA__", escena).replace("__TITLE__", title).replace("__LOGO__", logo).replace("__QR__", qr)
    open(os.path.join(here, out), "w", encoding="utf-8").write(html)
    print("ok", out, len(html))
