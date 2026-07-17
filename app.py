from __future__ import annotations

import io
import json
import os
import re
import unicodedata
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass, asdict
from html import escape, unescape

from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import load_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024

SAFARI_API = "https://api.distribuidorasafari.com.br/api/produto/pesquisar"
UA = "Mozilla/5.0 (compatible; SafariFotos/1.0)"

MANUFACTURERS = {
    "NATICON": "naticonindustria.com.br",
    "ASTRA": "astra-sa.com",
    "MAX": "maxeb.com.br",
    "SFORPLAST": "sforplast.com.br",
}


@dataclass
class Result:
    row: int
    code: str
    description: str
    supplier: str
    image_url: str = ""
    product_url: str = ""
    source: str = ""
    status: str = "Não encontrado"
    confidence: int = 0
    normal_price: float | None = None
    promo_price: float | None = None


def fetch(url: str, *, timeout: int = 20) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json,text/html,*/*"})
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return response.read()


def norm(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]+", " ", text.upper()).strip()


def similarity(a: str, b: str) -> int:
    sa, sb = set(norm(a).split()), set(norm(b).split())
    if not sa or not sb:
        return 0
    return round(100 * len(sa & sb) / len(sa | sb))


def safari_search(code: str, description: str) -> Result | None:
    for keyword in (code, description):
        query = urllib.parse.urlencode({"pagina": 1, "resultados_por_pagina": 12, "keyword": keyword})
        try:
            payload = json.loads(fetch(f"{SAFARI_API}?{query}"))
        except Exception:
            continue
        candidates = payload.get("resultados") or []
        if not candidates:
            continue
        exact = [p for p in candidates if str(p.get("codigo", "")).split(".")[0] == code]
        product = exact[0] if exact else max(candidates, key=lambda p: similarity(description, p.get("nome", "")))
        image = product.get("imagem") or ""
        if image and "sem_imagem" not in image.lower():
            score = 100 if exact else similarity(description, product.get("nome", ""))
            return Result(0, code, description, "", image, f"https://distribuidorasafari.com.br/produto/{product['id']}", "Safari", "Encontrado" if score >= 80 else "Revisar", score)
    return None


def manufacturer_domain(supplier: str, description: str) -> str:
    haystack = norm(f"{supplier} {description}")
    for name, domain in MANUFACTURERS.items():
        if name in haystack:
            return domain
    return ""


def manufacturer_search(description: str, supplier: str) -> Result | None:
    domain = manufacturer_domain(supplier, description)
    query = f'"{description}"' + (f" site:{domain}" if domain else " fabricante")
    url = "https://www.bing.com/images/search?" + urllib.parse.urlencode({"q": query, "form": "HDRSC2"})
    try:
        html = fetch(url).decode("utf-8", "ignore")
        for raw in re.findall(r'class="iusc"[^>]+m="([^"]+)"', html):
            meta = json.loads(unescape(raw))
            image, page = meta.get("murl", ""), meta.get("purl", "")
            if image and (not domain or domain in page):
                return Result(0, "", description, supplier, image, page, "Fabricante", "Revisar", 65)
    except Exception:
        pass
    return None


def parse_workbook(data: bytes) -> list[Result]:
    workbook = load_workbook(io.BytesIO(data), data_only=False)
    values_workbook = load_workbook(io.BytesIO(data), data_only=True)
    sheet = workbook.active
    values_sheet = values_workbook.active
    headers = {norm(cell.value): cell.column for cell in sheet[1] if cell.value}
    code_col = headers.get("PRODUTO")
    desc_col = headers.get("DESCRICAO DO PRODUTO")
    normal_col = headers.get("VENDA")
    promo_col = headers.get("PROMO")
    if not code_col or not desc_col:
        raise ValueError("A planilha precisa das colunas Produto e Descrição do Produto.")
    supplier = ""
    rows = []
    for index in range(2, sheet.max_row + 1):
        code = sheet.cell(index, code_col).value
        description = sheet.cell(index, desc_col).value
        first = sheet.cell(index, 1).value
        if first and code is None and description is None:
            supplier = str(first)
            continue
        if code is not None and description:
            normal_price = values_sheet.cell(index, normal_col).value if normal_col else None
            promo_price = values_sheet.cell(index, promo_col).value if promo_col else None
            rows.append(Result(index, str(code).strip(), str(description).strip(), supplier,
                               normal_price=float(normal_price) if isinstance(normal_price, (int, float)) else None,
                               promo_price=float(promo_price) if isinstance(promo_price, (int, float)) else None))
    return rows


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/import")
def import_sheet():
    file = request.files.get("file")
    if not file:
        return jsonify(error="Selecione uma planilha."), 400
    try:
        rows = parse_workbook(file.read())
        return jsonify(items=[asdict(row) for row in rows], total=len(rows))
    except Exception as exc:
        return jsonify(error=str(exc)), 400


@app.post("/api/search")
def search():
    item = request.get_json(force=True)
    base = Result(int(item.get("row", 0)), str(item["code"]), item["description"], item.get("supplier", ""),
                  normal_price=item.get("normal_price"), promo_price=item.get("promo_price"))
    found = safari_search(base.code, base.description)
    if not found:
        found = manufacturer_search(base.description, base.supplier)
    if found:
        found.row, found.code, found.supplier = base.row, base.code, base.supplier
        found.normal_price, found.promo_price = base.normal_price, base.promo_price
        return jsonify(asdict(found))
    return jsonify(asdict(base))


@app.post("/api/export")
def export():
    body = request.get_json(force=True)
    items = body.get("items", [])
    original = bytes.fromhex(body["workbook_hex"])
    workbook = load_workbook(io.BytesIO(original))
    sheet = workbook.active
    start = sheet.max_column + 1
    for offset, title in enumerate(("URL da Imagem", "Fonte da Imagem", "Página do Produto", "Status da Foto")):
        sheet.cell(1, start + offset, title)
    for item in items:
        row = int(item["row"])
        sheet.cell(row, start, item.get("image_url", ""))
        sheet.cell(row, start + 1, item.get("source", ""))
        sheet.cell(row, start + 2, item.get("product_url", ""))
        sheet.cell(row, start + 3, item.get("status", ""))
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="planilha_com_fotos.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/zip")
def images_zip():
    items = request.get_json(force=True).get("items", [])
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for item in items:
            url, code = item.get("image_url"), item.get("code")
            if not url:
                continue
            try:
                content = fetch(url)
                ext = os.path.splitext(urllib.parse.urlparse(url).path)[1].lower()
                if ext not in {".jpg", ".jpeg", ".png", ".webp"}:
                    ext = ".jpg"
                archive.writestr(f"{code}{ext}", content)
            except Exception:
                continue
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="fotos_produtos.zip", mimetype="application/zip")


def money(value: object) -> tuple[str, str]:
    try:
        integer, cents = f"{float(value):.2f}".split(".")
        return integer, cents
    except (TypeError, ValueError):
        return "0", "00"


def section_name(supplier: str) -> str:
    cleaned = re.sub(r"^\d+\s+", "", supplier or "Ofertas Safari").strip()
    aliases = {"NATICON": "Elétrica Naticon", "ASTRA": "Sanitários Astra", "MAX METALURGICA": "Ferramentas de Jardim e Obra Max", "ELITE": "Iluminação Elite", "SFORPLAST": "Pregos Sforplast", "FARROUPILHA": "Fixação Farroupilha"}
    normalized = norm(cleaned)
    for key, label in aliases.items():
        if key in normalized:
            return label
    if "MAC FER" in normalized:
        return "Ferramentas Thompson"
    return cleaned.title()


def build_tabloid(items: list[dict], start_date: str, end_date: str, density: str) -> str:
    per_page = {"compacto": 30, "equilibrado": 24, "destaque": 16}.get(density, 24)
    columns = {"compacto": 5, "equilibrado": 4, "destaque": 4}.get(density, 4)
    pages = [items[i:i + per_page] for i in range(0, len(items), per_page)] or [[]]
    page_html = []
    for page_number, page_items in enumerate(pages, 1):
        cards, current_section = [], None
        for item in page_items:
            section = section_name(item.get("supplier", ""))
            if section != current_section:
                cards.append(f'<div class="section-title"><span>{escape(section)}</span></div>')
                current_section = section
            normal_i, normal_c = money(item.get("normal_price"))
            promo_i, promo_c = money(item.get("promo_price"))
            image = escape(item.get("image_url") or "")
            photo = f'<img class="product-photo" src="{image}" alt="Produto">' if image else '<div class="no-photo">SEM FOTO</div>'
            cards.append(f'<article class="card">{photo}<div class="card-body"><div class="cod">Cód. {escape(str(item.get("code", "")))}</div><div class="desc">{escape(item.get("description", ""))}</div></div><div class="price-row"><span class="de">De R$ {normal_i},{normal_c}</span><span class="por">R$ <b>{promo_i}</b><sup>,{promo_c}</sup></span></div></article>')
        page_html.append(f'<section class="a4-page" style="--cols:{columns}"><header><div class="hero-title">SUPER<small>OFERTAS</small></div><div class="brand"><b>Distribuidora <em>Safari</em></b><span>VÁLIDA DE {escape(start_date)} ATÉ {escape(end_date)}</span></div></header><main class="catalog">{"".join(cards)}</main><footer><span>Após a validade, os preços voltarão ao normal. Imagens meramente ilustrativas.</span><b>(11) 2911-9888 · @distribuidorasafari</b><i>{page_number}/{len(pages)}</i></footer></section>')
    styles = '''@page{size:A4 portrait;margin:0}*{box-sizing:border-box}body{margin:0;background:#dfe8e1;font-family:Arial,sans-serif;color:#183024}.a4-page{width:210mm;height:297mm;margin:10mm auto;background:#fff;display:flex;flex-direction:column;overflow:hidden;page-break-after:always}header{height:29mm;background:linear-gradient(135deg,#0d5c2e,#1a8a44);padding:6mm 8mm;display:flex;justify-content:space-between;align-items:center;color:#fff}.hero-title{font-size:28pt;font-weight:1000;line-height:.75;color:#f4e4a8}.hero-title small{display:block;font-size:11pt;letter-spacing:4px;margin-top:4mm}.brand{text-align:right;display:flex;flex-direction:column;gap:3mm}.brand b{font-size:16pt}.brand em{color:#d4af37;font-style:normal}.brand span{background:#d4af37;color:#0d5c2e;padding:2mm 4mm;border-radius:99px;font-weight:800;font-size:8pt}.catalog{padding:4mm 6mm;display:grid;grid-template-columns:repeat(var(--cols),1fr);gap:2.3mm;align-content:start;flex:1;overflow:hidden}.section-title{grid-column:1/-1;border-bottom:1.2mm solid #0d5c2e;height:7mm;display:flex;align-items:end}.section-title span{background:#0d5c2e;color:#fff;padding:1.2mm 4mm;font-size:7.5pt;font-weight:900;text-transform:uppercase;border-radius:2mm 2mm 0 0}.card{border:.3mm solid #d9e3dc;border-radius:2mm;padding:2mm;display:flex;flex-direction:column;min-height:0;overflow:hidden;break-inside:avoid}.product-photo,.no-photo{width:100%;height:22mm;object-fit:contain;display:block}.no-photo{display:grid;place-items:center;background:#f1f3f1;color:#9aa39c;font-size:7pt;font-weight:bold}.card-body{flex:1}.cod{font-size:6pt;color:#6b7c72;font-weight:bold}.desc{font-size:6.8pt;font-weight:800;line-height:1.15;margin-top:1mm;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}.price-row{border-top:.3mm dashed #d9e3dc;margin-top:1.5mm;padding-top:1mm}.de{display:block;font-size:6pt;color:#6b7c72;text-decoration:line-through}.por{color:#d81f2b;font-size:12pt;font-weight:1000}.por sup{font-size:7pt}footer{height:12mm;border-top:1mm solid #0d5c2e;margin:0 6mm;padding:2.5mm 0;display:flex;align-items:center;gap:4mm;font-size:6pt;color:#6b7c72}footer b{color:#183024;margin-left:auto}footer i{font-style:normal;font-weight:bold}@media print{body{background:#fff}.a4-page{margin:0}}@media screen{.a4-page{box-shadow:0 8px 32px #0002}}'''
    return f'<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Tabloide Safari</title><style>{styles}</style></head><body>{"".join(page_html)}</body></html>'


@app.post("/api/tabloid")
def tabloid():
    body = request.get_json(force=True)
    html = build_tabloid(body.get("items", []), body.get("start_date", ""), body.get("end_date", ""), body.get("density", "equilibrado"))
    return send_file(io.BytesIO(html.encode("utf-8")), as_attachment=True, download_name="tabloide_safari.html", mimetype="text/html; charset=utf-8")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
