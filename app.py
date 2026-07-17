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
from html import unescape

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
    sheet = workbook.active
    headers = {norm(cell.value): cell.column for cell in sheet[1] if cell.value}
    code_col = headers.get("PRODUTO")
    desc_col = headers.get("DESCRICAO DO PRODUTO")
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
            rows.append(Result(index, str(code).strip(), str(description).strip(), supplier))
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
    base = Result(int(item.get("row", 0)), str(item["code"]), item["description"], item.get("supplier", ""))
    found = safari_search(base.code, base.description)
    if not found:
        found = manufacturer_search(base.description, base.supplier)
    if found:
        found.row, found.code, found.supplier = base.row, base.code, base.supplier
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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
