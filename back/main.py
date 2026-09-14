from __future__ import annotations

import io
import re
import zipfile
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
import pdfplumber
import pypdfium2 as pdfium
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process
from PIL import Image, ImageOps
import pytesseract

from ocr_engine import HAS_OCR, OCRWord, extract as ocr_extract, linguagem_disponivel

BASE_DIR = Path(__file__).resolve().parent
CADASTRO_DIR = BASE_DIR / "cadastro"
PRODUTOS_XLSX = CADASTRO_DIR / "produtos_totvs.xlsx"
SERVICOS_XLSX = CADASTRO_DIR / "servicos_totvs.xlsx"

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

CNPJ_RE = re.compile(r"(?<!\d)\d{2}[.\s]?\d{3}[.\s]?\d{3}[/\s]?\d{4}[-\s]?\d{2}(?!\d)")
DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
NCM_RE = re.compile(r"(?<!\d)(\d{7,8})(?!\d)")
MONEY_RE = re.compile(r"(?<![\d/])(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2,4}(?!\d)")
UNIDADES = {"UN", "UND", "UNID", "PC", "PÇ", "KG", "G", "L", "LT", "M", "M2", "M3", "CX", "FD", "SC", "HR", "H", "JG", "PAR", "TON", "RL", "ML", "SV"}
SERVICE_SYNONYMS = {
    "LIXO": ["RESIDUO", "DESTINACAO RESIDUO"],
    "RESIDUO": ["DESTINACAO RESIDUO"],
    "RESIDUOS": ["DESTINACAO RESIDUO"],
    "RESIDUO SOLIDO": ["DESTINACAO RESIDUO"],
    "SUCATA": ["DESTINACAO RESIDUO"],
    "DESCARTE": ["DESTINACAO RESIDUO"],
    "COLETA DE LIXO": ["DESTINACAO RESIDUO"],
    "MECANICO": ["MECANICA"],
    "MECANICA": ["MECANICA"],
    "MECANICOS": ["MECANICA"],
    "ELETRICO": ["ELETRICA"],
    "ELETRICA": ["ELETRICA"],
    "ELETRICOS": ["ELETRICA"],
    "TORNEARIA": ["TORNEARIA"],
    "TORNEAMENTO": ["TORNEARIA"],
    "USINAGEM": ["USINAGEM"],
    "CALIBRAR": ["CALIBRACAO"],
    "CALIBRACAO": ["CALIBRACAO"],
    "TRANSPORTAR": ["TRANSPORTE"],
    "TRANSPORTE": ["TRANSPORTE"],
    "ALUGUEL": ["LOCACAO"],
    "ALUGUEL DE": ["LOCACAO"],
    "LOCAR": ["LOCACAO"],
}


def normalizar_texto(v: Any) -> str:
    s = str(v or "")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().upper()


def normalizar_descricao(v: Any) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", normalizar_texto(v))).strip()


def parse_num(v: str | None) -> float | None:
    if not v:
        return None
    s = str(v).strip().replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def format_num(v: float | None, decimals: int = 3) -> str:
    if v is None:
        return "-"
    s = f"{v:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s.rstrip("0").rstrip(",") if "," in s else s


def clean_company(v: str) -> str:
    """
    Limpa uma razão social sem cortar palavras legítimas.

    Regra importante: NÃO procurar "ME"/"SA" no meio da frase, porque isso
    pode truncar empresas como "COMERCIO" ou "MATERIAL". O nome só é
    considerado encerrado por sufixo societário quando o sufixo aparece no
    final do candidato.
    """
    s = re.sub(r"\s+", " ", str(v or "")).strip(" :-|_[]()")
    s = re.split(r"(?i)\b(?:E-?MAIL|EMAIL)\b", s)[0]
    s = re.split(r"(?i)\b(?:CNPJ|CPF)\b", s)[0]
    s = s.strip(" :-|_[]()")

    # PDF nativo às vezes cola espaços: "LORENOKOCHLTDA".
    # Só aceitamos o sufixo se ele estiver no FINAL da string.
    suffix = re.search(r"(?:LTDA|EIRELI|EPP|ME|S\s*\.?\s*A\.?|S/A)\s*$", s, re.I)
    if suffix:
        return s

    return s


def rows_from_words(words: list[OCRWord]) -> list[dict]:
    """Agrupa palavras por posição vertical, não apenas pelo line_num do Tesseract."""
    words = sorted(words, key=lambda w: (w.cy, w.left))
    rows: list[dict] = []
    for w in words:
        if not rows or abs(w.cy - rows[-1]["cy"]) > max(13, w.height * 0.9):
            rows.append({"cy": w.cy, "top": w.top, "bottom": w.top + w.height, "words": [w]})
        else:
            r = rows[-1]
            r["words"].append(w)
            r["cy"] = sum(x.cy for x in r["words"]) / len(r["words"])
            r["top"] = min(r["top"], w.top)
            r["bottom"] = max(r["bottom"], w.top + w.height)
    out = []
    for r in rows:
        ws = sorted(r["words"], key=lambda x: x.left)
        out.append({"cy": r["cy"], "top": r["top"], "bottom": r["bottom"], "text": " ".join(x.text for x in ws), "words": ws})
    return out


def process_pdf_pages(pdf_bytes: bytes) -> list[dict]:
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            natives = [p.extract_text() or "" for p in pdf.pages]
    except Exception as e:
        return [{"text": "", "lines": [], "native": [], "ocr": False, "conf": 0, "words": [], "rows": [], "image": None, "error": f"PDF inválido: {e}"}]

    doc = pdfium.PdfDocument(pdf_bytes)
    out = []
    for i, native in enumerate(natives):
        text = native.strip()
        ocr_used = False
        conf = 0.0
        words: list[OCRWord] = []
        image = None
        error = ""
        try:
            # OCR só é acionado quando o PDF realmente parece escaneado/híbrido.
            # NFS-e, boletos e CT-e digitais têm texto nativo suficiente e não devem
            # gastar tempo renderizando a página inteira para OCR.
            un = normalizar_texto(text)
            structured_digital = any(k in un for k in [
                "DANFSE", "VALOR TOTAL DA NFS-E", "PRESTADOR / FORNECEDOR",
                "RECIBO DO PAGADOR", "COMPROVANTE DE ENTREGA", "NOSSO NUMERO",
                "DACTE", "CONHECIMENTO DE TRANSPORTE"
            ])
            possible_nfe = any(k in un for k in ["DANFE", "NF-E", "RECEBEMOS DE", "DESTINATARIO / REMETENTE"])
            needs_ocr = HAS_OCR and (
                len(text) < 250
                or (possible_nfe and not any(k in un for k in ["DADOS DOS PRODUTOS", "DADOS DO PRODUTO", "DADOS DO PRODUTOS / SERVICOS", "DADOS DO PRODUTO / SERVICO"]) and not structured_digital)
            )
            if needs_ocr:
                image = doc[i].render(scale=300 / 72).to_pil()
                r = ocr_extract(image, detalhar=True)
                conf = r.confidence
                words = r.words
                ocr_used = True
                if len(text) < 300:
                    text = r.text
                else:
                    existing = {normalizar_texto(x) for x in text.splitlines()}
                    merged = text.splitlines()
                    for line in r.lines:
                        if normalizar_texto(line) not in existing:
                            merged.append(line)
                    text = "\n".join(merged)
        except Exception as e:
            error = f"Falha no OCR: {e}"
        rows = rows_from_words(words)
        out.append({"text": text, "lines": text.splitlines(), "native": native.splitlines(), "ocr": ocr_used, "conf": conf, "words": words, "rows": rows, "image": image, "error": error})
    return out


def classify(text: str) -> str:
    u = normalizar_texto(text)
    if any(x in u for x in ["DACTE", "CT-E", "DOCUMENTO AUXILIAR DO CONHECIMENTO DE TRANSPORTE"]):
        return "CTE"
    nfse = sum(x in u for x in ["DANFSE", "NFS-E", "PRESTADOR / FORNECEDOR", "PRESTADOR DO SERVICO", "VALOR TOTAL DA NFS-E"])
    boleto = sum(x in u for x in ["RECIBO DO PAGADOR", "COMPROVANTE DE ENTREGA", "NOSSO NUMERO", "VALOR DO DOCUMENTO", "LINHA DIGITAVEL"])
    nfe = sum(x in u for x in ["DANFE", "NATUREZA DA OPERACAO", "DADOS DOS PRODUTOS", "NF-E", "RECEBEMOS DE"])
    if nfse >= 1 and nfse >= max(boleto, nfe):
        return "NFSE"
    if boleto >= 2 and boleto > nfe:
        return "BOLETO"
    if nfe >= 1:
        return "NFE"
    return "OUTRO"


def extract_nfe_emitter(lines: list[str], image: Image.Image | None) -> str:
    """Extrai o emitente da NF-e priorizando âncoras legais e o bloco do DANFE."""
    full = "\n".join(lines)
    patterns = [
        # Layouts tradicionais
        r"RECEBEMOS\s+DE\s+(.+?)\s+(?:OS\s+PRODUTOS(?:\s+E/?OU\s+SERVI[CÇ]OS)?|OS\s+SERVI[CÇ]OS)\s+CONSTANTES",
        r"RECEBEMOS\s+DE\s+(.+?)\s+OS\s+PRODUTOS\s+E?/\?OU\s+SERVI[CÇ]OS\s+CONSTANTES",
        r"RECEBEMOS\s+DE\s+(.+?)(?:\s+OS\s+PRODUTOS|\s+OS\s+SERVI[CÇ]OS)",
    ]
    for pat in patterns:
        m = re.search(pat, full, re.I | re.S)
        if m:
            cand = clean_company(m.group(1))
            if len(normalizar_descricao(cand)) >= 4:
                return cand

    # Muitos DANFEs colocam a razão social na mesma linha de "DANFE".
    # Ex.: "COOPERATIVA CERCAR DANFE ..." / "LORENO KOCH DANFE ..."
    for i, line in enumerate(lines[:40]):
        if "DANFE" not in normalizar_texto(line):
            continue
        before = re.split(r"(?i)\bDANFE\b", line, maxsplit=1)[0].strip(" :-|_[]()")
        before = clean_company(before)
        nb = normalizar_descricao(before)
        if len(nb) >= 4 and not any(k in nb for k in ["DOCUMENTO AUXILIAR", "NOTA FISCAL", "CONTROLE DO FISCO"]):
            return before

        # Se a razão social estiver na linha imediatamente anterior ao DANFE.
        for j in (i - 1, i + 1):
            if 0 <= j < len(lines):
                cand = clean_company(lines[j])
                nc = normalizar_descricao(cand)
                if len(nc) >= 4 and not any(k in nc for k in ["DOCUMENTO AUXILIAR", "NOTA FISCAL", "CHAVE DE ACESSO", "CONTROLE DO FISCO"]):
                    return cand

    if image:
        try:
            top = image.crop((0, 0, image.width, int(image.height * .28)))
            txt = pytesseract.image_to_string(top, lang=linguagem_disponivel(), config="--oem 3 --psm 6")
            for line in txt.splitlines():
                m = re.search(r"RECEBEMOS\s+DE\s+(.+?)(?:\s+OS\s+PRODUTOS|\s+OS\s+SERVI[CÇ]OS)", line, re.I)
                if m:
                    cand = clean_company(m.group(1))
                    if len(normalizar_descricao(cand)) >= 4:
                        return cand
                if "DANFE" in normalizar_texto(line):
                    cand = clean_company(re.split(r"(?i)\bDANFE\b", line, 1)[0])
                    if len(normalizar_descricao(cand)) >= 4:
                        return cand
        except Exception:
            pass
    return ""

def extract_nfe_number(lines: list[str]) -> str:
    """Extrai o número completo da NF-e, inclusive formatos 000.335.546."""
    for line in lines[:55]:
        u = normalizar_texto(line)
        # Captura "Nº 60", "Nº 000.335.546", "N° 000.003.254" etc.
        m = re.search(r"(?:N[º°]|N[O0]\b)\s*[:.-]?\s*(\d{1,3}(?:[.\s]\d{3})+|\d{1,9})\b", u)
        if m:
            valor = re.sub(r"\D", "", m.group(1))
            if valor:
                return valor.lstrip("0") or "0"
        # Fallback quando o rótulo aparece como NF-e Nº ...
        m = re.search(r"\bNF[- ]?E?\s*(?:N[º°:]?\s*)?(\d{1,3}(?:[.]\d{3})+|\d{1,9})\b", u)
        if m and "CNPJ" not in u:
            valor = re.sub(r"\D", "", m.group(1))
            if valor:
                return valor.lstrip("0") or "0"
    return "SemNumero"

def _numeric_candidates(text: str) -> list[str]:
    vals = []
    for x in re.findall(r"(?<!\d)(?:\d{1,3}(?:\.\d{3})+|\d+)(?:[.,]\d{1,4})?(?!\d)", text):
        if parse_num(x) is not None:
            vals.append(x)
    return vals


def _find_nfe_header_centers(page: dict) -> dict:
    """Encontra o centro das colunas da tabela DANFE usando as palavras do cabeçalho.
    A leitura da quantidade nunca depende apenas da ordem dos tokens do OCR.
    """
    centers: dict[str, float] = {}
    rows = page.get("rows", [])
    if not rows:
        return centers

    labels = {
        "quantidade": ("QUANT", "QTDE", "QTD", "QUANTIDADE"),
        "unitario": ("V UNIT", "VUNIT", "UNIT"),
        "total": ("V TOTAL", "VTOTAL", "TOTAL"),
        "unidade": ("UNID", "UN", "UND"),
    }

    # Primeiro procura o cabeçalho inteiro. Os PDFs podem separar "V." e "UNIT."
    # em duas palavras, portanto avaliamos também janelas de 3 palavras.
    for row in rows:
        ws = sorted(row["words"], key=lambda w: w.left)
        txt = normalizar_texto(row["text"])
        if not any(k in txt for k in ["QUANT", "QTDE", "QTD"]):
            continue
        for i, w in enumerate(ws):
            u = normalizar_texto(w.text).replace(".", "")
            if any(k in u for k in labels["quantidade"]):
                centers["quantidade"] = w.cx
            elif any(k in u for k in labels["unidade"]):
                centers.setdefault("unidade", w.cx)
            # Ex.: "V." + "UNIT." ou "V" + "TOTAL".
            nxt = " ".join(normalizar_texto(x.text).replace(".", "") for x in ws[i:i+2])
            if "V UNIT" in nxt or "VUNIT" in nxt or "UNIT" == u:
                centers["unitario"] = (ws[i].cx + ws[min(i+1, len(ws)-1)].cx) / 2
            if "V TOTAL" in nxt or "VTOTAL" in nxt:
                centers["total"] = (ws[i].cx + ws[min(i+1, len(ws)-1)].cx) / 2
        if len(centers) >= 3:
            break

    # Segunda tentativa: procurar os rótulos globalmente caso o OCR tenha separado
    # o cabeçalho em linhas diferentes.
    if "quantidade" not in centers:
        for row in rows:
            for w in row["words"]:
                u = normalizar_texto(w.text).replace(".", "")
                if any(k in u for k in labels["quantidade"]):
                    centers["quantidade"] = w.cx
                    break
            if "quantidade" in centers:
                break

    # "UNID" pode aparecer tanto no cabeçalho quanto na própria linha do item.
    # Só usamos o valor do cabeçalho se ele veio da mesma região da quantidade.
    return centers


def _ocr_column_number(image: Image.Image, x_center: float, y0: int, y1: int, width_ratio: float = 0.045) -> str:
    if image is None or x_center is None:
        return ""
    x0 = max(0, int(x_center - image.width * width_ratio / 2))
    x1 = min(image.width, int(x_center + image.width * width_ratio / 2))
    crop = image.crop((x0, max(0, y0 - 12), x1, min(image.height, y1 + 12))).convert("L")
    crop = ImageOps.autocontrast(crop)
    crop = crop.resize((max(220, crop.width * 8), max(80, crop.height * 8)))
    candidatos = []
    for psm in (7, 6, 11, 13):
        try:
            raw = pytesseract.image_to_string(
                crop, lang=linguagem_disponivel(),
                config=f"--oem 3 --psm {psm} -c tessedit_char_whitelist=0123456789,."
            ).strip()
        except Exception:
            continue
        for token in _numeric_candidates(raw):
            n = parse_num(token)
            if n is None or n <= 0:
                continue
            score = 100 + (5 if re.search(r"[,.]\d", token) else 0)
            candidatos.append((score, token))
    return sorted(candidatos, reverse=True)[0][1] if candidatos else ""


def _numeric_tokens_between(words, x_left: float, x_right: float):
    vals = []
    for w in sorted(words, key=lambda x: x.left):
        if w.cx < x_left or w.cx > x_right:
            continue
        for token in _numeric_candidates(w.text):
            n = parse_num(token)
            if n is not None and n > 0:
                vals.append((w, token, n))
    return vals


def _quantity_from_row_words(page: dict, row: dict, headers: dict, unit_center: float | None = None) -> str:
    """Lê QUANTIDADE isolando a célula entre as linhas/centros das colunas.
    Isso evita capturar V.TOTAL, que é o erro mais comum nos scans.
    """
    image = page.get("image")
    if image is None:
        return ""

    qx = headers.get("quantidade")
    ux = headers.get("unitario")
    tx = headers.get("total")

    # Quando conseguimos os centros do cabeçalho, delimitamos a célula por metades.
    if qx is not None:
        left = (unit_center + qx) / 2 if unit_center is not None else max(0, qx - image.width * 0.045)
        right_candidates = [x for x in (ux, tx) if x is not None and x > qx]
        right = (qx + min(right_candidates)) / 2 if right_candidates else qx + image.width * 0.035
        vals = _numeric_tokens_between(row.get("words", []), left, right)
        valid = []
        for w, token, n in vals:
            if 0 < n <= 100000000:
                score = w.confidence + 20
                if re.search(r"[,.]\d{1,4}$", token):
                    score += 12
                # Quantidade deve ficar depois da UN e antes de V.UNIT.
                if unit_center is not None and w.cx <= unit_center:
                    score -= 25
                valid.append((score, token))
        if valid:
            return sorted(valid, reverse=True)[0][1]

        # Se a célula estiver vazia no OCR, uma leitura dirigida é muito mais segura
        # do que OCR da linha inteira.
        if ux is not None:
            return _ocr_column_number(image, qx, row["top"], row["bottom"], 0.04)

    # Fallback para scans sem cabeçalho reconhecido: a quantidade fica entre a
    # unidade e o primeiro bloco numérico grande à direita.
    if unit_center is not None:
        vals = []
        for w, token, n in _numeric_tokens_between(row.get("words", []), unit_center + image.width * 0.015, image.width * 0.86):
            # Descartamos explicitamente valores tipicamente pertencentes a desconto,
            # valor unitário ou total apenas pela posição/ordem. O primeiro candidato
            # decimal razoável depois da UN é a quantidade.
            if n <= 0 or n > 100000000:
                continue
            score = w.confidence
            if re.search(r"[,.]\d{1,4}$", token):
                score += 8
            score -= max(0, (w.cx - unit_center) / image.width) * 20
            vals.append((w.cx, score, token, n))
        if vals:
            vals.sort(key=lambda x: x[0])
            # Pula descontos 0,00/0 e pega o primeiro valor positivo significativo.
            for _, _, token, n in vals:
                if n > 0.0001:
                    return token

    return ""

def extract_nfe_items(page: dict) -> list[dict]:
    """Extrai itens da DANFE. Para scans, a quantidade é lida pela coluna QUANTIDADE."""
    if not page["ocr"]:
        return extract_nfe_native(page["lines"])
    out = []
    seen = set()
    rows = page.get("rows", [])
    headers = _find_nfe_header_centers(page)
    in_table = False
    for row in rows:
        u = normalizar_texto(row["text"])
        if "DADOS DOS PRODUTOS" in u or ("COD" in u and "PRODUTO" in u and "QUANT" in u):
            in_table = True
            continue
        if in_table and any(k in u for k in ["DADOS ADICIONAIS", "INFORMACOES COMPLEMENTARES"]):
            in_table = False
            continue
        if not in_table:
            # Alguns scans não reconhecem o cabeçalho inteiro; aceita a linha se parecer um item.
            if not re.match(r"^\d{4,}\s", u):
                continue
        ws = sorted(row["words"], key=lambda x: x.left)
        if not ws:
            continue
        code_i = None
        code = ""
        for i, w in enumerate(ws[:8]):
            digits = re.sub(r"\D", "", w.text)
            if len(digits) >= 4:
                code_i, code = i, digits
                break
        if code_i is None:
            continue
        # Procura unidade depois do código. NCM/CST/CFOP podem estar colados e não são usados como âncora.
        unit_i = None
        unit = ""
        for i, w in enumerate(ws[code_i + 1:], start=code_i + 1):
            t = re.sub(r"[^A-Za-zÀ-ÿ]", "", w.text).upper()
            if t in UNIDADES:
                unit_i, unit = i, t
                break
        if unit_i is None:
            continue
        parts = []
        for w in ws[code_i + 1:unit_i]:
            t = w.text.strip("[]()|_")
            if not t:
                continue
            digits = re.sub(r"\D", "", t)
            # descarta NCM/CST/CFOP fundidos ou isolados
            if digits and not re.search(r"[A-Za-zÀ-ÿ]", t) and len(digits) >= 4:
                continue
            parts.append(t)
        desc = re.sub(r"\s+", " ", " ".join(parts)).strip(" -")
        # Alguns scans fundem CST/CSOSN ao final da descrição.
        desc = re.sub(r"(?:\s+|^)(?:0\d\d|\d{3,4})$", "", desc).strip(" -")
        if len(desc) < 3:
            continue
        unit_center = ws[unit_i].cx if unit_i is not None else None
        quantidade = _quantity_from_row_words(page, row, headers, unit_center)
        # O NCM é um sinal auxiliar muito útil para o cadastro de produtos.
        # Pegamos somente sequências de 8 dígitos entre o código e a UN,
        # evitando CST/CFOP (normalmente 2-3/4 dígitos).
        ncm = ""
        for w in ws[code_i + 1:unit_i]:
            digits = re.sub(r"\D", "", w.text)
            if len(digits) == 8:
                ncm = digits
                break
        qnum = parse_num(quantidade)
        if qnum is None or qnum <= 0:
            quantidade = "1"
            qnum = 1.0
        key = (code, normalizar_descricao(desc))
        if key in seen:
            continue
        seen.add(key)
        out.append({"codigo_fornecedor": code, "descricao": desc, "ncm": ncm, "un": unit, "quantidade": format_num(qnum, 3)})
    return out


def extract_nfe_native(lines: list[str]) -> list[dict]:
    """Parser robusto para DANFE digitais: extrai código, descrição, NCM, unidade e quantidade."""
    out = []
    in_table = False
    item_re = re.compile(
        r"^(?P<codigo>[^\s]+)\s+(?P<desc>.*?)\s+"
        r"(?P<ncm>\d{8}|\d{4}\.\d{2}\.\d{2})\s+"
        r"(?P<trib>[A-Za-z0-9/.\-]+)\s+(?P<cfop>\d{4})\s+"
        r"(?P<un>UN|UND|UNID|PC|PÇ|KG|G|L|LT|M|M2|M3|CX|FD|SC|HR|H|JG|PAR|TON|RL|ML|SV)\s+"
        r"(?P<qtd>\d+(?:[.,]\d+)?)\b",
        re.I,
    )
    for i, line in enumerate(lines):
        u = normalizar_texto(line)
        if "DADOS DOS PRODUTOS" in u or "DADOS DO PRODUTO" in u:
            in_table = True
            continue
        if in_table and any(k in u for k in ["DADOS ADICIONAIS", "INFORMACOES COMPLEMENTARES"]):
            break
        if not in_table:
            continue
        m = item_re.match(line.strip())
        if not m:
            continue
        qtd = parse_num(m.group("qtd"))
        if qtd is None or qtd <= 0:
            qtd = 1.0
        out.append({
            "codigo_fornecedor": m.group("codigo"),
            "descricao": m.group("desc").strip(" -"),
            "ncm": re.sub(r"\D", "", m.group("ncm")),
            "un": m.group("un").upper(),
            "quantidade": format_num(qtd, 3),
        })
    return out

def _compactar_rotulo(v: str) -> str:
    """Normaliza rótulos do DANFSe removendo espaços artificiais do PDF."""
    return re.sub(r"[^A-Z0-9]", "", normalizar_texto(v))


def _eh_descricao_generica_nfse(desc: str) -> bool:
    n = normalizar_descricao(desc)
    return n in {
        "PRESTACAO DE SERVICO",
        "SERVICO PRESTADO",
        "SERVICOS PRESTADOS",
        "SERVICO",
        "SERVICOS",
        "PRESTACAO",
    }


def _parece_razao_social(v: str) -> bool:
    u = normalizar_texto(v).strip()
    return bool(
        u
        and re.search(r"(?:LTDA|S\.?A\.?|S/A|EIRELI|ME|EPP)$", u, re.I)
        and len(u) >= 4
    )


def _linha_descritiva_servico(s: str) -> str:
    """Retorna uma linha limpa de descrição da área SERVIÇO PRESTADO ou ''."""
    if not s or len(normalizar_descricao(s)) < 5:
        return ""
    raw=re.sub(r"\s+"," ",s).strip()
    u=_compactar_rotulo(raw)
    if any(k in u for k in ["SERVICOPRESTADO","DESCRICAODOSERVICO","CODIGODETRIBUTACAO","CODIGODANBS","LOCALDAPRESTACAO","TRIBUTACAOMUNICIPAL","TRIBUTACAOFEDERAL","TRIBUTACAOIBSCBS","VALORTOTAL","VALORDAOPERACAO","BCISSQN","ALIQUOTAAPLICADA"]):
        return ""
    # Remove códigos/NBS e município que o extrator cola ao final da linha.
    # Ex.: "... eletrônica, 310103 / 120015000 MARECHAL..." -> "... eletrônica,".
    raw2 = re.split(r"\s+\d{3,}(?:[./]\d+)*\s*/", raw, maxsplit=1)[0].strip()
    if re.fullmatch(r"[\d./\-\s]+",raw2):
        return ""
    if re.search(r"(?i)\b(?:R\$|CEP|CNPJ|CPF|E-MAIL|EMAIL|MUNICIPIO|SIGLA UF)\b",raw2):
        return ""
    # Linhas iniciadas por código tributário/localização não são descrição.
    if re.match(r"^\s*(?:\d{2}(?:[./]\d{2}){1,3}|\d{3,}(?:[./]\d+)*)(?:\s*/|\s+-|\s+\d)",raw2):
        return ""
    letters=len(re.sub(r"[^A-Za-zÀ-ÿ]","",raw2))
    digits=len(re.sub(r"[^0-9]","",raw2))
    if letters < 5 or (digits > letters*2 and digits >= 8):
        return ""
    return raw2.strip(" |-")


def _extrair_servico_prestado_completo(lines: list[str]) -> tuple[str,str]:
    compact=[_compactar_rotulo(x) for x in lines]
    inicio=next((i for i,u in enumerate(compact) if "SERVICOPRESTADO" in u),None)
    if inicio is None: return "",""
    fim=len(lines)
    for i in range(inicio+1,len(lines)):
        if any(k in compact[i] for k in ["TRIBUTACAOMUNICIPAL","VALORTOTALDANFSE","VALORDAOPERACAOSERVICO"]):
            fim=i; break
    desc_idx=next((i for i in range(inicio+1,fim) if "DESCRICAODOSERVICO" in compact[i]),None)
    before=[]; after=[]
    for i in range(inicio+1,fim):
        if desc_idx is not None and i>desc_idx:
            cleaned=_linha_descritiva_servico(lines[i])
            if cleaned: after.append(cleaned)
        elif desc_idx is None:
            cleaned=_linha_descritiva_servico(lines[i])
            if cleaned: before.append(cleaned)
        else:
            cleaned=_linha_descritiva_servico(lines[i])
            if cleaned: before.append(cleaned)
    def dedupe(parts):
        out=[]
        for x in parts:
            nx=normalizar_descricao(x)
            if nx and all(nx!=normalizar_descricao(y) for y in out): out.append(x)
        return out
    before,after=dedupe(before),dedupe(after)
    after = [x for x in after if not _eh_descricao_generica_nfse(x)]
    completo=" ".join(before+after).strip()
    if not completo and before:
        completo=" ".join(before).strip()
    return completo," ".join(after).strip()

def _extrair_ordem_servico(lines: list[str]) -> dict|None:
    norm_full=normalizar_texto("\n".join(lines))
    if "PRESTACAO DE SERVICOS" not in norm_full or "SERVICOS ALOCADOS" not in norm_full: return None
    numero="SemNumero"
    m=re.search(r"PRESTACAO\s+DE\s+SERVICOS\s+N[º°oO0]?\s*:?\s*(\d{3,12})",norm_full)
    if m: numero=str(int(m.group(1)))
    empresa=""
    for line in lines[:15]:
        s=re.sub(r"\s+"," ",line).strip()
        m=re.search(r"^(.*?)\s+Emiss[aã]o\s*:",s,re.I)
        if m and len(m.group(1))>=5: empresa=m.group(1).strip(" :-"); break
    inicio=next((i for i,x in enumerate(lines) if "SERVICOSALOCADOS" in _compactar_rotulo(x)),None)
    descricoes=[]; qtd_total=0.0
    if inicio is not None:
        for raw in lines[inicio+1:]:
            s=re.sub(r"\s+"," ",raw).strip()
            cu=_compactar_rotulo(s)
            if "TOTALDOSSERVICOS" in cu or "TOTAISDAOS" in cu: break
            mrow=re.match(r"^\d{2}/\d{2}/\d{4}\s+\d+\s+(.*?)\s+(\d+(?:[.,]\d+)?)\s+[\d.]+,\d{2}\s+[\d.]+,\d{2}\s+[\d.]+,\d{2}\s+[\d.]+,\d{2}\s*$",s,re.I)
            if mrow:
                descricoes.append(mrow.group(1).strip()); qtd_total += parse_num(mrow.group(2)) or 1.0
    valor="-"
    for raw in lines:
        m=re.search(r"(?i)VALOR\s+TOTAL\s+DA\s+OS\s*:\s*R?\$?\s*([\d.]+,\d{2,4})",raw)
        if m: valor=m.group(1); break
    if valor=="-":
        for raw in lines:
            m=re.search(r"(?i)TOTAL\s+DOS\s+SERVI[CÇ]OS\.*\s*:?\s*\d+\s+([\d.]+,\d{2,4})",raw)
            if m: valor=m.group(1); break
    desc="; ".join(dict.fromkeys(descricoes)) or "Prestacao de servico"
    return {"numero":numero,"empresa":empresa or "Empresa não identificada","descricao":desc,"descricao_match":desc,"contexto_servico":desc,"quantidade":qtd_total if qtd_total>0 else 1.0,"valor_total":valor,"tipo_documento":"ORDEM_SERVICO"}


def extract_nfse(lines: list[str]) -> dict:
    os_doc = _extrair_ordem_servico(lines)
    if os_doc:
        return os_doc
    norm = [normalizar_texto(x) for x in lines]
    compact = [_compactar_rotulo(x) for x in lines]

    numero = "SemNumero"
    for i, u in enumerate(compact):
        if "NUMERODANFSE" in u:
            for x in lines[i:i+5]:
                # Formato comum: "14 02/09/2026 ..."
                m = re.search(r"(?<![/\d])(\d{1,8})\s+(?:\d{2}/\d{2}/\d{4}|\d{2}/\d{4})", x)
                if m and int(m.group(1)) <= 99999999:
                    numero = str(int(m.group(1)))
                    break
                # Formato nacional padronizado: "0000.00000000004 09/2026 ..."
                m = re.search(r"(?<![0-9A-Za-z])(\d{2,}\.\d{2,})(?=\s+\d{2}/(?:\d{2}/)?\d{4})", x)
                if m:
                    digits = re.sub(r"\D", "", m.group(1))
                    numero = str(int(digits)) if digits else "SemNumero"
                    break
            if numero == "SemNumero":
                for x in lines[i+1:i+5]:
                    token = re.match(r"\s*([0-9.]{2,})\s+", x)
                    if token and re.search(r"\d", token.group(1)):
                        digits = re.sub(r"\D", "", token.group(1))
                        if digits:
                            numero = str(int(digits))
                            break
            break

    empresa = ""
    # Procuramos explicitamente o cabeçalho "Nome / Nome Empresarial"
    # imediatamente dentro do bloco PRESTADOR/FORNECEDOR.
    prestador_idx = next((i for i,u in enumerate(compact) if "PRESTADORFORNECEDOR" in u), None)
    if prestador_idx is not None:
        for i in range(prestador_idx, min(len(lines), prestador_idx + 12)):
            if "NOMENOMEEMPRESARIAL" in compact[i]:
                for x in lines[i+1:i+5]:
                    raw = re.sub(r"\s+", " ", str(x or "")).strip()
                    # Em DANFSe nacional, o nome do prestador pode vir colado
                    # ao município/UF e ao CEP, por exemplo:
                    # "LORENOKOCHLTDA MarechalCândidoRondon/PR 41.14609/85.969-899".
                    # Primeiro extraímos explicitamente a razão social pelo sufixo
                    # societário, evitando que o restante da linha mate a detecção.
                    mraz = re.search(
                        r"(?i)([A-ZÀ-Ý0-9&.'\- ]+?(?:LTDA|EIRELI|EPP|S\s*\.?A\s*\.?|S/A|ME))(?=\s|$)",
                        raw
                    )
                    candidato = mraz.group(1).strip() if mraz else raw
                    c = clean_company(candidato)
                    cu = normalizar_texto(c)
                    if (
                        len(cu) >= 4
                        and not DATE_RE.search(c)
                        and "EMAIL" not in cu
                        and "MUNICIPIO" not in cu
                        and _parece_razao_social(c)
                    ):
                        empresa = c
                        break
                break

    if not empresa:
        # Fallback para layouts em que os rótulos são separados.
        for i, u in enumerate(norm):
            if "NOME / NOME EMPRES" in u:
                for x in lines[i+1:i+4]:
                    c = clean_company(x)
                    if (
                        len(normalizar_texto(c)) >= 4
                        and "EMAIL" not in normalizar_texto(c)
                        and re.search(r"\b(?:LTDA|S\.?A\.?|S/A|EIRELI|ME|EPP)\b", normalizar_texto(c))
                    ):
                        empresa = c
                        break
            if empresa:
                break

    desc = ""
    contexto = ""
    desc_idx = None
    for i,u in enumerate(compact):
        if "DESCRICAODOSERVICO" in u:
            desc_idx=i
            vals=[]
            for x in lines[i+1:i+8]:
                xu=_compactar_rotulo(x)
                if not x.strip(): continue
                if any(marker in xu for marker in ["TRIBUTACAOMUNICIPAL","TRIBUTACAOFEDERAL","TRIBUTACAOIBSCBS","VALORTOTALDANFSE","VALORDAOPERACAOSERVICO"]): break
                vals.append(x.strip())
                if len(" ".join(vals))>=250: break
            desc=re.sub(r"\s+"," "," ".join(vals)).strip()
            break

    servico_completo, desc_explicita = _extrair_servico_prestado_completo(lines)
    if servico_completo:
        # A planilha recebe o serviço completo, não apenas o resumo do campo
        # "Descrição do Serviço". Isso também recupera notas que escrevem
        # apenas "Prestação de serviço" nesse campo.
        desc = servico_completo

    if not desc:
        full="\n".join(lines)
        m=re.search(r"Descri(?:ç|c)ã?o\s*do\s*Servi(?:ç|c)o\s*[:\-]?\s*(.+?)(?=\n(?:TRIBUTA|VALOR TOTAL|VALOR DA OPERA)|$)",full,re.I|re.S)
        if m: desc=re.sub(r"\s+"," ",m.group(1)).strip()
    if not desc: desc="Prestacao de servico"
    desc=re.sub(r"(?i)^servi[cç]o(?:s)?de", "serviço de ", desc)
    desc=re.sub(r"\s+"," ",desc).strip()
    contexto=servico_completo or desc
    descricao_para_match=contexto

    valor = "-"
    for i, u in enumerate(compact):
        if "VALORDAOPERACAOSERVICO" in u or "VALORDOSERVICO" in u:
            m = re.search(r"R\$\s*([\d.]+,\d{2,4})", " ".join(lines[i:i+7]))
            if m:
                valor = m.group(1)
                break
    if valor == "-":
        for i, u in enumerate(compact):
            if "VALORTOTALDANFSE" in u:
                m = re.search(r"R\$\s*([\d.]+,\d{2,4})", " ".join(lines[i:i+10]))
                if m:
                    valor = m.group(1)
                    break

    full = normalizar_texto(" ".join(lines))
    quantidade = 1.0
    hour_patterns = [
        r"(?<!\d)(\d+(?:[.,]\d+)?)\s*(?:HORAS|HORA|HRS|HR)(?![A-Z])",
        r"QUANTIDADE\s*[:=]?\s*(\d+(?:[.,]\d+)?)",
        r"(\d+(?:[.,]\d+)?)\s*(?:H|HH)\b",
    ]
    for pat in hour_patterns:
        m = re.search(pat, full)
        if m:
            quantidade = parse_num(m.group(1)) or 1.0
            break

    return {
        "numero": numero,
        "empresa": empresa or "Empresa não identificada",
        "descricao": desc,
        "descricao_match": descricao_para_match,
        "contexto_servico": contexto,
        "quantidade": quantidade,
        "valor_total": valor,
    }


def carregar_cadastro(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Cadastro fixo não encontrado: {path.name}")
    raw = pd.read_excel(path, header=0, dtype=object)
    raw = raw.dropna(how="all").reset_index(drop=True)
    if raw.empty:
        raise ValueError(f"Cadastro vazio: {path.name}")
    return raw


def escolher_colunas_cadastro(df: pd.DataFrame) -> tuple[str, str, str | None]:
    keys = {c: normalizar_descricao(c) for c in df.columns}
    cod = next((c for c, k in keys.items() if k == "CODIGO" or k.endswith(" CODIGO")), None)
    if not cod:
        cod = next((c for c, k in keys.items() if "CODIGO" in k and "BARRAS" not in k), None)
    desc_short = next((c for c, k in keys.items() if k == "DESCRICAO"), None)
    desc_cols = [c for c, k in keys.items() if k.startswith("DESCRICAO")]
    desc = desc_short or (desc_cols[0] if desc_cols else None)
    desc_long = desc_cols[1] if len(desc_cols) > 1 else None
    if not cod or not desc:
        raise ValueError(f"Cadastro {Path(df.__class__.__name__).name if False else ''}: colunas necessárias não encontradas. Encontradas: {', '.join(map(str, df.columns))}")
    return cod, desc, desc_long


def cadastro_preparado(path: Path) -> dict:
    df = carregar_cadastro(path)
    cod_col, desc_col, desc2_col = escolher_colunas_cadastro(df)
    rows = []
    for _, r in df.iterrows():
        code = str(r.get(cod_col, "") or "").strip()
        d1 = str(r.get(desc_col, "") or "").strip()
        d2 = str(r.get(desc2_col, "") or "").strip() if desc2_col else ""
        if len(normalizar_descricao(d2)) < 5:
            d2 = ""
        ncm = str(r.get("Pos.IPI/NCM", "") or "")
        if not code or not d1:
            continue
        rows.append({
            "codigo": code,
            "descricao": d1,
            "descricao_extra": d2,
            "texto_match": normalizar_descricao(d1),
            "texto_match_extra": normalizar_descricao(d2),
            "ncm": re.sub(r"\D", "", ncm),
        })
    return {"rows": rows}


def _tokens_informativos(texto: str) -> list[str]:
    stops = {
        "DE","DA","DO","DAS","DOS","PARA","COM","SEM","EM","NO","NA",
        "NOS","NAS","E","A","O","AS","OS","UM","UMA","UN","UND",
        "UNID","PC","PCES","PÇ","SV","SERVICO","SERVICOS","SERVIÇO","SERVIÇOS",
    }
    tokens = re.findall(r"[A-Z0-9]+", normalizar_texto(texto))
    return [t for t in tokens if t not in stops and len(t) >= 2]


def _score_descricao(a: str, b: str) -> float:
    na, nb = normalizar_descricao(a), normalizar_descricao(b)
    if not na or not nb:
        return 0.0
    ta, tb = _tokens_informativos(na), _tokens_informativos(nb)
    sa, sb = set(ta), set(tb)
    common = sa & sb
    coverage = len(common) / max(1, len(sa))
    precision = len(common) / max(1, len(sb))
    f1 = 2 * coverage * precision / max(0.001, coverage + precision)
    ratio = fuzz.ratio(na, nb)
    sort = fuzz.token_sort_ratio(na, nb)
    score = f1 * 48 + sort * 0.28 + ratio * 0.24
    if na == nb:
        return 100.0
    if len(ta) >= 3 and len(tb) >= 3 and (na in nb or nb in na):
        score = max(score, min(97.0, 72 + f1 * 25))
    numeric_a = set(re.findall(r"\d+(?:[.,]\d+)?", na))
    numeric_b = set(re.findall(r"\d+(?:[.,]\d+)?", nb))
    if numeric_a and numeric_b:
        num_cov = len(numeric_a & numeric_b) / len(numeric_a)
        score += min(12, num_cov * 12)
    return round(min(100.0, score), 1)


def _candidatos_cadastro(desc: str, cadastro_rows: list[dict], ncm: str = "", service: bool = False):
    if not cadastro_rows:
        return []
    choices = []
    meta = []
    for global_i, r in enumerate(cadastro_rows):
        for field in ["texto_match", "texto_match_extra"]:
            if r.get(field):
                choices.append(r[field])
                meta.append(global_i)
    if not choices:
        return []
    q = normalizar_descricao(desc)

    # Exato é decisivo e elimina ambiguidades artificiais.
    exact_indices = [i for i, c in enumerate(choices) if c == q]
    if exact_indices:
        ranked = []
        seen = set()
        for i in exact_indices:
            gi = meta[i]
            if gi not in seen:
                ranked.append((100.0, gi)); seen.add(gi)
        return ranked

    candidate_indices = set()
    limits = 60 if not service else len(choices)
    for scorer in (fuzz.token_set_ratio, fuzz.token_sort_ratio, fuzz.WRatio, fuzz.ratio):
        for _, _, j in process.extract(q, choices, scorer=scorer, limit=min(limits, len(choices))):
            candidate_indices.add(meta[j])

    ranked = []
    for global_i in candidate_indices:
        r = cadastro_rows[global_i]
        scores = [_score_descricao(desc, r.get("descricao", ""))]
        if r.get("descricao_extra"):
            scores.append(_score_descricao(desc, r["descricao_extra"]))
        score = max(scores)

        qt = _tokens_informativos(desc)
        cand_tokens = set(_tokens_informativos(r.get("descricao", "")) + _tokens_informativos(r.get("descricao_extra", "")))
        common = set(qt) & cand_tokens
        if common:
            score += min(10.0, len(common) * 3.0)

        qnums = set(re.findall(r"\d+(?:[.,]\d+)?", normalizar_texto(desc)))
        cnums = set(re.findall(r"\d+(?:[.,]\d+)?", normalizar_texto(r.get("descricao", ""))))
        if qnums and cnums:
            score += min(8.0, (len(qnums & cnums) / len(qnums)) * 8.0)

        if (not service) and ncm and r.get("ncm") == ncm:
            score = min(100.0, score + 6.0)
        ranked.append((round(min(100.0, score), 1), global_i))
    return sorted(ranked, key=lambda x: (-x[0], x[1]))


# Mapa semântico específico para os 14 serviços do TOTVS.
# Cada entrada contém termos fortes, termos de apoio e exclusões para evitar
# que palavras genéricas como "serviço" criem falsos empates.
SERVICE_SEMANTIC_MAP = {
    "700186": {
        "name": "SERVICO DE CALIBRACAO",
        "strong": ["CALIBRACAO", "CALIBRAR", "AFERICAO", "AFERIR", "VERIFICACAO METROLOGICA", "METROLOGIA"],
        "support": ["INSTRUMENTO", "BALANCA", "MEDICAO", "MEDICAO TECNICA"],
        "negative": ["LOCAÇÃO"],
    },
    "715237": {
        "name": "LOCACAO DE ANDAIMES",
        "strong": ["ANDAIME", "ANDAIMES", "ESCORAMENTO"],
        "support": ["LOCACAO", "ALUGUEL"],
        "negative": ["MARTELETE", "GERADOR", "CACAMBA", "PLATAFORMA"],
    },
    "717703": {
        "name": "LOCACAO DE MARTELETE ROMPEDOR",
        "strong": ["MARTELETE", "ROMPEDOR", "MARTELO ROMPEDOR"],
        "support": ["LOCACAO", "ALUGUEL", "DEMOLICAO"],
        "negative": ["GERADOR", "PLATAFORMA", "ANDAIME"],
    },
    "730009": {
        "name": "SERVICO LOCACAO EQUIPAMENTOS CONSTRUCAO CIVIL",
        "strong": ["EQUIPAMENTO CONSTRUCAO", "EQUIPAMENTOS CONSTRUCAO", "MAQUINA CONSTRUCAO", "EQUIPAMENTO PARA CONSTRUCAO"],
        "support": ["LOCACAO", "ALUGUEL", "CONSTRUCAO CIVIL", "EQUIPAMENTO", "MAQUINA"],
        "negative": ["GERADOR", "MARTELETE", "ANDAIME", "PLATAFORMA", "CACAMBA", "DATADORA"],
    },
    "730170": {
        "name": "SERVICO DE CONCRETAGEM",
        "strong": ["CONCRETAGEM", "CONCRETAR", "CONCRETO"],
        "support": ["MATERIAL CONCRETO", "BOMBEAMENTO CONCRETO", "BOMBA DE CONCRETO"],
        "negative": [],
    },
    "750027": {
        "name": "SERVICO ASSISTENCIA TECNICA",
        "strong": [
            "ASSISTENCIA TECNICA", "MANUTENCAO", "MANUTENCAO MECANICA", "MANUTENCAO ELETRICA",
            "REPARO", "CONSERTO", "FURAR", "FURACAO", "MANDRILHAR", "MANDRILAGEM",
            "RETIFICAR", "RETIFICA", "USINAGEM", "USINAR", "FABRICAR", "FABRICACAO",
            "SERVICO MECANICO", "SERVICO ELETRICO", "ELETRICA", "MECANICA", "BUCHA", "EIXO", "CHAPA", "INSTALACAO E MONTAGEM", "INSTALACAO DE EQUIPAMENTOS", "PASTEURIZACAO",
        ],
        "support": ["OFICINA", "TECNICA", "TECNICO", "EQUIPAMENTO", "MOTOR", "MAQUINA", "PECA"],
        "negative": ["HORA HOMEM", "HORA-HOMEM", "MAO DE OBRA", "CALIBRACAO", "GERADOR", "ANDAIME"],
    },
    "750685": {
        "name": "SERVICO ASSISTENCIA TECNICA (HORA HOMEM)",
        "strong": ["HORA HOMEM", "HORA-HOMEM", "MAO DE OBRA", "HOMEM HORA", "HH", "HORAS TECNICAS", "HORA TECNICA"],
        "support": ["ASSISTENCIA TECNICA", "MANUTENCAO", "MECANICA", "ELETRICA", "TECNICO", "TECNICA", "HORAS", "HORA"],
        "negative": [],
    },
    "760000": {
        "name": "SERVICO HOSPEDAGEM",
        "strong": ["HOSPEDAGEM", "HOTEL", "POUSADA", "ALOJAMENTO"],
        "support": ["DIARIA", "DIARIAS", "ESTADIA", "HOSPEDE"],
        "negative": [],
    },
    "771206": {
        "name": "SERVICO DESTINACAO RESIDUO",
        "strong": ["DESTINACAO RESIDUO", "DESTINACAO DE RESIDUO", "RESIDUO", "RESIDUOS", "LIXO", "DESCARTE", "SUCATA", "COLETA DE RESIDUO", "DESTINACAO"],
        "support": ["AMBIENTAL", "REJEITO", "MATERIAL DESCARTADO"],
        "negative": ["CACAMBA", "LOCACAO DE CACAMBA"],
    },
    "780007": {
        "name": "SERVICO TRANSPORTE",
        "strong": ["TRANSPORTE", "TRANSPORTAR", "FRETE", "CARGA", "ENTREGA"],
        "support": ["CAMINHAO", "VEICULO", "LOGISTICA", "COLETA"],
        "negative": ["RESIDUO", "LIXO", "CACAMBA", "PLATAFORMA"],
    },
    "780117": {
        "name": "LOCACAO DE PLATAFORMA ELEVATORIA",
        "strong": ["PLATAFORMA ELEVATORIA", "PLATAFORMA ELEVATORIA", "PLATAFORMA ARTICULADA", "PLATAFORMA TESOURA"],
        "support": ["PLATAFORMA", "LOCACAO", "ALUGUEL", "ELEVATORIA"],
        "negative": ["GERADOR", "MARTELETE", "ANDAIME", "CACAMBA"],
    },
    "830132": {
        "name": "LOCACAO DE CACAMBAS LIXO/ENTULHO",
        "strong": ["CACAMBA", "CACAMBAS", "CAÇAMBA", "CAÇAMBAS"],
        "support": ["ENTULHO", "LIXO", "LOCACAO", "ALUGUEL", "RESIDUO"],
        "negative": ["DESTINACAO RESIDUO", "TRANSPORTE DE RESIDUO"],
    },
    "900076": {
        "name": "SERVICO DE LOCACAO DATADORA",
        "strong": ["DATADORA", "CODIFICADORA", "CODIFICADOR", "MARCADOR DE DATA", "MARCADOR DE LOTE", "IMPRESSORA DE DATA", "IMPRESSORA DE LOTE"],
        "support": ["LOCACAO", "ALUGUEL", "VALIDADE", "LOTE"],
        "negative": [],
    },
    "900136": {
        "name": "LOCACAO DE GERADOR",
        "strong": ["GERADOR", "GRUPO GERADOR", "GERADOR DE ENERGIA"],
        "support": ["LOCACAO", "ALUGUEL", "ENERGIA", "KVA", "KVA"],
        "negative": [],
    },
}


def _contains_phrase(texto: str, termo: str) -> bool:
    n = normalizar_descricao(texto)
    t = normalizar_descricao(termo)
    return bool(t and t in n)


def _service_semantic_score(desc: str, row: dict) -> float:
    code = str(row.get("codigo", "")).strip()
    profile = SERVICE_SEMANTIC_MAP.get(code)
    base = _score_descricao(desc, row.get("descricao_extra") or row.get("descricao", ""))
    if not profile:
        return base

    nd = normalizar_descricao(desc)
    strong_hits = [t for t in profile["strong"] if _contains_phrase(nd, t)]
    support_hits = [t for t in profile["support"] if _contains_phrase(nd, t)]
    negative_hits = [t for t in profile["negative"] if _contains_phrase(nd, t)]

    score = base * 0.35
    # Termos fortes dominam o resultado; um único termo muito discriminativo já
    # pode ser suficiente quando o cadastro tem poucos serviços.
    if strong_hits:
        # Nosso cadastro tem poucos serviços. Um termo realmente distintivo
        # (ex.: GERADOR, CACAMBA, DATADORA, CALIBRACAO) vale muito mais que
        # a mera semelhança textual.
        score += min(84.0, 68.0 + 12.0 * (len(strong_hits) - 1))
    score += min(22.0, 7.0 * len(support_hits))
    score -= min(35.0, 12.0 * len(negative_hits))

    # Regras específicas para pares muito parecidos.
    if code == "750685":
        if any(k in nd for k in ["HORA HOMEM", "HOMEM HORA", "MAO DE OBRA", "HORAS TECNICAS", "HORA TECNICA"]):
            score = max(score, 94.0)
        elif any(k in nd for k in ["HORA", "HORAS"]) and any(k in nd for k in ["ASSISTENCIA", "MANUTENCAO", "MECANICA", "ELETRICA", "TECNICA", "TECNICO"]):
            score = max(score, 92.0)
        elif "HORA" in nd or "HORAS" in nd:
            score = max(score, 78.0)
    if code == "750027" and any(k in nd for k in ["ASSISTENCIA TECNICA", "MANUTENCAO", "REPARO", "CONSERTO"]):
        if any(k in nd for k in ["HORA HOMEM", "MAO DE OBRA", "HOMEM HORA"]) or (any(k in nd for k in ["HORA", "HORAS"]) and any(k in nd for k in ["ASSISTENCIA", "MANUTENCAO", "TECNICA", "TECNICO"])):
            score = min(score, 60.0)
        else:
            score = max(score, 86.0)

    # O serviço de caçamba precisa ganhar de "destinação de resíduo" quando há
    # a palavra caçamba explicitamente.
    if code == "830132" and any(k in nd for k in ["CACAMBA", "CAÇAMBA"]):
        score = max(score, 94.0)
    if code == "771206" and any(k in nd for k in ["DESTINACAO RESIDUO", "RESIDUO", "RESIDUOS"]):
        if not any(k in nd for k in ["CACAMBA", "CAÇAMBA"]):
            score = max(score, 92.0)

    return round(min(100.0, max(0.0, score)), 1)


def _service_interpretation(desc: str, cadastro_rows: list[dict]):
    """Classifica a NFS-e pelos 14 serviços conhecidos e retorna ranking + evidências."""
    ranked = []
    nd = normalizar_descricao(desc)
    for i, row in enumerate(cadastro_rows):
        score = _service_semantic_score(desc, row)
        profile = SERVICE_SEMANTIC_MAP.get(str(row.get("codigo", "")).strip(), {})
        strong = sum(1 for t in profile.get("strong", []) if _contains_phrase(nd, t))
        support = sum(1 for t in profile.get("support", []) if _contains_phrase(nd, t))
        negative = sum(1 for t in profile.get("negative", []) if _contains_phrase(nd, t))
        evidence = strong * 3 + support - negative * 2
        ranked.append((score, evidence, i))
    ranked.sort(key=lambda x: (-x[0], -x[1], x[2]))
    return ranked


def _formatar_outras_pecas(ranked, rows, best_index: int, best_score: float, forcar: bool = False) -> str:
    alternativas=[]
    if forcar:
        rel=max(12.0,min(25.0,best_score*0.30))
        for pos,(score,idx) in enumerate(ranked):
            if idx==best_index: continue
            if score>=max(30.0,best_score-rel) or pos<=2:
                codigo=str(rows[idx].get("codigo","")).strip()
                if codigo: alternativas.append(codigo)
            if len(alternativas)>=5: break
    else:
        rel=max(8.0,min(15.0,best_score*0.12))
        for score,idx in ranked:
            if idx==best_index or score<65.0: continue
            if score>=best_score-rel or score>=82.0:
                codigo=str(rows[idx].get("codigo","")).strip()
                if codigo: alternativas.append(codigo)
            if len(alternativas)>=5: break
    if not alternativas: return ""
    return "Outras peças viáveis: " + ", ".join(dict.fromkeys(alternativas))

def correlacionar(items: list[dict], cadastro: dict, centro_custo: str, tipo: str) -> list[dict]:
    rows = cadastro["rows"]
    out = []
    service = tipo == "NFS-E"

    for item in items:
        desc_exibicao = str(item.get("descricao_nota", "") or "").strip()
        desc_match = str(item.get("descricao_match") or desc_exibicao).strip()
        ocr = float(item.get("confianca_ocr") or 0)

        if service:
            ranked = _service_interpretation(desc_match, rows)
            best = ranked[0] if ranked else None
            second_score = ranked[1][0] if len(ranked) > 1 else 0.0
            best_score = best[0] if best else 0.0
            best_row = rows[best[2]] if best else None
            nd = normalizar_descricao(desc_match)
            generic = _eh_descricao_generica_nfse(desc_exibicao) and not (
                len(_tokens_informativos(desc_match)) > len(_tokens_informativos(desc_exibicao))
            )
            profile = SERVICE_SEMANTIC_MAP.get(str(best_row.get("codigo", "")).strip(), {}) if best_row else {}
            strong_hit_count = sum(1 for t in profile.get("strong", []) if _contains_phrase(nd, t))
            gap = best_score - second_score

            # Classificações semânticas fortes podem ser aceitas mesmo quando o
            # segundo candidato é relativamente próximo. Em caso de serviço genérico
            # ou conflito real entre duas famílias, continuamos em revisão.
            accepted = bool(
                best_row
                and not generic
                and (
                    strong_hit_count >= 1 and best_score >= 78.0
                    or best_score >= 88.0 and gap >= 6.0
                )
            )

            # Um contexto tributário útil pode tornar uma descrição aparentemente
            # genérica classificável (ex.: "Prestação de serviço" + "serviços técnicos
            # ... mecânica"). Mas não exibimos o contexto como se fosse a descrição
            # original.
            if _eh_descricao_generica_nfse(desc_exibicao) and desc_match != desc_exibicao:
                generic = len(_tokens_informativos(desc_match)) < 2

            confidence = best_score
            if best_row and ocr > 0:
                confidence = min(100.0, best_score * 0.85 + ocr * 0.15)

            if accepted:
                code = best_row.get("codigo", "")
                desc_db = best_row.get("descricao", "-")
                final = round(confidence, 1)
                obs = ""
            else:
                code = ""
                desc_db = "-"
                final = round(min(69.0, confidence) if best_row and generic else confidence, 1)
                if not best_row:
                    obs = "REVISAR: não foi possível correlacionar com o cadastro de serviços."
                elif generic:
                    obs = (
                        "REVISAR: a NFS-e não apresenta informação descritiva suficiente "
                        "para definir o serviço com segurança."
                    )
                elif best_score < 78.0:
                    obs = "REVISAR: correspondência insuficiente para definir o código de serviço com segurança."
                else:
                    obs = "REVISAR: há mais de uma possível correspondência de serviço com evidência semelhante."

        else:
            ranked = _candidatos_cadastro(
                desc_match,
                rows,
                ncm=str(item.get("ncm", "") or ""),
                service=False,
            )
            best = ranked[0] if ranked else None
            best_score = best[0] if best else 0.0
            best_index = best[1] if best else -1
            best_row = rows[best_index] if best else None
            base_ocr = ocr if ocr > 0 else 96.0
            confidence = round(min(100.0, best_score * 0.82 + base_ocr * 0.18), 1) if best_row else 0.0

            # Diferentemente da lógica anterior, dois candidatos próximos NÃO anulam
            # automaticamente o melhor. O melhor candidato preenche as colunas, e
            # os outros candidatos plausíveis aparecem nas observações.
            accepted = bool(best_row and best_score >= 58.0)

            if accepted:
                code = best_row.get("codigo", "")
                desc_db = best_row.get("descricao", "-")
                final = confidence
                outras = _formatar_outras_pecas(ranked, rows, best_index, best_score, forcar=(final < 70))
                if outras:
                    obs = outras
                    if final < 70:
                        obs += " | Conferência recomendada devido à confiança da correlação."
                elif final < 70:
                    obs = "Conferência recomendada: confiança da correlação abaixo de 70%."
                else:
                    obs = ""
            else:
                code = ""
                desc_db = "-"
                final = confidence
                outras = _formatar_outras_pecas(ranked, rows, best_index, best_score, forcar=True) if best else ""
                if not best_row:
                    obs = "REVISAR: não foi possível correlacionar com o cadastro de produtos."
                elif outras:
                    obs = (
                        "REVISAR: nenhuma correspondência atingiu segurança suficiente. "
                        + outras
                    )
                else:
                    obs = "REVISAR: correspondência insuficiente para definir o código do produto com segurança."

        base = {
            "Descrição na Nota": desc_exibicao,
            "Quantidade": item.get("quantidade", "1"),
            "Código no Banco de Dados": code,
            "Descrição no Banco de Dados": desc_db,
            "Centro de Custos": centro_custo,
            "Confiança": f"{round(final)}%",
            "Observações": obs,
        }

        if service:
            base = {
                "Número da Nota": item.get("numero", "SemNumero"),
                "Empresa": item.get("empresa", "Empresa não identificada"),
                **base,
                "Valor Total": item.get("valor_total", "-"),
                "Arquivo": item.get("arquivo", ""),
            }
        else:
            base = {
                "Número da Nota": item.get("numero", "SemNumero"),
                "Empresa": item.get("empresa", "Empresa não identificada"),
                **base,
                "Arquivo": item.get("arquivo", ""),
            }

        out.append(base)

    return out


def motivo_outro(tipo: str, page: dict) -> str:
    if page.get("error"): return page["error"]
    if tipo == "BOLETO": return "Boleto detectado; documentos bancários não são processados pelo sistema."
    if tipo == "CTE": return "CT-e/DACTE detectado; somente NF-e e NFS-e são processadas."
    if page.get("ocr") and page.get("conf", 0) < 50: return "Documento não legível o suficiente após OCR (baixa qualidade da imagem)."
    if len(page.get("text", "").strip()) < 30: return "Texto insuficiente para identificar o documento."
    return "Tipo de documento não identificado como NF-e ou NFS-e."


def style_workbook(writer) -> None:
    blue = PatternFill("solid", fgColor="D9EAF7")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    orange = PatternFill("solid", fgColor="FCE4D6")
    red = PatternFill("solid", fgColor="F4CCCC")
    for ws in writer.book.worksheets:
        ws.freeze_panes = "A2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = blue
            c.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=2):
            conf = 100
            for c in row:
                if str(ws.cell(1,c.column).value or "") == "Confiança":
                    m = re.search(r"\d+", str(c.value or ""))
                    conf = int(m.group()) if m else 0
                    break
            fill = red if conf < 50 else orange if conf < 70 else yellow if conf <= 80 else None
            if fill:
                for c in row: c.fill = fill
        for col in range(1, ws.max_column + 1):
            samples = [len(str(ws.cell(r,col).value or "")) for r in range(1, min(ws.max_row, 100)+1)]
            ws.column_dimensions[get_column_letter(col)].width = min(55, max(12, max(samples, default=12)+2))
            for r in range(1, ws.max_row+1):
                ws.cell(r,col).alignment = Alignment(vertical="top", wrap_text=True)


def load_fixed_cadastros() -> tuple[dict, dict]:
    return cadastro_preparado(PRODUTOS_XLSX), cadastro_preparado(SERVICOS_XLSX)


@app.post("/processar")
async def processar_notas(centroCusto: str = Form(...), notas: UploadFile = File(...)):
    try:
        produtos, servicos = load_fixed_cadastros()
    except Exception as e:
        raise RuntimeError(f"Falha ao carregar cadastros fixos: {e}")

    data = await notas.read()
    arquivos: list[tuple[str, bytes]] = []
    nome = notas.filename or "documento.pdf"
    arquivos_outros_iniciais: list[dict] = []
    if nome.lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for entry in z.infolist():
                    if entry.is_dir():
                        continue
                    entry_name = Path(entry.filename).name
                    if entry_name.lower().endswith(".pdf"):
                        arquivos.append((entry_name, z.read(entry.filename)))
                    else:
                        arquivos_outros_iniciais.append({
                            "Arquivo": entry_name,
                            "Tipo Detectado": "Arquivo não suportado",
                            "Motivo": "O ZIP contém um arquivo que não é PDF; o sistema processa somente NF-e/NFS-e em PDF.",
                            "Confiança": "100%",
                            "Observações": "Movido para Outros automaticamente."
                        })
        except zipfile.BadZipFile:
            raise RuntimeError("O arquivo enviado não é um ZIP válido.")
    elif nome.lower().endswith(".pdf"):
        arquivos = [(nome, data)]
    else:
        arquivos = [(nome, data)]

    nfe_rows: list[dict] = []
    nfse_rows: list[dict] = []
    outros_rows: list[dict] = list(arquivos_outros_iniciais)

    for nome_arq, pdf_bytes in arquivos:
        try:
            pages = process_pdf_pages(pdf_bytes)
        except Exception as e:
            outros_rows.append({"Arquivo": nome_arq, "Tipo Detectado": "Falha", "Motivo": "Falha no sistema ao processar o PDF.", "Confiança": "0%", "Observações": str(e)})
            continue

        combined_text = "\n".join(p.get("text","") for p in pages)
        combined_norm = normalizar_texto(combined_text)
        if "PRESTACAO DE SERVICOS" in combined_norm and "SERVICOS ALOCADOS" in combined_norm:
            d=extract_nfse(combined_text.splitlines())
            item={"numero":d["numero"],"empresa":d["empresa"],"descricao_nota":d["descricao"],"descricao_match":d.get("descricao_match",d["descricao"]),"quantidade":d["quantidade"],"valor_total":d["valor_total"],"confianca_ocr":max([float(p.get("conf",0) or 0) for p in pages] or [0]),"arquivo":nome_arq}
            nfse_rows.extend(correlacionar([item],servicos,centroCusto,"NFS-E"))
            continue

        for pnum, page in enumerate(pages, start=1):
            tipo = classify(page["text"])
            arquivo_label = nome_arq if len(pages) == 1 else f"{nome_arq} (página {pnum})"
            if tipo not in {"NFE", "NFSE"}:
                detection_conf = 100.0 if tipo in {"BOLETO", "CTE"} else float(page.get("conf", 0) or 0)
                outros_rows.append({"Arquivo": arquivo_label, "Tipo Detectado": {"BOLETO":"Boleto", "CTE":"CT-e"}.get(tipo, "Desconhecido"), "Motivo": motivo_outro(tipo, page), "Confiança": f"{round(detection_conf)}%", "Observações": "" if tipo in {"BOLETO","CTE"} else page.get("error", "")})
                continue

            if tipo == "NFSE":
                d = extract_nfse(page["lines"])
                item = {"numero": d["numero"], "empresa": d["empresa"], "descricao_nota": d["descricao"], "descricao_match": d.get("descricao_match", d["descricao"]), "quantidade": d["quantidade"], "valor_total": d["valor_total"], "confianca_ocr": page["conf"], "arquivo": arquivo_label}
                nfse_rows.extend(correlacionar([item], servicos, centroCusto, "NFS-E"))
            else:
                empresa = extract_nfe_emitter(page["lines"], page.get("image")); numero = extract_nfe_number(page["lines"])
                items = extract_nfe_items(page)
                if not items:
                    outros_rows.append({"Arquivo": arquivo_label, "Tipo Detectado": "NF-e", "Motivo": "NF-e identificada, mas nenhum item pôde ser extraído da tabela de produtos.", "Confiança": f"{round(page.get('conf',0))}%", "Observações": "Provável baixa qualidade do OCR ou layout não reconhecido."})
                    continue
                payload = [{"numero": numero, "empresa": empresa or "Empresa não identificada", "descricao_nota": x["descricao"], "quantidade": x["quantidade"], "ncm": x.get("ncm", ""), "confianca_ocr": page["conf"], "arquivo": arquivo_label} for x in items]
                nfe_rows.extend(correlacionar(payload, produtos, centroCusto, "NF-E"))

    if not nfe_rows: nfe_rows = [{"Aviso": "Nenhuma NF-e processada"}]
    if not nfse_rows: nfse_rows = [{"Aviso": "Nenhuma NFS-e processada"}]
    if not outros_rows: outros_rows = [{"Aviso": "Nenhum documento enviado para Outros"}]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(nfe_rows).to_excel(writer, index=False, sheet_name="Produtos (NF-e)")
        pd.DataFrame(nfse_rows).to_excel(writer, index=False, sheet_name="Serviços (NFS-e)")
        pd.DataFrame(outros_rows).to_excel(writer, index=False, sheet_name="Outros")
        style_workbook(writer)
    buf.seek(0)
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=Resultado_Notas.xlsx"})


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)