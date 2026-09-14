"""Motor de OCR local para documentos fiscais.

Tesseract + OpenCV, retornando texto, confiança, palavras e coordenadas.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from functools import lru_cache
from typing import Iterable

import pytesseract
from PIL import Image, ImageFilter, ImageOps

try:
    import cv2
    import numpy as np
except ImportError:  # pragma: no cover
    cv2 = None
    np = None


@dataclass
class OCRWord:
    text: str
    confidence: float
    left: int
    top: int
    width: int
    height: int
    block_num: int = 0
    par_num: int = 0
    line_num: int = 0

    @property
    def cx(self) -> float:
        return self.left + self.width / 2

    @property
    def cy(self) -> float:
        return self.top + self.height / 2


@dataclass
class OCRResult:
    text: str
    confidence: float
    words: list[OCRWord] = field(default_factory=list)
    lines: list[str] = field(default_factory=list)
    variant: str = ""


def configurar_tesseract() -> bool:
    candidatos = [
        os.getenv("TESSERACT_CMD", ""),
        r"C:\Users\otavio.seidinger\AppData\Local\Tesseract-OCR\tesseract.exe",
        os.path.join(os.path.expanduser("~"), r"AppData\Local\Tesseract-OCR\tesseract.exe"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.path.expanduser("~"), r"Tesseract-OCR\tesseract.exe"),
        "tesseract",
    ]
    for caminho in candidatos:
        try:
            if caminho != "tesseract" and not (caminho and os.path.isfile(caminho)):
                continue
            if caminho != "tesseract":
                pytesseract.pytesseract.tesseract_cmd = caminho
            pytesseract.get_tesseract_version()
            if caminho != "tesseract":
                tessdata = os.path.join(os.path.dirname(caminho), "tessdata")
                if os.path.isdir(tessdata):
                    os.environ.setdefault("TESSDATA_PREFIX", tessdata)
            return True
        except Exception:
            continue
    return False


HAS_OCR = configurar_tesseract()


@lru_cache(maxsize=1)
def linguagem_disponivel() -> str:
    try:
        langs = set(pytesseract.get_languages(config=""))
    except Exception:
        langs = set()
    if {"por", "eng"} <= langs:
        return "por+eng"
    if "por" in langs:
        return "por"
    if "eng" in langs:
        return "eng"
    return "eng"


def _preprocessamentos(imagem: Image.Image) -> Iterable[tuple[str, Image.Image]]:
    gray = ImageOps.autocontrast(ImageOps.grayscale(imagem))
    yield "gray", gray
    if cv2 is None or np is None:
        yield "sharp", gray.filter(ImageFilter.SHARPEN)
        return
    arr = np.asarray(gray)
    denoise = cv2.fastNlMeansDenoising(arr, None, 8, 7, 21)
    yield "denoise", Image.fromarray(denoise)
    _, otsu = cv2.threshold(denoise, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    yield "otsu", Image.fromarray(otsu)
    adaptive = cv2.adaptiveThreshold(
        denoise, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 11
    )
    yield "adaptive", Image.fromarray(adaptive)


def executar_passe(imagem: Image.Image, config: str = "--oem 3 --psm 6 -c preserve_interword_spaces=1") -> OCRResult:
    data = pytesseract.image_to_data(
        imagem, lang=linguagem_disponivel(), config=config,
        output_type=pytesseract.Output.DICT,
    )
    words: list[OCRWord] = []
    grupos: dict[tuple[int, int, int], list[OCRWord]] = {}
    confs: list[float] = []
    chars = 0
    n = len(data.get("text", []))
    for i in range(n):
        text = str(data["text"][i] or "").strip()
        if not text:
            continue
        try:
            conf = float(data["conf"][i])
        except Exception:
            conf = -1
        if conf < 0:
            continue
        w = OCRWord(
            text=text, confidence=conf,
            left=int(data["left"][i]), top=int(data["top"][i]),
            width=int(data["width"][i]), height=int(data["height"][i]),
            block_num=int(data.get("block_num", [0]*n)[i]),
            par_num=int(data.get("par_num", [0]*n)[i]),
            line_num=int(data.get("line_num", [0]*n)[i]),
        )
        words.append(w); confs.append(conf); chars += len(text)
        key=(w.block_num,w.par_num,w.line_num)
        grupos.setdefault(key, []).append(w)
    linhas_pos=[]
    for ws in grupos.values():
        ws.sort(key=lambda x:x.left)
        if ws:
            linhas_pos.append((min(x.top for x in ws), min(x.left for x in ws), " ".join(x.text for x in ws)))
    linhas=[x[2] for x in sorted(linhas_pos)]
    media=sum(confs)/len(confs) if confs else 0.0
    cobertura=min(1.0, chars/120.0)
    score=media*(0.65+0.35*cobertura)
    return OCRResult("\n".join(linhas), round(score,2), words, linhas, config)


def extract(imagem: Image.Image, detalhar: bool = True) -> OCRResult:
    if not HAS_OCR:
        return OCRResult("",0.0,[],[],"indisponivel")
    base=ImageOps.autocontrast(ImageOps.grayscale(imagem))
    melhor=None
    try:
        melhor=executar_passe(base)
        melhor.variant="gray:psm6"
    except Exception:
        pass
    precisa=detalhar and (melhor is None or melhor.confidence < 82 or len(melhor.text)<180)
    if precisa:
        for nome,img in _preprocessamentos(imagem):
            if nome=="gray":
                continue
            for cfg in ("--oem 3 --psm 6 -c preserve_interword_spaces=1","--oem 3 --psm 11"):
                try:
                    r=executar_passe(img,cfg); r.variant=f"{nome}:{cfg}"
                except Exception:
                    continue
                if melhor is None or (r.confidence,len(r.text))>(melhor.confidence,len(melhor.text)):
                    melhor=r
    return melhor or OCRResult("",0.0,[],[],"falha")
