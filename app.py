# ============================================================
# FECHAMENTOS PREMIER
# Versão consolidada: Alex, Oscar, Demetra e Strong.
# OCR das tabelas por nome do agente, sem depender da ordem das linhas.
# ============================================================

import io
import re
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps

try:
    import cv2
    import numpy as np
    CV2_OK = True
except Exception:
    cv2 = None
    np = None
    CV2_OK = False

try:
    import pytesseract
    TESSERACT_OK = True
except Exception:
    pytesseract = None
    TESSERACT_OK = False

st.set_page_config(page_title="Fechamentos Premier", layout="wide")

# =========================
# CONFIGURAÇÕES
# =========================
RB_DEMETRA_PLANILHA = 70.0
RB_DEMETRA_IMAGEM = 70.0
REBATE_DEMETRA = -5.0
REBATE_OSCAR = -10.0
REBATE_ALEX_POSITIVO = -5.0
REBATE_ALEX_NEGATIVO = 5.0

RB_ALEX_IMAGEM_REAL = 65.0
RB_ALEX_IMAGEM_EXIBIDO = 70.0

AGENTE_PLANILHA_DEMETRA = "TheShark_ (ID11719117)"
ORIGEM_PLANILHA_DEMETRA = "TheShark_"
ID_PLANILHA_DEMETRA = "11719117"
CODIGO_PLANILHA_DEMETRA = "802606"

MAPA_IDS_PDF = {
    "12968708": {"cliente": "Demetra", "rb": 70.0},
    "1607968": {"cliente": "Demetra", "rb": 65.0},
    "1527106": {"cliente": "Demetra", "rb": 65.0},
    "13357678": {"cliente": "Demetra", "rb": 65.0},
    "13648965": {"cliente": "Demetra", "rb": 60.0},
    "13713445": {"cliente": "Demetra", "rb": 60.0},
    "12970177": {"cliente": "Alex", "rb": 70.0},
    "13019559": {"cliente": "Alex", "rb": 50.0},
    "13018880": {"cliente": "Alex", "rb": 45.0},
    "13213751": {"cliente": "Alex", "rb": 70.0},
    "13265647": {"cliente": "Alex", "rb": 50.0},
    "13319248": {"cliente": "Alex", "rb": 70.0},
    "13379845": {"cliente": "Alex", "rb": 60.0},
    "13104440": {"cliente": "Alex", "rb": 50.0},
    "13590389": {"cliente": "Alex", "rb": 60.0},
    "13472941": {"cliente": "Oscar", "rb": 65.0},
    "13470981": {"cliente": "Oscar", "rb": 65.0},
    "2733985": {"cliente": "Oscar", "rb": 65.0},
    "13489882": {"cliente": "Oscar", "rb": 65.0},
    "3891202": {"cliente": "Oscar", "rb": 65.0},
    "4085350": {"cliente": "Oscar", "rb": 45.0},
    "13696313": {"cliente": "Demetra", "rb": 60.0},
    "0": {"cliente": "Oscar", "rb": 40.0},
}

# =========================
# CORES / LAYOUT
# =========================
NAVY = (7, 29, 69)
GOLD = (199, 143, 43)
GREEN = (0, 120, 0)
RED = (180, 0, 0)
WHITE = (255, 255, 255)
GRAY = (245, 245, 245)
LIGHT_BG = (234, 241, 235)
YELLOW = (248, 238, 27)
LIGHT_GRAY = (230, 230, 230)
GRID = (90, 90, 90)

FONT_TITLE = 45
FONT_SUBTITLE = 36
FONT_HEADER = 30
FONT_CELL = 30
FONT_AGENT = 24
FONT_TOTAL = 30
FONT_STATUS = 36
FONT_STATUS_VALUE = 36

SUMMARY_W = 1400
SUMMARY_H = 980
TABLE_W = 1450
TABLE_ROW_H_MIN = 52
TABLE_TOP = 220
TABLE_HEADER_H = 52


# =========================
# UTILITÁRIOS
# =========================
def fmt_brl(v: float) -> str:
    s = f"{float(v):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def parse_money(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    text = text.replace("R$", "").replace(" ", "")
    text = text.replace("—", "-").replace("–", "-").replace("−", "-")
    text = text.replace("(", "").replace(")", "")
    text = re.sub(r"[^0-9,.\-]", "", text)

    if text in {"", "-", ".", "-."}:
        return 0.0

    if re.match(r"^-?\d{1,3}(,\d{3})+\.\d{1,2}$", text):
        text = text.replace(",", "")
        return float(text)

    if re.match(r"^-?\d{1,3}(\.\d{3})+,\d{1,2}$", text):
        text = text.replace(".", "").replace(",", ".")
        return float(text)

    if "," in text and "." not in text:
        text = text.replace(",", ".")
        return float(text)

    if "." in text and "," not in text:
        return float(text)

    if "," in text and "." in text:
        if text.rfind(".") > text.rfind(","):
            text = text.replace(",", "")
        else:
            text = text.replace(".", "").replace(",", ".")
        return float(text)

    return float(text)


def parse_money_misto(text: str) -> float:
    text = str(text).strip()
    text = text.replace("R$", "").replace(" ", "")
    text = text.replace("—", "-").replace("–", "-").replace("−", "-")
    text = re.sub(r"[^0-9,.\-]", "", text)

    if not text or text in {"-", ".", ","}:
        return 0.0

    if re.match(r"^-?\d{1,3}(,\d{3})+\.\d{1,2}$", text):
        return float(text.replace(",", ""))

    if re.match(r"^-?\d{1,3}(\.\d{3})+,\d{1,2}$", text):
        return float(text.replace(".", "").replace(",", "."))

    if "," in text and "." in text:
        if text.rfind(".") > text.rfind(","):
            return float(text.replace(",", ""))
        return float(text.replace(".", "").replace(",", "."))

    if "," in text:
        return float(text.replace(",", "."))

    return float(text)


def extract_all_money_misto(text: str):
    matches = re.findall(
        r"-?\d{1,3}(?:[,.]\d{3})+[,.]\d{1,2}|-?\d+[,.]\d{1,2}",
        text,
    )
    return [parse_money_misto(m) for m in matches]


def to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def get_font(size: int, bold: bool = False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for p in candidates:
        if Path(p).exists():
            return ImageFont.truetype(p, size=size)
    return ImageFont.load_default()


def measure(draw, text, font):
    box = draw.textbbox((0, 0), text, font=font)
    return box[2] - box[0], box[3] - box[1]


def normalize_id(v) -> str:
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[^\d]", "", s)


def wrap_text(draw, text, font, max_width):
    words = str(text).split()
    lines = []
    current = ""
    for word in words:
        candidate = current + (" " if current else "") + word
        w, _ = measure(draw, candidate, font)
        if w <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines if lines else [""]


# =========================
# OCR
# =========================
def preprocess_for_ocr(img: Image.Image) -> Image.Image:
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    g = ImageEnhance.Contrast(g).enhance(2.8)
    g = g.filter(ImageFilter.SHARPEN)
    return g


def ocr_image(img: Image.Image, psm: int = 6) -> str:
    if not TESSERACT_OK:
        return ""
    proc = preprocess_for_ocr(img)
    try:
        return pytesseract.image_to_string(proc, lang="eng", config=f"--oem 3 --psm {psm}") or ""
    except Exception:
        return ""


def first_money(text: str) -> float:
    for pat in [
        r"([\-]?[0-9]{1,3}(?:[,\.][0-9]{3})*[,\.][0-9]{2})",
        r"([\-]?[0-9]+[,\.][0-9]{2})",
    ]:
        m = re.search(pat, text)
        if m:
            return parse_money(m.group(1))
    return 0.0


def extract_all_money(text: str):
    matches = re.findall(r"[\-]?[0-9]{1,3}(?:[,\.][0-9]{3})*[,\.][0-9]{2}|[\-]?[0-9]+[,\.][0-9]{2}", text)
    return [parse_money(m) for m in matches]


def ocr_crop_value(img: Image.Image, box) -> tuple[str, float]:
    crop = img.crop(box)
    txt = ocr_image(crop, psm=6) + "\n" + ocr_image(crop, psm=11)
    vals = extract_all_money(txt)
    for v in vals:
        if abs(v) < 100000:
            return txt, v
    return txt, 0.0



# =========================
# OSCAR - SUPREMA (imagem)
# =========================
def detect_suprema_image(img: Image.Image) -> bool:
    text = (ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)).upper()
    return "SUPREMA" in text and ("W/L" in text or "RAKE" in text or "RB" in text)


def extract_suprema_values(img: Image.Image) -> dict:
    """
    Lê a imagem da SUPREMA:
    - W/L = ganhos, apenas exibido
    - RAKE = rake
    - TOTAL do Oscar = rake * 65%, sem somar ganhos
    """
    w, h = img.size

    wl_box = (int(w * 0.30), int(h * 0.42), int(w * 0.58), int(h * 0.63))
    rake_box = (int(w * 0.55), int(h * 0.42), int(w * 0.80), int(h * 0.63))
    total_box = (int(w * 0.45), int(h * 0.63), int(w * 0.78), int(h * 0.83))

    wl_txt, ganhos = ocr_crop_value(img, wl_box)
    rake_txt, rake = ocr_crop_value(img, rake_box)
    total_txt, total_imagem = ocr_crop_value(img, total_box)

    text = ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)

    total_calculado = rake * 0.65
    diferenca = abs(total_calculado - total_imagem)
    leitura_valida = diferenca <= 3.0

    if not leitura_valida:
        vals = extract_all_money(text)
        if len(vals) >= 4:
            ganhos2 = vals[0]
            rake2 = vals[1]
            total2 = vals[3]
            total_calc2 = rake2 * 0.65
            if abs(total_calc2 - total2) <= 3.0:
                ganhos = ganhos2
                rake = rake2
                total_imagem = total2
                leitura_valida = True

    return {
        "agente": "SUPREMA | Agents",
        "ganhos": ganhos,
        "rake": rake,
        "rb_percentual": 65.0,
        "total_imagem": total_imagem,
        "total_calculado": rake * 0.65,
        "leitura_valida": leitura_valida,
        "ocr_text": (
            text
            + "\n\nOCR W/L:\n" + wl_txt
            + "\n\nOCR RAKE:\n" + rake_txt
            + "\n\nOCR TOTAL:\n" + total_txt
        ),
    }


# =========================
# ALEX - CASARICA
# =========================
def detect_alex_casarica_image(img: Image.Image) -> bool:
    text = (ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)).upper()
    return (
        ("GUSTAVO" in text and "CASARICA" in text)
        or ("PROFIT" in text and "LOSS" in text and "RAKE" in text)
        or ("TAXA" in text and "GANHOS" in text)
    )


def extract_labeled_money(text: str, label_pattern: str) -> float:
    m = re.search(label_pattern + r"[^0-9\-]*([\-]?[0-9][0-9\.,]+)", text, flags=re.IGNORECASE)
    if m:
        return parse_money(m.group(1))
    return 0.0


def extract_alex_casarica_values(img: Image.Image) -> dict:
    text = ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)
    text_upper = text.upper()

    if "PROFIT" in text_upper and "LOSS" in text_upper and "RAKE" in text_upper:
        ganhos = extract_labeled_money(text, r"PROFIT\s*/\s*LOSS")
        rake70 = extract_labeled_money(text, r"RAKE\s*70%?")
        rakeback = rake70 * (RB_ALEX_IMAGEM_REAL / RB_ALEX_IMAGEM_EXIBIDO)
        total_base = ganhos + rakeback
        if total_base > 0:
            rebate = total_base * (REBATE_ALEX_POSITIVO / 100.0)
        elif total_base < 0:
            rebate = total_base * (REBATE_ALEX_NEGATIVO / 100.0)
        else:
            rebate = 0.0
        total_final = total_base + rebate
        return {
            "agente": "Gustavo | Casarica",
            "ganhos": ganhos,
            "rake": rake70,
            "rb_percentual": RB_ALEX_IMAGEM_REAL,
            "total_base": total_base,
            "rebate": rebate,
            "total_final": total_final,
            "ocr_text": text,
            "formato": "profit_loss",
            "ocr_taxa_crop": "",
            "ocr_ganhos_crop": "",
        }

    w, h = img.size
    taxa_box = (int(w * 0.06), int(h * 0.48), int(w * 0.34), int(h * 0.68))
    ganhos_box = (int(w * 0.66), int(h * 0.48), int(w * 0.94), int(h * 0.68))

    taxa_txt, rake = ocr_crop_value(img, taxa_box)
    ganhos_txt, ganhos = ocr_crop_value(img, ganhos_box)

    if abs(rake) > 50000:
        rake = 0.0
    if abs(ganhos) > 50000:
        ganhos = 0.0

    if rake == 0.0 and ganhos == 0.0:
        return {
            "agente": "Gustavo | Casarica",
            "ganhos": 0.0,
            "rake": 0.0,
            "rb_percentual": RB_ALEX_IMAGEM_REAL,
            "total_base": 0.0,
            "rebate": 0.0,
            "total_final": 0.0,
            "ocr_text": text,
            "formato": "taxa_ganhos_falhou",
            "ocr_taxa_crop": taxa_txt,
            "ocr_ganhos_crop": ganhos_txt,
        }

    rakeback = rake * (RB_ALEX_IMAGEM_REAL / 100.0)
    total_base = ganhos + rakeback

    if total_base > 0:
        rebate = total_base * (REBATE_ALEX_POSITIVO / 100.0)
    elif total_base < 0:
        rebate = total_base * (REBATE_ALEX_NEGATIVO / 100.0)
    else:
        rebate = 0.0

    total_final = total_base + rebate
    return {
        "agente": "Gustavo | Casarica",
        "ganhos": ganhos,
        "rake": rake,
        "rb_percentual": RB_ALEX_IMAGEM_REAL,
        "total_base": total_base,
        "rebate": rebate,
        "total_final": total_final,
        "ocr_text": text,
        "formato": "taxa_ganhos_recorte",
        "ocr_taxa_crop": taxa_txt,
        "ocr_ganhos_crop": ganhos_txt,
    }


# =========================
# DEMETRA
# =========================
def process_demetra_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, usecols=[5, 6, 7, 8, 9, 29])
    df.columns = ["origem", "id_conta", "nick", "codigo", "ganhos", "rake"]
    df["id_conta"] = df["id_conta"].apply(normalize_id)
    df["codigo"] = df["codigo"].apply(normalize_id)
    df["origem"] = df["origem"].astype(str).str.strip()
    df["ganhos"] = pd.to_numeric(df["ganhos"], errors="coerce").fillna(0.0)
    df["rake"] = pd.to_numeric(df["rake"], errors="coerce").fillna(0.0)
    df = df[df["codigo"] == CODIGO_PLANILHA_DEMETRA].copy()
    if df.empty:
        return pd.DataFrame()
    mask = (df["id_conta"] == ID_PLANILHA_DEMETRA) & (df["origem"] == ORIGEM_PLANILHA_DEMETRA)
    return df[mask].copy()


def extract_demetra_killuminatti_novo(img: Image.Image) -> dict:
    w, h = img.size

    taxa_box = (int(w * 0.02), int(h * 0.54), int(w * 0.40), int(h * 0.72))
    ganhos_box = (int(w * 0.64), int(h * 0.54), int(w * 0.99), int(h * 0.72))
    retorno_box = (int(w * 0.35), int(h * 0.73), int(w * 0.67), int(h * 0.93))
    total_box = (int(w * 0.66), int(h * 0.73), int(w * 0.99), int(h * 0.93))

    def crop_ocr_money(box):
        crop = img.crop(box)
        crop = crop.resize((crop.width * 4, crop.height * 4))
        txt = ocr_image(crop, psm=6) + "\n" + ocr_image(crop, psm=7) + "\n" + ocr_image(crop, psm=11)
        vals = extract_all_money_misto(txt)
        return txt, vals

    taxa_txt, taxa_vals = crop_ocr_money(taxa_box)
    ganhos_txt, ganhos_vals = crop_ocr_money(ganhos_box)
    retorno_txt, retorno_vals = crop_ocr_money(retorno_box)
    total_txt, total_vals = crop_ocr_money(total_box)

    rake = taxa_vals[0] if taxa_vals else 0.0
    ganhos = ganhos_vals[0] if ganhos_vals else 0.0
    retorno_taxa = retorno_vals[0] if retorno_vals else 0.0
    total_tela = total_vals[0] if total_vals else 0.0

    if retorno_taxa > 0:
        rake_por_retorno = retorno_taxa / 0.75
        if rake == 0 or abs((rake * 0.75) - retorno_taxa) > 10:
            rake = rake_por_retorno

    if total_tela > 0 and retorno_taxa > 0:
        ganhos_por_total = total_tela - retorno_taxa
        if ganhos == 0 or abs((ganhos + retorno_taxa) - total_tela) > 10:
            ganhos = ganhos_por_total

    if ganhos > 0 and ganhos < 100 and total_tela > 100 and retorno_taxa > 100:
        ganhos = total_tela - retorno_taxa

    text = ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)

    return {
        "agente": "Killuminatti",
        "ganhos": ganhos,
        "rake": rake,
        "rb_percentual": RB_DEMETRA_IMAGEM,
        "ocr_text": (
            text
            + "\n\nOCR TAXA:\n" + taxa_txt
            + "\n\nOCR GANHOS:\n" + ganhos_txt
            + "\n\nOCR RETORNO:\n" + retorno_txt
            + "\n\nOCR GANHOS + RETORNO:\n" + total_txt
            + f"\n\nGANHOS FINAL VALIDADO: {ganhos}"
            + f"\nRAKE FINAL VALIDADO: {rake}"
            + f"\nRETORNO DE TAXA LIDO: {retorno_taxa}"
            + f"\nGANHOS + RETORNO LIDO: {total_tela}"
        ),
    }


def detect_demetra_image(img: Image.Image) -> bool:
    text = (ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)).upper()
    return ("KILLUMINATTI" in text or "SUPER AGENTE" in text) and ("RAKE" in text or "GANHOS" in text)


def extract_demetra_image_values(img: Image.Image) -> dict:
    """
    Lê a imagem/tabela do Demetra no formato Killuminatti 25%.

    Correção:
    - soma a coluna RAKE de todas as linhas de agente;
    - soma a coluna GANHOS de todas as linhas de agente;
    - ignora as colunas -25%, Resultado e TOTAL;
    - a linha ADIANTAMENTO pode aparecer como 0,00 e não altera a soma.

    Exemplo esperado:
    RAKE = 6751,75 + 23,05 = 6774,80
    GANHOS = -2231,40 + 592,35 = -1639,05
    """
    text_global = ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)
    w, h = img.size

    def ocr_money_column(box, label):
        crop = img.crop(box)
        # aumenta o recorte para melhorar OCR de números pequenos
        crop_big = crop.resize((crop.width * 4, crop.height * 4))
        txt = (
            ocr_image(crop_big, psm=6)
            + "\n"
            + ocr_image(crop_big, psm=7)
            + "\n"
            + ocr_image(crop_big, psm=11)
        )
        vals = extract_all_money_misto(txt)

        # proteção: remove leituras absurdas que normalmente vêm de ruído/OCR
        vals = [v for v in vals if abs(v) < 100000]

        return txt, vals

    # Recortes por coluna para a tabela "Killuminatti 25%".
    # As proporções foram feitas para pegar somente as linhas de dados,
    # sem cabeçalho e sem a faixa TOTAL no rodapé.
    rake_box = (
        int(w * 0.28),
        int(h * 0.36),
        int(w * 0.43),
        int(h * 0.79),
    )
    ganhos_box = (
        int(w * 0.56),
        int(h * 0.36),
        int(w * 0.72),
        int(h * 0.79),
    )

    rake_txt, rake_vals = ocr_money_column(rake_box, "RAKE")
    ganhos_txt, ganhos_vals = ocr_money_column(ganhos_box, "GANHOS")

    rake = sum(rake_vals)
    ganhos = sum(ganhos_vals)

    # Fallback por linhas: usa somente linhas que pareçam ser agentes.
    # Ignora cabeçalho, adiantamento, total e resultado final.
    if rake == 0.0 and ganhos == 0.0:
        linhas = [ln.strip() for ln in text_global.splitlines() if ln.strip()]
        for line in linhas:
            upper = line.upper()
            if any(palavra in upper for palavra in ["SUPER", "AGENTE", "RAKE", "GANHOS", "RESULTADO", "ADIANTAMENTO", "TOTAL"]):
                continue

            vals = extract_all_money_misto(line)
            if len(vals) >= 3:
                # padrão da linha:
                # agente | rake | -25% | ganhos | resultado
                rake += vals[0]
                ganhos += vals[2]

    return {
        "agente": "Killuminatti",
        "ganhos": ganhos,
        "rake": rake,
        "rb_percentual": RB_DEMETRA_IMAGEM,
        "ocr_text": (
            text_global
            + "\n\n=== LEITURA POR TABELA/COLUNAS ==="
            + "\nOCR COLUNA RAKE:\n" + rake_txt
            + "\nVALORES RAKE LIDOS: " + str(rake_vals)
            + "\nSOMA RAKE: " + str(rake)
            + "\n\nOCR COLUNA GANHOS:\n" + ganhos_txt
            + "\nVALORES GANHOS LIDOS: " + str(ganhos_vals)
            + "\nSOMA GANHOS: " + str(ganhos)
        ),
    }

# =========================
# PDF
# =========================
# =========================
# PDF
# =========================
def extract_pdf_lines(uploaded_file):
    lines = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend([ln.strip() for ln in text.splitlines() if ln.strip()])
    return lines


def process_pdf_by_client(uploaded_file, cliente_alvo: str):
    """Lê linhas do PDF e aplica o RB% cadastrado no MAPA_IDS_PDF.

    Para Oscar, também captura a terceira coluna monetária como REBATE.
    O RAKEBACK presente no PDF nunca é usado como percentual.
    """
    rows = []
    for line in extract_pdf_lines(uploaded_file):
        if "R$" not in line:
            continue
        id_match = re.search(r"\b(0|\d{6,9})\b", line)
        if not id_match:
            continue
        id_agente = normalize_id(id_match.group(1))
        info = MAPA_IDS_PDF.get(id_agente)
        if not info or info["cliente"] != cliente_alvo:
            continue

        money_matches = re.findall(r"R\$\s*-?\d[\d\.,]*", line)
        if len(money_matches) < 2:
            continue

        ganhos = parse_money(money_matches[0])
        rake = parse_money(money_matches[1])
        rebate = parse_money(money_matches[2]) if cliente_alvo == "Oscar" and len(money_matches) >= 3 else 0.0
        agente = re.sub(r"\s{2,}", " ", line[:id_match.start()].strip())
        rows.append({
            "agente": agente,
            "id_agente": id_agente,
            "ganhos": ganhos,
            "rake": rake,
            "rebate": rebate,
            "rb_percentual": float(info["rb"]),
        })
    return pd.DataFrame(rows)


# =========================
# OCR DINÂMICO DE TABELA
# =========================
def _ocr_data(img: Image.Image):
    """Executa OCR na imagem inteira e preserva as coordenadas de cada token."""
    if not TESSERACT_OK:
        return None
    try:
        proc = preprocess_for_ocr(img)
        data = pytesseract.image_to_data(
            proc,
            lang="eng",
            config="--oem 3 --psm 6",
            output_type=pytesseract.Output.DATAFRAME,
        )
        if data is None or data.empty:
            return None

        data = data.dropna(subset=["text"]).copy()
        data["text"] = data["text"].astype(str).str.strip()
        data = data[data["text"] != ""].copy()
        for col in ["left", "top", "width", "height", "conf"]:
            data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

        data["x_center"] = data["left"] + data["width"] / 2
        data["y_center"] = data["top"] + data["height"] / 2
        return data
    except Exception:
        return None


def _norm_name(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def _name_similarity(candidate: str, target: str) -> float:
    """Pontuação tolerante a pequenas falhas do OCR."""
    candidate = _norm_name(candidate)
    target = _norm_name(target)
    if not candidate or not target:
        return 0.0
    if candidate == target:
        return 1.0
    if candidate in target or target in candidate:
        return 0.92
    return SequenceMatcher(None, candidate, target).ratio()


def _cluster_ocr_rows(data: pd.DataFrame) -> list[dict]:
    """Agrupa tokens em linhas usando a posição vertical real, não a ordem do OCR."""
    if data is None or data.empty:
        return []

    heights = data.loc[data["height"] > 0, "height"]
    median_h = float(heights.median()) if not heights.empty else 16.0
    tolerance = max(7.0, median_h * 0.75)

    tokens = data.sort_values(["y_center", "left"]).to_dict("records")
    groups: list[list[dict]] = []
    centers: list[float] = []

    for token in tokens:
        y = float(token["y_center"])
        best_idx = None
        best_dist = None
        for idx, center in enumerate(centers):
            dist = abs(y - center)
            if dist <= tolerance and (best_dist is None or dist < best_dist):
                best_idx = idx
                best_dist = dist

        if best_idx is None:
            groups.append([token])
            centers.append(y)
        else:
            groups[best_idx].append(token)
            centers[best_idx] = sum(float(t["y_center"]) for t in groups[best_idx]) / len(groups[best_idx])

    rows = []
    for group in groups:
        group = sorted(group, key=lambda t: float(t["left"]))
        text = " ".join(str(t["text"]) for t in group if str(t["text"]).strip())
        rows.append({
            "tokens": group,
            "text": text,
            "norm": _norm_name(text),
            "y": sum(float(t["y_center"]) for t in group) / len(group),
            "top": min(float(t["top"]) for t in group),
            "bottom": max(float(t["top"] + t["height"]) for t in group),
        })

    return sorted(rows, key=lambda r: r["y"])


def _best_token_center(row: dict, aliases: list[str]) -> tuple[float | None, float]:
    """Retorna o centro X do token mais semelhante a um cabeçalho."""
    best_x = None
    best_score = 0.0
    for token in row["tokens"]:
        token_text = str(token["text"])
        for alias in aliases:
            score = _name_similarity(token_text, alias)
            if score > best_score:
                best_score = score
                best_x = float(token["x_center"])
    return best_x, best_score


def _detect_table_structure(img: Image.Image) -> dict:
    """Detecta cabeçalhos, centros de colunas e linhas da tabela inteira."""
    data = _ocr_data(img)
    rows = _cluster_ocr_rows(data) if data is not None else []
    if not rows:
        return {"ok": False, "reason": "OCR não encontrou linhas na imagem.", "rows": []}

    header = None
    header_score = 0.0
    for row in rows:
        rake_x, rake_score = _best_token_center(row, ["rake"])
        ganhos_x, ganhos_score = _best_token_center(row, ["ganhos", "ganho"])
        score = rake_score + ganhos_score
        if rake_x is not None and ganhos_x is not None and score > header_score:
            header = row
            header_score = score

    if header is None:
        return {"ok": False, "reason": "Cabeçalhos RAKE e GANHOS não foram localizados.", "rows": rows}

    header_aliases = {
        "agente": ["super", "agente"],
        "rake": ["rake"],
        "desconto": ["25", "-25", "25%"],
        "ganhos": ["ganhos", "ganho"],
        "resultado": ["resultado", "result"] ,
    }

    centers = {}
    for name, aliases in header_aliases.items():
        x, score = _best_token_center(header, aliases)
        if x is not None and score >= 0.35:
            centers[name] = x

    if "rake" not in centers or "ganhos" not in centers:
        return {"ok": False, "reason": "Não foi possível definir as colunas RAKE e GANHOS.", "rows": rows}

    # Cria limites de coluna a partir dos centros detectados no próprio cabeçalho.
    ordered = sorted((x, name) for name, x in centers.items())
    bounds = {}
    for i, (x, name) in enumerate(ordered):
        left = 0.0 if i == 0 else (ordered[i - 1][0] + x) / 2
        right = float(img.width) if i == len(ordered) - 1 else (x + ordered[i + 1][0]) / 2
        bounds[name] = (left, right)

    data_rows = [r for r in rows if r["y"] > header["y"] + max(5.0, (header["bottom"] - header["top"]) * 0.5)]
    return {
        "ok": True,
        "rows": rows,
        "header": header,
        "centers": centers,
        "bounds": bounds,
        "data_rows": data_rows,
    }


def _numeric_ocr_from_cell(img: Image.Image, box: tuple[int, int, int, int]) -> tuple[str, float]:
    """Lê uma célula já localizada dinamicamente pelos cabeçalhos e pela linha."""
    x1, y1, x2, y2 = box
    if x2 <= x1 or y2 <= y1:
        return "", 0.0

    # Pequena margem interna para evitar que as linhas da grade prejudiquem o OCR.
    mx = max(2, int((x2 - x1) * 0.04))
    my = max(1, int((y2 - y1) * 0.10))
    x1, x2 = x1 + mx, x2 - mx
    y1, y2 = y1 + my, y2 - my
    crop = img.crop((x1, y1, x2, y2))
    scale = 5
    crop_big = crop.resize((max(1, crop.width * scale), max(1, crop.height * scale)))

    gray = ImageOps.autocontrast(crop_big.convert("L"))
    threshold = gray.point(lambda p: 255 if p > 170 else 0)
    variants = [crop_big, gray, threshold]
    texts = []

    configs = [
        "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789.,-",
        "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789.,-",
        "--oem 3 --psm 11 -c tessedit_char_whitelist=0123456789.,-",
    ]

    for variant in variants:
        for config in configs:
            try:
                txt = pytesseract.image_to_string(variant, lang="eng", config=config) or ""
            except Exception:
                txt = ""
            if txt.strip():
                texts.append(txt.strip())

    joined = "\n".join(texts)
    candidates = []
    for txt in texts:
        candidates.extend(extract_all_money_misto(txt))

    candidates = [v for v in candidates if abs(v) < 1_000_000]
    if not candidates:
        return joined, 0.0

    # O mesmo valor costuma aparecer repetido em vários modos; usa o mais frequente.
    rounded = [round(v, 2) for v in candidates]
    value = max(set(rounded), key=rounded.count)
    return joined, float(value)


def _find_agent_row(structure: dict, aliases: list[str]) -> tuple[dict | None, float]:
    targets = [_norm_name(a) for a in aliases if _norm_name(a)]
    best_row = None
    best_score = 0.0

    for row in structure.get("data_rows", []):
        full_score = max((_name_similarity(row["text"], target) for target in targets), default=0.0)
        token_score = 0.0
        for token in row["tokens"]:
            token_score = max(
                token_score,
                max((_name_similarity(str(token["text"]), target) for target in targets), default=0.0),
            )
        score = max(full_score, token_score)
        if score > best_score:
            best_score = score
            best_row = row

    if best_score < 0.45:
        return None, best_score
    return best_row, best_score


def _row_vertical_bounds(structure: dict, target_row: dict, img_height: int) -> tuple[int, int]:
    rows = structure.get("data_rows", [])
    idx = rows.index(target_row)
    current_y = float(target_row["y"])

    if idx > 0:
        top = int((float(rows[idx - 1]["y"]) + current_y) / 2)
    else:
        header_y = float(structure["header"]["y"])
        top = int((header_y + current_y) / 2)

    if idx < len(rows) - 1:
        bottom = int((current_y + float(rows[idx + 1]["y"])) / 2)
    else:
        row_h = max(18.0, float(target_row["bottom"] - target_row["top"]))
        bottom = int(min(img_height, current_y + row_h * 1.2))

    return max(0, top), min(img_height, bottom)



def _merge_close_positions(values: list[float], tolerance: float = 4.0) -> list[int]:
    if not values:
        return []
    values = sorted(float(v) for v in values)
    groups = [[values[0]]]
    for value in values[1:]:
        if value - groups[-1][-1] <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [int(round(sum(group) / len(group))) for group in groups]


def _detect_table_grid(img: Image.Image) -> dict:
    """Detecta automaticamente as linhas da grade da tabela na imagem inteira."""
    if not CV2_OK:
        return {"ok": False, "reason": "OpenCV indisponível."}

    rgb = np.array(img.convert("RGB"))
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    binary = cv2.threshold(gray, 185, 255, cv2.THRESH_BINARY_INV)[1]
    height, width = gray.shape

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, width // 15), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(15, height // 8)))
    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
    vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel)

    h_contours, _ = cv2.findContours(horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    v_contours, _ = cv2.findContours(vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    h_positions = []
    h_segments = []
    for contour in h_contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w >= width * 0.38 and h <= max(8, height * 0.025):
            h_positions.append(y + h / 2)
            h_segments.append((x, y, w, h))

    v_positions = []
    for contour in v_contours:
        x, y, w, h = cv2.boundingRect(contour)
        if h >= height * 0.30 and w <= max(8, width * 0.01):
            v_positions.append(x + w / 2)

    ys = _merge_close_positions(h_positions, tolerance=max(3.0, height * 0.008))
    xs = _merge_close_positions(v_positions, tolerance=max(3.0, width * 0.004))

    # A borda esquerda às vezes encosta no limite da imagem e não vira contorno vertical.
    if h_segments:
        table_left = min(seg[0] for seg in h_segments)
        table_right = max(seg[0] + seg[2] for seg in h_segments)
        if not xs or abs(xs[0] - table_left) > width * 0.03:
            xs.insert(0, int(table_left))
        if abs(xs[-1] - table_right) > width * 0.03:
            xs.append(int(table_right))

    xs = sorted(set(max(0, min(width - 1, int(x))) for x in xs))
    ys = sorted(set(max(0, min(height - 1, int(y))) for y in ys))

    if len(xs) < 4 or len(ys) < 3:
        return {
            "ok": False,
            "reason": f"Grade insuficiente: {len(xs)} linhas verticais e {len(ys)} horizontais.",
            "xs": xs,
            "ys": ys,
        }

    return {"ok": True, "xs": xs, "ys": ys}


def _ocr_cell_text(img: Image.Image, box: tuple[int, int, int, int], numeric: bool = False) -> str:
    x1, y1, x2, y2 = box
    if x2 <= x1 or y2 <= y1:
        return ""
    mx = max(2, int((x2 - x1) * 0.02))
    my = max(1, int((y2 - y1) * 0.06))
    crop = img.crop((x1 + mx, y1 + my, x2 - mx, y2 - my))
    scale = 6 if numeric else 4
    crop = crop.resize((max(1, crop.width * scale), max(1, crop.height * scale)))

    gray = ImageOps.autocontrast(crop.convert("L"))
    threshold = gray.point(lambda p: 255 if p > 170 else 0)
    variants = [crop, threshold, gray]
    configs = []
    if numeric:
        configs = [
            "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789.,-",
            "--oem 3 --psm 13 -c tessedit_char_whitelist=0123456789.,-",
            "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789.,-",
        ]
    else:
        configs = ["--oem 3 --psm 7", "--oem 3 --psm 6", "--oem 3 --psm 11"]

    best = ""
    for variant in variants:
        for config in configs:
            try:
                text = pytesseract.image_to_string(variant, lang="eng", config=config).strip()
            except Exception:
                text = ""
            if text:
                if numeric and extract_all_money_misto(text):
                    return text
                if not numeric:
                    return text
                if len(text) > len(best):
                    best = text
    return best


def _header_matches(text: str, aliases: list[str], threshold: float = 0.72) -> bool:
    candidate = _norm_name(text)
    if not candidate:
        return False
    for alias in aliases:
        target = _norm_name(alias)
        if candidate == target or (len(target) >= 4 and target in candidate):
            return True
        if len(candidate) >= max(3, len(target) - 2) and _name_similarity(candidate, target) >= threshold:
            return True
    return False


def _read_table_by_grid(img: Image.Image) -> dict:
    """Lê a tabela célula a célula depois de detectar sua grade dinamicamente."""
    grid = _detect_table_grid(img)
    if not grid.get("ok"):
        return {"ok": False, "reason": grid.get("reason", "Grade não detectada."), "grid": grid}

    xs = grid["xs"]
    ys = grid["ys"]
    matrix = []
    for r in range(len(ys) - 1):
        row = []
        for c in range(len(xs) - 1):
            box = (xs[c], ys[r], xs[c + 1], ys[r + 1])
            row.append({"box": box, "text": _ocr_cell_text(img, box, numeric=False)})
        matrix.append(row)

    header_index = None
    header_map = {}
    for r, row in enumerate(matrix):
        local = {}
        for c, cell in enumerate(row):
            text = cell["text"]
            if _header_matches(text, ["rake"], threshold=0.78):
                local["rake"] = c
            if _header_matches(text, ["ganhos", "ganho"], threshold=0.72):
                local["ganhos"] = c
            if _header_matches(text, ["super agente", "agente"], threshold=0.68):
                local["agente"] = c
        if "rake" in local and "ganhos" in local:
            header_index = r
            header_map = local
            break

    if header_index is None:
        return {
            "ok": False,
            "reason": "Cabeçalho RAKE/GANHOS não identificado nas células da grade.",
            "grid": grid,
            "matrix": matrix,
        }

    if "agente" not in header_map:
        header_map["agente"] = 0

    data_rows = []
    for r in range(header_index + 1, len(matrix)):
        row = matrix[r]
        agent_col = header_map["agente"]
        if agent_col >= len(row):
            continue
        agent_text = row[agent_col]["text"].strip()
        norm = _norm_name(agent_text)
        if not norm or any(key in norm for key in ["total", "adiantamento"]):
            continue

        rake_cell = row[header_map["rake"]]
        ganhos_cell = row[header_map["ganhos"]]
        rake_text = _ocr_cell_text(img, rake_cell["box"], numeric=True)
        ganhos_text = _ocr_cell_text(img, ganhos_cell["box"], numeric=True)
        rake_vals = extract_all_money_misto(rake_text)
        ganhos_vals = extract_all_money_misto(ganhos_text)
        data_rows.append({
            "row_index": r,
            "agente": agent_text,
            "rake": rake_vals[0] if rake_vals else 0.0,
            "ganhos": ganhos_vals[0] if ganhos_vals else 0.0,
            "rake_text": rake_text,
            "ganhos_text": ganhos_text,
            "rake_box": rake_cell["box"],
            "ganhos_box": ganhos_cell["box"],
        })

    return {
        "ok": True,
        "grid": grid,
        "header_index": header_index,
        "header_map": header_map,
        "rows": data_rows,
    }


def _find_agent_in_grid(table: dict, aliases: list[str]) -> tuple[dict | None, float]:
    targets = [_norm_name(alias) for alias in aliases if _norm_name(alias)]
    best = None
    best_score = 0.0
    for row in table.get("rows", []):
        score = max((_name_similarity(row["agente"], target) for target in targets), default=0.0)
        if score > best_score:
            best = row
            best_score = score
    if best_score < 0.45:
        return None, best_score
    return best, best_score

def _extract_agent_by_headers(img: Image.Image, aliases, display_name: str) -> dict:
    """Lê RAKE e GANHOS da linha do agente sem usar coordenadas fixas.

    A imagem inteira é analisada. Os cabeçalhos definem as posições das colunas,
    e o nome do agente define a linha, portanto a ordem das linhas e a largura das
    colunas podem mudar.
    """
    structure = _detect_table_structure(img)
    if not structure.get("ok"):
        return {
            "found": False,
            "agente": display_name,
            "ganhos": 0.0,
            "rake": 0.0,
            "ocr_text": structure.get("reason", "Estrutura da tabela não identificada."),
        }

    row, score = _find_agent_row(structure, list(aliases))
    if row is None:
        visible_rows = "\n".join(f"- {r['text']}" for r in structure.get("data_rows", []))
        return {
            "found": False,
            "agente": display_name,
            "ganhos": 0.0,
            "rake": 0.0,
            "ocr_text": f"Agente não localizado. Melhor similaridade: {score:.2f}\n\nLinhas detectadas:\n{visible_rows}",
        }

    y1, y2 = _row_vertical_bounds(structure, row, img.height)
    rake_left, rake_right = structure["bounds"]["rake"]
    ganhos_left, ganhos_right = structure["bounds"]["ganhos"]

    rake_box = (int(rake_left), y1, int(rake_right), y2)
    ganhos_box = (int(ganhos_left), y1, int(ganhos_right), y2)

    rake_txt, rake = _numeric_ocr_from_cell(img, rake_box)
    ganhos_txt, ganhos = _numeric_ocr_from_cell(img, ganhos_box)

    centers_txt = ", ".join(f"{k}={v:.1f}" for k, v in sorted(structure["centers"].items(), key=lambda item: item[1]))
    return {
        "found": True,
        "agente": display_name,
        "ganhos": ganhos,
        "rake": rake,
        "ocr_text": (
            f"Linha localizada: {row['text']}\n"
            f"Similaridade: {score:.2f}\n"
            f"Y da linha: {row['y']:.1f}\n"
            f"Cabeçalhos/centros detectados: {centers_txt}\n"
            f"Recorte RAKE: {rake_box}\n"
            f"OCR RAKE:\n{rake_txt}\n"
            f"RAKE FINAL: {rake}\n\n"
            f"Recorte GANHOS: {ganhos_box}\n"
            f"OCR GANHOS:\n{ganhos_txt}\n"
            f"GANHOS FINAL: {ganhos}"
        ),
    }


def extract_agent_from_adamantium_table(img: Image.Image, aliases, display_name: str) -> dict:
    """Lê a imagem inteira, detecta a grade e identifica a linha pelo nome do agente.

    Primeiro tenta a leitura estrutural da tabela por linhas/colunas detectadas.
    Se a grade não for reconhecida, usa o OCR por cabeçalhos como fallback.
    """
    table = _read_table_by_grid(img)
    if table.get("ok"):
        row, score = _find_agent_in_grid(table, list(aliases))
        if row is not None:
            detected_rows = "\n".join(
                f"- {item['agente']}: RAKE={item['rake']} | GANHOS={item['ganhos']}"
                for item in table.get("rows", [])
            )
            return {
                "found": True,
                "agente": display_name,
                "ganhos": float(row["ganhos"]),
                "rake": float(row["rake"]),
                "ocr_text": (
                    "MÉTODO: leitura dinâmica pela grade da tabela inteira\n"
                    f"Linhas verticais detectadas: {table['grid']['xs']}\n"
                    f"Linhas horizontais detectadas: {table['grid']['ys']}\n"
                    f"Colunas identificadas: {table['header_map']}\n"
                    f"Agente reconhecido: {row['agente']}\n"
                    f"Similaridade: {score:.2f}\n"
                    f"Recorte RAKE: {row['rake_box']}\n"
                    f"OCR RAKE: {row['rake_text']}\n"
                    f"RAKE FINAL: {row['rake']}\n"
                    f"Recorte GANHOS: {row['ganhos_box']}\n"
                    f"OCR GANHOS: {row['ganhos_text']}\n"
                    f"GANHOS FINAL: {row['ganhos']}\n\n"
                    f"Linhas reconhecidas:\n{detected_rows}"
                ),
            }

    fallback = _extract_agent_by_headers(img, aliases, display_name)
    grid_reason = table.get("reason", "Agente não encontrado na grade.")
    fallback["ocr_text"] = (
        f"A leitura pela grade não foi concluída: {grid_reason}\n"
        "Foi usado o fallback por cabeçalhos e coordenadas detectadas.\n\n"
        + fallback.get("ocr_text", "")
    )
    return fallback


def extract_suprema_total(img: Image.Image) -> dict:
    """Extrai somente o campo TOTAL da imagem Suprema."""
    text = "\n".join([ocr_image(img, psm=6), ocr_image(img, psm=11)])
    patterns = [
        r"TOTAL[^0-9\-]*R?\$?\s*([\-−–—]?\s*[0-9][0-9\.,]*)",
        r"TOTAL[^0-9\-]*([\-−–—]?[0-9][0-9\.,]*)",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            return {"total": parse_money(m.group(1)), "ocr_text": text}

    # fallback: recorte da região inferior-central onde aparece TOTAL
    w, h = img.size
    crop = img.crop((int(w * 0.32), int(h * 0.55), int(w * 0.72), int(h * 0.82))).resize((int(w * 1.6), int(h * 1.08)))
    crop_text = "\n".join([ocr_image(crop, psm=6), ocr_image(crop, psm=11)])
    vals = extract_all_money_misto(crop_text)
    total = vals[-1] if vals else 0.0
    return {"total": total, "ocr_text": text + "\n\nOCR RECORTE TOTAL:\n" + crop_text}


# =========================
# CÁLCULOS
# =========================
def rb_value(rake: float, rb_percentual: float) -> float:
    return float(rake) * float(rb_percentual) / 100.0


def base_row(agent: str, ganhos: float, rake: float, rb_percentual: float, total_override=None, rebate=0.0):
    rb = rb_value(rake, rb_percentual)
    total = ganhos + rb if total_override is None else float(total_override)
    return {
        "AGENTE": agent,
        "GANHOS": float(ganhos),
        "RAKE": float(rake),
        "RB(%)": f"{int(rb_percentual) if float(rb_percentual).is_integer() else rb_percentual}%",
        "RB": rb,
        "REBATE": float(rebate),
        "TOTAL": total,
    }


# =========================
# RELATÓRIOS
# =========================
def generate_summary_report(titulo: str, periodo: str, headers, values, rebate: float, total_final: float) -> Image.Image:
    W, H = SUMMARY_W, SUMMARY_H
    img = Image.new("RGB", (W, H), GRAY)
    draw = ImageDraw.Draw(img)
    title_font = get_font(FONT_TITLE, bold=True)
    subtitle_font = get_font(FONT_SUBTITLE, bold=False)
    subtitle_bold = get_font(FONT_SUBTITLE, bold=True)
    header_font = get_font(FONT_HEADER, bold=True)
    value_font = get_font(FONT_CELL, bold=False)
    total_font = get_font(FONT_TOTAL, bold=True)
    status_font = get_font(FONT_STATUS, bold=True)
    status_value_font = get_font(FONT_STATUS_VALUE, bold=True)

    tw, _ = measure(draw, titulo, title_font)
    draw.text(((W - tw) / 2, 40), titulo, fill=NAVY, font=title_font)
    label = "Período do fechamento:"
    l_w, _ = measure(draw, label, subtitle_font)
    p_w, _ = measure(draw, periodo, subtitle_bold)
    x0 = (W - (l_w + 12 + p_w)) / 2
    draw.text((x0, 150), label, fill=NAVY, font=subtitle_font)
    draw.text((x0 + l_w + 12, 150), periodo, fill=NAVY, font=subtitle_bold)

    x1, x2, y1, y2 = 60, 1340, 250, 455
    draw.rounded_rectangle((x1, y1, x2, y2), radius=14, outline=NAVY, width=3, fill=WHITE)
    draw.rectangle((x1, y1, x2, 325), fill=NAVY)
    col_w = (x2-x1)/len(headers)
    for i, htxt in enumerate(headers):
        cx=x1+(i+.5)*col_w; sw,_=measure(draw,htxt,header_font)
        draw.text((cx-sw/2,272),htxt,fill=WHITE,font=header_font)
        val=fmt_brl(values[i]); vw,_=measure(draw,val,value_font)
        draw.text((cx-vw/2,365),val,fill=NAVY,font=value_font)
    draw.rectangle((60,550,1340,635),fill=YELLOW)
    val=fmt_brl(total_final); vw,_=measure(draw,val,total_font)
    draw.text((980,575),"TOTAL",fill=(0,0,0),font=total_font)
    draw.text((1310-vw,575),val,fill=(0,0,0),font=total_font)
    status_text = "PREMIER TEM A PAGAR" if total_final > 0 else ("PREMIER TEM A RECEBER" if total_final < 0 else "SEM VALORES")
    status_value=f"R$ {fmt_brl(abs(total_final))}"
    draw.rounded_rectangle((60,710,1340,835),radius=16,outline=NAVY,width=3,fill=LIGHT_BG)
    stw,_=measure(draw,status_text,status_font); svw,_=measure(draw,status_value,status_value_font)
    sx=(W-(stw+24+svw))/2
    draw.text((sx,748),status_text,fill=NAVY,font=status_font)
    draw.text((sx+stw+24,744),status_value,fill=GREEN if total_final>=0 else RED,font=status_value_font)
    return img




def generate_client_table_image(titulo: str, periodo: str, df: pd.DataFrame, total_geral: float,
                                adjustment_rows=None, total_base_exibido: float | None = None) -> Image.Image:
    """Gera tabela com colunas variáveis e linhas-resumo ao final."""
    adjustment_rows = adjustment_rows or []
    columns = list(df.columns)
    n = len(columns)
    W = TABLE_W
    x1 = 40
    usable = W - 80
    agent_w = 390 if n <= 6 else 300
    remaining = usable - agent_w
    other_w = remaining / max(1, n-1)
    widths = [agent_w] + [other_w]*(n-1)
    xs=[x1]
    for w in widths: xs.append(xs[-1]+w)

    temp=Image.new("RGB",(10,10),WHITE); td=ImageDraw.Draw(temp)
    agent_font=get_font(22); line_h=measure(td,"Ag",agent_font)[1]+4
    row_heights=[]
    for _,row in df.iterrows():
        lines=wrap_text(td,str(row[columns[0]]),agent_font,widths[0]-20)
        row_heights.append(max(TABLE_ROW_H_MIN,len(lines)*line_h+16))

    summary_count = 1 + len(adjustment_rows)
    H = TABLE_TOP + TABLE_HEADER_H + sum(row_heights) + summary_count*TABLE_ROW_H_MIN + 180
    img=Image.new("RGB",(W,H),GRAY); draw=ImageDraw.Draw(img)
    title_font=get_font(FONT_TITLE,True); subtitle_font=get_font(FONT_SUBTITLE); subtitle_bold=get_font(FONT_SUBTITLE,True)
    header_font=get_font(24,True); cell_font=get_font(24); total_font=get_font(FONT_TOTAL,True)
    status_font=get_font(FONT_STATUS,True); status_value_font=get_font(FONT_STATUS_VALUE,True)

    tw,_=measure(draw,titulo,title_font); draw.text(((W-tw)/2,20),titulo,fill=NAVY,font=title_font)
    label="Período do fechamento:"; lw,_=measure(draw,label,subtitle_font); pw,_=measure(draw,periodo,subtitle_bold)
    sx=(W-(lw+12+pw))/2; draw.text((sx,100),label,fill=NAVY,font=subtitle_font); draw.text((sx+lw+12,100),periodo,fill=NAVY,font=subtitle_bold)

    draw.rectangle((x1,TABLE_TOP,xs[-1],TABLE_TOP+TABLE_HEADER_H),fill=NAVY)
    for i,c in enumerate(columns):
        txt=str(c); sw,_=measure(draw,txt,header_font); cx=(xs[i]+xs[i+1])/2
        draw.text((cx-sw/2,TABLE_TOP+13),txt,fill=WHITE,font=header_font)

    y=TABLE_TOP+TABLE_HEADER_H
    for pos,(_,row) in enumerate(df.iterrows()):
        rh=row_heights[pos]; draw.rectangle((x1,y,xs[-1],y+rh),fill=WHITE,outline=GRID,width=2)
        wrapped=wrap_text(draw,str(row[columns[0]]),agent_font,widths[0]-20)
        for j,line in enumerate(wrapped): draw.text((xs[0]+10,y+8+j*line_h),line,fill=(0,0,0),font=agent_font)
        for i,c in enumerate(columns[1:],start=1):
            v=row[c]
            txt=str(v) if isinstance(v,str) else fmt_brl(v)
            sw,sh=measure(draw,txt,cell_font); cx=(xs[i]+xs[i+1])/2
            draw.text((cx-sw/2,y+(rh-sh)/2-2),txt,fill=(0,0,0),font=cell_font)
        y+=rh

    total_base = total_geral if total_base_exibido is None else total_base_exibido
    summary = [("SUBTOTAL" if adjustment_rows else "TOTAL", total_base, NAVY, WHITE)] + adjustment_rows
    for idx,(label,value,bg,label_color) in enumerate(summary):
            draw.rectangle((x1,y,xs[-1],y+TABLE_ROW_H_MIN),fill=bg,outline=GRID,width=2)
            label_end=xs[-2]
            sw,_=measure(draw,label,total_font)
            draw.text((label_end-sw-14,y+10),label,fill=label_color,font=total_font)

            val=fmt_brl(value)
            vw,_=measure(draw,val,total_font)

            # valor branco apenas na linha SUBTOTAL/TOTAL azul
            value_color = WHITE if idx == 0 else (0,0,0)

            draw.text(
                ((xs[-2]+xs[-1])/2-vw/2, y+10),
                val,
                fill=value_color,
                font=total_font
            )

            y+=TABLE_ROW_H_MIN

    status_text="PREMIER TEM A PAGAR" if total_geral>0 else ("PREMIER TEM A RECEBER" if total_geral<0 else "SEM VALORES")
    status_value=f"R$ {fmt_brl(abs(total_geral))}"; by1=y+35; by2=by1+90
    draw.rounded_rectangle((40,by1,xs[-1],by2),radius=14,outline=NAVY,width=3,fill=LIGHT_BG)
    sw,_=measure(draw,status_text,status_font); vw,_=measure(draw,status_value,status_value_font); start=(W-(sw+20+vw))/2
    draw.text((start,by1+28),status_text,fill=NAVY,font=status_font)
    draw.text((start+sw+20,by1+24),status_value,fill=GREEN if total_geral>=0 else RED,font=status_value_font)
    return img


# =========================
# PÁGINAS
# =========================


def page_alex():
    st.subheader("Alex")
    periodo=st.text_input("Período do fechamento",key="periodo_alex",placeholder="06/04/2026 a 12/04/2026")
    pdf=st.file_uploader("Envie o PDF do Alex",type=["pdf"],key="alex_pdf")
    rows=[]
    if pdf is not None:
        for _,r in process_pdf_by_client(pdf,"Alex").iterrows():
            rows.append(base_row(r["agente"],r["ganhos"],r["rake"],r["rb_percentual"]))
    if st.button("Gerar fechamento Alex",type="primary",key="btn_alex"):
        if not rows: st.warning("Envie o PDF do Alex."); return
        df=pd.DataFrame(rows)[["AGENTE","GANHOS","RAKE","RB(%)","RB","TOTAL"]]
        total=float(df["TOTAL"].sum())
        report=generate_client_table_image("ALEX",periodo.strip() or "-",df,total)
        st.image(report,caption="Pronto para print",use_container_width=True)
        st.download_button("Baixar relatório em PNG",data=to_png_bytes(report),file_name="alex_fechamento.png",mime="image/png")


def page_oscar():
    st.subheader("Oscar")
    periodo=st.text_input("Período do fechamento",key="periodo_oscar",placeholder="06/04/2026 a 12/04/2026")
    pdf=st.file_uploader("Envie o PDF",type=["pdf"],key="oscar_pdf")
    rows=[]
    if pdf is not None:
        for _,r in process_pdf_by_client(pdf,"Oscar").iterrows():
            row=base_row(r["agente"],r["ganhos"],r["rake"],r["rb_percentual"],rebate=r["rebate"])
            row["TOTAL"]=row["GANHOS"]+row["RB"]+row["REBATE"]
            rows.append(row)
    if st.button("Gerar fechamento Oscar",type="primary",key="btn_oscar"):
        if not rows: st.warning("Envie o PDF do Oscar."); return
        df=pd.DataFrame(rows)[["AGENTE","GANHOS","RAKE","RB(%)","RB","REBATE","TOTAL"]]
        total=float(df["TOTAL"].sum())
        report=generate_client_table_image("OSCAR",periodo.strip() or "-",df,total)
        st.image(report,caption="Pronto para print",use_container_width=True)
        st.download_button("Baixar relatório em PNG",data=to_png_bytes(report),file_name="oscar_fechamento.png",mime="image/png")


def page_demetra():
    st.subheader("Demetra")
    periodo=st.text_input("Período do fechamento",key="periodo_demetra",placeholder="06/04/2026 a 12/04/2026")
    pdf=st.file_uploader("Envie o PDF",type=["pdf"],key="demetra_pdf")
    imagem=st.file_uploader("Envie a imagem Killuminatti",type=["png","jpg","jpeg","webp"],key="demetra_killuminatti_img")
    rows=[]
    if pdf is not None:
        for _,r in process_pdf_by_client(pdf,"Demetra").iterrows():
            rows.append(base_row(r["agente"],r["ganhos"],r["rake"],r["rb_percentual"]))
    if imagem is not None:
        img=Image.open(imagem)
        dados=extract_agent_from_adamantium_table(img,["Killuminatti","KILLUMINATTI"],"Killuminatti")
        with st.expander("Diagnóstico OCR - Killuminatti",expanded=False): st.code(dados["ocr_text"])
        if not dados["found"] or (dados["ganhos"]==0 and dados["rake"]==0):
            st.warning("Não consegui localizar ou ler a linha Killuminatti com segurança.")
        else:
            rows.append(base_row("Killuminatti",dados["ganhos"],dados["rake"],70.0))
    if st.button("Gerar fechamento Demetra",type="primary",key="btn_demetra"):
        if not rows: st.warning("Envie o PDF e/ou a imagem Killuminatti."); return
        df=pd.DataFrame(rows)[["AGENTE","GANHOS","RAKE","RB(%)","RB","TOTAL"]]
        subtotal=float(df["TOTAL"].sum())
        rebate=subtotal*(REBATE_DEMETRA/100.0) if subtotal>0 else 0.0
        total=subtotal+rebate
        adjustments=[("-5% total",rebate,LIGHT_GRAY,(0,0,0)),("TOTAL",total,YELLOW,(0,0,0))]
        report=generate_client_table_image("DEMETRA",periodo.strip() or "-",df,total,adjustments,subtotal)
        st.image(report,caption="Pronto para print",use_container_width=True)
        st.download_button("Baixar relatório em PNG",data=to_png_bytes(report),file_name="demetra_fechamento.png",mime="image/png")


def page_strong():
    st.subheader("Strong")
    periodo=st.text_input("Período do fechamento",key="periodo_strong",placeholder="06/04/2026 a 12/04/2026")
    suprema=st.file_uploader("Envie a imagem SUPREMA",type=["png","jpg","jpeg","webp"],key="strong_suprema_img")
    adamantium=st.file_uploader("Envie a imagem ADAMANTIUM",type=["png","jpg","jpeg","webp"],key="strong_adamantium_img")
    rows=[]; total_suprema=0.0
    if suprema is not None:
        s=extract_suprema_total(Image.open(suprema)); total_suprema=s["total"]
        with st.expander("Diagnóstico OCR - SUPREMA",expanded=False): st.code(s["ocr_text"])
    if adamantium is not None:
        img=Image.open(adamantium)
        p=extract_agent_from_adamantium_table(img,["PPFICHAS","PP FICHAS"],"PPFICHAS")
        m=extract_agent_from_adamantium_table(img,["MrLeo79","MrLeo","MRLEO79"],"MrLeo79")
        with st.expander("Diagnóstico OCR - ADAMANTIUM",expanded=False):
            st.code("PPFICHAS\n"+p["ocr_text"]+"\n\nMrLeo79\n"+m["ocr_text"])
        if p["found"] and (abs(p["ganhos"])>0.0001 or abs(p["rake"])>0.0001):
            rows.append(base_row("PPFICHAS",p["ganhos"],p["rake"],70.0))
        if m["found"] and (abs(m["ganhos"])>0.0001 or abs(m["rake"])>0.0001):
            rows.append(base_row("MrLeo79",m["ganhos"],m["rake"],70.0,total_override=rb_value(m["rake"],70.0)))
    if st.button("Gerar fechamento Strong",type="primary",key="btn_strong"):
        if adamantium is None and suprema is None: st.warning("Envie as imagens SUPREMA e ADAMANTIUM."); return
        if not rows and total_suprema==0: st.warning("Nenhum valor diferente de zero foi identificado."); return
        df=pd.DataFrame(rows,columns=["AGENTE","GANHOS","RAKE","RB(%)","RB","REBATE","TOTAL"])
        if "REBATE" in df.columns: df=df.drop(columns=["REBATE"])
        df=df[["AGENTE","GANHOS","RAKE","RB(%)","RB","TOTAL"]]
        subtotal=float(df["TOTAL"].sum()) if not df.empty else 0.0
        total=subtotal+total_suprema
        adjustments=[("Total SUPREMA",total_suprema,LIGHT_GRAY,(0,0,0)),("TOTAL",total,YELLOW,(0,0,0))]
        report=generate_client_table_image("Adamantium PPPoker - STRONG",periodo.strip() or "-",df,total,adjustments,subtotal)
        st.image(report,caption="Pronto para print",use_container_width=True)
        st.download_button("Baixar relatório em PNG",data=to_png_bytes(report),file_name="strong_fechamento.png",mime="image/png")


st.title("Fechamentos Premier")
cliente = st.selectbox("Escolha o cliente", ["Demetra", "Oscar", "Alex", "Strong"])
if cliente == "Demetra":
    page_demetra()
elif cliente == "Oscar":
    page_oscar()
elif cliente == "Alex":
    page_alex()
else:
    page_strong()
