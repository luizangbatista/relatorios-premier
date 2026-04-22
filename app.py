# ============================================================
# FECHAMENTOS PREMIER - APP FINAL CORRIGIDO DEFINITIVO
# ------------------------------------------------------------
# Clientes:
# - Harnefer: OCR em imagem
# - Demetra: planilha + PDF
# - Oscar: PDF
# - Alex: PDF + imagem Casarica (2 formatos)
#
# CORREÇÃO DEFINITIVA DO CASARICA NOVO:
# - leitura por recorte obrigatório dos blocos "Taxa" e "Ganhos"
# - ignora ID, %, "retorno de taxa" e outros números da tela
# - só usa fallback textual no formato antigo (profit/loss)
# ============================================================

import io
import re
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps

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
RB_HARNEFER = 76.0
RB_DEMETRA_PLANILHA = 70.0
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
    "12970177": {"cliente": "Alex", "rb": 70.0},
    "13019559": {"cliente": "Alex", "rb": 50.0},
    "13018880": {"cliente": "Alex", "rb": 45.0},
    "13213751": {"cliente": "Alex", "rb": 70.0},
    "13265647": {"cliente": "Alex", "rb": 50.0},
    "13319248": {"cliente": "Alex", "rb": 70.0},
    "13379845": {"cliente": "Alex", "rb": 60.0},
    "13104440": {"cliente": "Alex", "rb": 50.0},
    "13472941": {"cliente": "Oscar", "rb": 65.0},
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

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")

    if text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]

    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return 0.0

    try:
        return float(text)
    except Exception:
        return 0.0


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
    # escolhe primeiro valor plausível
    for v in vals:
        if abs(v) < 100000:
            return txt, v
    return txt, 0.0


# =========================
# HARNEFER
# =========================
def detect_harnefer_image(img: Image.Image) -> bool:
    text = (ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)).upper()
    return ("TOTAL FEE" in text and "WINNINGS" in text) or ("GAMES" in text and "ADMIN FEE" in text)


def crop_harnefer_summary(img: Image.Image) -> Image.Image:
    w, h = img.size
    return img.crop((int(w * 0.08), int(h * 0.32), int(w * 0.93), int(h * 0.64)))


def extract_harnefer_values(img: Image.Image) -> dict:
    crop = crop_harnefer_summary(img)
    w, h = crop.size
    col_w = w // 4
    cols = [crop.crop((i * col_w, 0, (i + 1) * col_w if i < 3 else w, h)) for i in range(4)]
    texts = [ocr_image(c, psm=6) + "\n" + ocr_image(c, psm=11) for c in cols]
    rake = first_money(texts[1])
    ganhos = first_money(texts[3])
    rakeback = rake * (RB_HARNEFER / 100.0)
    total = ganhos + rakeback
    return {
        "ganhos": ganhos,
        "rake": rake,
        "rakeback": rakeback,
        "total_final": total,
        "ocr_fee": texts[1],
        "ocr_winnings": texts[3],
        "crop": crop,
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

    # FORMATO ANTIGO - ainda pode usar texto
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

    # FORMATO NOVO - recorte obrigatório, sem fallback perigoso
    w, h = img.size

    # Recortes mais estreitos para fugir do ID, 70% e campos inferiores
    taxa_box = (
        int(w * 0.06),
        int(h * 0.48),
        int(w * 0.34),
        int(h * 0.68),
    )
    ganhos_box = (
        int(w * 0.66),
        int(h * 0.48),
        int(w * 0.94),
        int(h * 0.68),
    )

    taxa_txt, rake = ocr_crop_value(img, taxa_box)
    ganhos_txt, ganhos = ocr_crop_value(img, ganhos_box)

    # Validação forte: não aceitar valores absurdos
    if abs(rake) > 50000:
        rake = 0.0
    if abs(ganhos) > 50000:
        ganhos = 0.0

    # Se ambos falharem, não tenta leitura global perigosa
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
    rows = []
    for line in extract_pdf_lines(uploaded_file):
        if "R$" not in line:
            continue
        id_match = re.search(r"\b(\d{6,9})\b", line)
        if not id_match:
            continue
        id_agente = normalize_id(id_match.group(1))
        money_matches = re.findall(r"R\$\s*-?\d[\d\.,]*", line)
        if len(money_matches) < 2:
            continue
        info = MAPA_IDS_PDF.get(id_agente)
        if not info or info["cliente"] != cliente_alvo:
            continue
        ganhos = parse_money(money_matches[0])
        rake = parse_money(money_matches[1])
        agente = re.sub(r"\s{2,}", " ", line.split(id_agente)[0].strip())
        rows.append({"agente": agente, "ganhos": ganhos, "rake": rake, "rb_percentual": float(info["rb"])})
    return pd.DataFrame(rows)


# =========================
# CÁLCULOS
# =========================
def calc_row(ganhos: float, rake: float, rb_percentual: float, rebate_percentual: float):
    rb_valor = rake * (rb_percentual / 100.0)
    total_base = ganhos + rb_valor
    rebate = total_base * (rebate_percentual / 100.0) if total_base > 0 else 0.0
    total_final = total_base + rebate
    return total_base, rebate, total_final


def calc_alex_row(ganhos: float, rake: float, rb_percentual: float):
    rb_valor = rake * (rb_percentual / 100.0)
    total_base = ganhos + rb_valor
    if total_base > 0:
        rebate = total_base * (REBATE_ALEX_POSITIVO / 100.0)
    elif total_base < 0:
        rebate = total_base * (REBATE_ALEX_NEGATIVO / 100.0)
    else:
        rebate = 0.0
    total_final = total_base + rebate
    return total_base, rebate, total_final


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
    small_font = get_font(FONT_CELL, bold=False)
    total_font = get_font(FONT_TOTAL, bold=True)
    status_font = get_font(FONT_STATUS, bold=True)
    status_value_font = get_font(FONT_STATUS_VALUE, bold=True)

    tw, _ = measure(draw, titulo, title_font)
    draw.text(((W - tw) / 2, 40), titulo, fill=NAVY, font=title_font)
    draw.line((70, 115, 300, 115), fill=GOLD, width=3)
    draw.line((1100, 115, 1330, 115), fill=GOLD, width=3)

    label = "Período do fechamento:"
    l_w, _ = measure(draw, label, subtitle_font)
    p_w, _ = measure(draw, periodo, subtitle_bold)
    x0 = (W - (l_w + 12 + p_w)) / 2
    draw.text((x0, 150), label, fill=NAVY, font=subtitle_font)
    draw.text((x0 + l_w + 12, 150), periodo, fill=NAVY, font=subtitle_bold)

    table_x1, table_x2 = 60, 1340
    header_y1, header_y2 = 250, 325
    values_y1, values_y2 = 325, 455
    draw.rounded_rectangle((table_x1, header_y1, table_x2, values_y2), radius=14, outline=NAVY, width=3, fill="white")
    draw.rounded_rectangle((table_x1, header_y1, table_x2, header_y2), radius=14, fill=NAVY)
    draw.rectangle((table_x1, header_y2 - 14, table_x2, header_y2), fill=NAVY)

    col_w = (table_x2 - table_x1) / 4
    for i in range(1, 4):
        x = table_x1 + i * col_w
        draw.line((x, header_y1, x, values_y2), fill=(160, 160, 160), width=2)

    for i, col in enumerate(headers):
        cx = table_x1 + i * col_w + col_w / 2
        cw, _ = measure(draw, col, header_font)
        draw.text((cx - cw / 2, header_y1 + 22), col, fill=WHITE, font=header_font)

    for i, val in enumerate(values):
        sval = fmt_brl(val)
        cx = table_x1 + i * col_w + col_w / 2
        vw, _ = measure(draw, sval, value_font)
        draw.text((cx - vw / 2, values_y1 + 45), sval, fill=NAVY, font=value_font)

    draw.rectangle((60, 490, 1340, 550), fill=LIGHT_GRAY)
    rebate_label = "-5% total" if rebate != 0 else "Sem rebate"
    rebate_value = fmt_brl(rebate) if rebate != 0 else "0,00"
    rl_w, _ = measure(draw, rebate_label, small_font)
    rv_w, _ = measure(draw, rebate_value, small_font)
    draw.text((1000 - rl_w, 506), rebate_label, fill=NAVY, font=small_font)
    draw.text((1310 - rv_w, 506), rebate_value, fill=NAVY, font=small_font)

    draw.rectangle((60, 550, 1340, 635), fill=YELLOW)
    tlabel, tvalue = "TOTAL", fmt_brl(total_final)
    tl_w, _ = measure(draw, tlabel, total_font)
    tv_w, _ = measure(draw, tvalue, total_font)
    draw.text((1040 - tl_w, 575), tlabel, fill=(0, 0, 0), font=total_font)
    draw.text((1310 - tv_w, 575), tvalue, fill=(0, 0, 0), font=total_font)

    status_text = "PREMIER TEM A PAGAR" if total_final > 0 else ("PREMIER TEM A RECEBER" if total_final < 0 else "SEM VALORES")
    status_value = f"R$ {fmt_brl(abs(total_final))}"
    draw.rounded_rectangle((60, 710, 1340, 835), radius=16, outline=NAVY, width=3, fill=LIGHT_BG)
    stw, _ = measure(draw, status_text, status_font)
    svw, _ = measure(draw, status_value, status_value_font)
    start_x = (W - (stw + 24 + svw)) / 2
    draw.text((start_x, 748), status_text, fill=NAVY, font=status_font)
    draw.text((start_x + stw + 24, 744), status_value, fill=GREEN if total_final >= 0 else RED, font=status_value_font)
    return img


def generate_harnefer_report(periodo: str, ganhos: float, rake: float) -> Image.Image:
    return generate_summary_report(
        "HARNEFER",
        periodo,
        ["GANHOS", "RAKE", f"RB ({int(RB_HARNEFER)}%)", "TOTAL"],
        [ganhos, rake, rake * (RB_HARNEFER / 100.0), ganhos + rake * (RB_HARNEFER / 100.0)],
        0.0,
        ganhos + rake * (RB_HARNEFER / 100.0),
    )


def generate_client_table_image(titulo: str, periodo: str, df: pd.DataFrame, total_geral: float, rebate_total: float, rebate_label: str) -> Image.Image:
    temp_img = Image.new("RGB", (10, 10), WHITE)
    temp_draw = ImageDraw.Draw(temp_img)
    cell_font = get_font(FONT_CELL, bold=False)
    agent_font = get_font(FONT_AGENT, bold=False)
    line_h = measure(temp_draw, "Ag", agent_font)[1] + 4
    agente_col_width = 500 - 24

    row_heights = []
    for _, row in df.iterrows():
        lines = wrap_text(temp_draw, str(row["AGENTE"]), agent_font, agente_col_width)
        row_heights.append(max(TABLE_ROW_H_MIN, len(lines) * line_h + 16))

    H = TABLE_TOP + TABLE_HEADER_H + sum(row_heights) + (3 * TABLE_ROW_H_MIN) + 180
    W = TABLE_W
    img = Image.new("RGB", (W, H), GRAY)
    draw = ImageDraw.Draw(img)

    title_font = get_font(FONT_TITLE, bold=True)
    subtitle_font = get_font(FONT_SUBTITLE, bold=False)
    subtitle_bold = get_font(FONT_SUBTITLE, bold=True)
    header_font = get_font(FONT_HEADER, bold=True)
    cell_font = get_font(FONT_CELL, bold=False)
    agent_font = get_font(FONT_AGENT, bold=False)
    total_font = get_font(FONT_TOTAL, bold=True)
    status_font = get_font(FONT_STATUS, bold=True)
    status_value_font = get_font(FONT_STATUS_VALUE, bold=True)

    tw, _ = measure(draw, titulo, title_font)
    draw.text(((W - tw) / 2, 20), titulo, fill=NAVY, font=title_font)
    draw.line((60, 70, 250, 70), fill=GOLD, width=3)
    draw.line((1190, 70, 1380, 70), fill=GOLD, width=3)

    label = "Período do fechamento:"
    l_w, _ = measure(draw, label, subtitle_font)
    p_w, _ = measure(draw, periodo, subtitle_bold)
    x0 = (W - (l_w + 12 + p_w)) / 2
    draw.text((x0, 100), label, fill=NAVY, font=subtitle_font)
    draw.text((x0 + l_w + 12, 100), periodo, fill=NAVY, font=subtitle_bold)

    x1 = 50
    widths = [500, 210, 210, 130, 260]
    headers = ["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]
    xs = [x1]
    for w in widths:
        xs.append(xs[-1] + w)
    table_w = sum(widths)

    draw.rectangle((x1, TABLE_TOP, x1 + table_w, TABLE_TOP + TABLE_HEADER_H), fill=NAVY)
    for i, htxt in enumerate(headers):
        cx = (xs[i] + xs[i + 1]) / 2
        cw, _ = measure(draw, htxt, header_font)
        draw.text((cx - cw / 2, TABLE_TOP + 12), htxt, fill=WHITE, font=header_font)

    y = TABLE_TOP + TABLE_HEADER_H
    for idx, row in df.iterrows():
        row_h = row_heights[idx]
        draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=WHITE, outline=GRID, width=2)

        wrapped = wrap_text(draw, str(row["AGENTE"]), agent_font, agente_col_width)
        for j, line in enumerate(wrapped):
            draw.text((xs[0] + 12, y + 8 + j * line_h), line, fill=(0, 0, 0), font=agent_font)

        vals = [fmt_brl(row["GANHOS"]), fmt_brl(row["RAKE"]), str(row["RB"]), fmt_brl(row["TOTAL"])]
        for i, val in enumerate(vals, start=1):
            vw, vh = measure(draw, val, cell_font)
            cx = (xs[i] + xs[i + 1]) / 2
            cy = y + (row_h - vh) / 2 - 2
            draw.text((cx - vw / 2, cy), val, fill=(0, 0, 0), font=cell_font)

        y += row_h

    draw.rectangle((x1, y, x1 + table_w, y + TABLE_ROW_H_MIN), fill=WHITE, outline=GRID, width=2)
    draw.rectangle((x1, y, xs[4], y + TABLE_ROW_H_MIN), fill=NAVY)
    total_label = "TOTAL"
    tlw, _ = measure(draw, total_label, total_font)
    draw.text((xs[4] - tlw - 16, y + 10), total_label, fill=WHITE, font=total_font)
    total_str = fmt_brl(total_geral)
    tvw, _ = measure(draw, total_str, total_font)
    draw.text((((xs[4] + xs[5]) / 2) - tvw / 2, y + 10), total_str, fill=(0, 0, 0), font=total_font)
    y += TABLE_ROW_H_MIN

    draw.rectangle((x1, y, x1 + table_w, y + TABLE_ROW_H_MIN), fill=LIGHT_GRAY, outline=GRID, width=2)
    rlabel = rebate_label if rebate_total != 0 else "Sem rebate"
    rvalue = fmt_brl(rebate_total) if rebate_total != 0 else "0,00"
    rlw, _ = measure(draw, rlabel, cell_font)
    rvw, _ = measure(draw, rvalue, cell_font)
    draw.text((xs[4] - rlw - 16, y + 12), rlabel, fill=(0, 0, 0), font=cell_font)
    draw.text((((xs[4] + xs[5]) / 2) - rvw / 2, y + 12), rvalue, fill=(0, 0, 0), font=cell_font)
    y += TABLE_ROW_H_MIN

    draw.rectangle((x1, y, x1 + table_w, y + TABLE_ROW_H_MIN), fill=YELLOW, outline=GRID, width=2)
    flw, _ = measure(draw, total_label, total_font)
    fvw, _ = measure(draw, total_str, total_font)
    draw.text((xs[4] - flw - 16, y + 10), total_label, fill=(0, 0, 0), font=total_font)
    draw.text((((xs[4] + xs[5]) / 2) - fvw / 2, y + 10), total_str, fill=(0, 0, 0), font=total_font)
    y += TABLE_ROW_H_MIN

    status_text = "PREMIER TEM A PAGAR" if total_geral > 0 else ("PREMIER TEM A RECEBER" if total_geral < 0 else "SEM VALORES")
    status_value = f"R$ {fmt_brl(abs(total_geral))}"
    box_y1 = y + 35
    box_y2 = box_y1 + 90
    draw.rounded_rectangle((50, box_y1, x1 + table_w, box_y2), radius=14, outline=NAVY, width=3, fill=LIGHT_BG)
    stw, _ = measure(draw, status_text, status_font)
    svw, _ = measure(draw, status_value, status_value_font)
    start_x = (W - (stw + 20 + svw)) / 2
    draw.text((start_x, box_y1 + 28), status_text, fill=NAVY, font=status_font)
    draw.text((start_x + stw + 20, box_y1 + 24), status_value, fill=GREEN if total_geral >= 0 else RED, font=status_value_font)
    return img


# =========================
# PÁGINAS
# =========================
def page_harnefer():
    st.subheader("Harnefer")
    periodo = st.text_input("Período do fechamento", key="periodo_harnefer", placeholder="13/04/2026 a 19/04/2026")
    arquivo = st.file_uploader("Envie a imagem do Harnefer", type=["png", "jpg", "jpeg", "webp"], key="harnefer_img")

    if not TESSERACT_OK:
        st.error("OCR indisponível: instale `pytesseract` e `tesseract-ocr`.")
        return

    if arquivo and st.button("Ler imagem e gerar fechamento", type="primary", key="btn_harnefer"):
        img = Image.open(arquivo)
        if not detect_harnefer_image(img):
            st.warning("Não identifiquei a imagem do Harnefer com segurança.")
            return

        dados = extract_harnefer_values(img)

        with st.expander("Diagnóstico OCR", expanded=False):
            st.image(dados["crop"], caption="Recorte usado para leitura", width=700)
            st.code(dados["ocr_fee"])
            st.code(dados["ocr_winnings"])

        report = generate_harnefer_report(periodo.strip() or "-", dados["ganhos"], dados["rake"])
        st.image(report, caption="Relatório final", use_container_width=True)
        st.download_button("Baixar relatório em PNG", data=to_png_bytes(report), file_name="harnefer_fechamento.png", mime="image/png")


def page_demetra():
    st.subheader("Demetra")
    periodo = st.text_input("Período do fechamento", key="periodo_demetra", placeholder="06/04/2026 a 12/04/2026")
    planilha = st.file_uploader("Envie a planilha 2101...", type=["xlsx", "xls"], key="demetra_xlsx")
    pdf = st.file_uploader("Envie o PDF", type=["pdf"], key="demetra_pdf")

    rows = []
    if planilha is not None:
        demetra_df = process_demetra_excel(planilha)
        if not demetra_df.empty:
            ganhos_excel = demetra_df["ganhos"].sum()
            rake_excel = demetra_df["rake"].sum()
            total_base, rebate, total_final = calc_row(ganhos_excel, rake_excel, RB_DEMETRA_PLANILHA, REBATE_DEMETRA)
            rows.append({"AGENTE": AGENTE_PLANILHA_DEMETRA, "GANHOS": ganhos_excel, "RAKE": rake_excel, "RB": f"{int(RB_DEMETRA_PLANILHA)}%", "TOTAL": total_base, "_REBATE": rebate, "_TOTAL_FINAL": total_final})
        else:
            st.info("Nenhuma linha encontrada para TheShark_ com ID 11719117 na planilha.")

    if pdf is not None:
        df_pdf = process_pdf_by_client(pdf, "Demetra")
        for _, row in df_pdf.iterrows():
            total_base, rebate, total_final = calc_row(float(row["ganhos"]), float(row["rake"]), float(row["rb_percentual"]), REBATE_DEMETRA)
            rows.append({"AGENTE": row["agente"], "GANHOS": float(row["ganhos"]), "RAKE": float(row["rake"]), "RB": f"{int(float(row['rb_percentual']))}%", "TOTAL": total_base, "_REBATE": rebate, "_TOTAL_FINAL": total_final})

    if st.button("Gerar fechamento Demetra", type="primary", key="btn_demetra"):
        if not rows:
            st.warning("Envie a planilha e/ou o PDF.")
            return
        detalhado = pd.DataFrame(rows)
        rebate_total = detalhado["_REBATE"].sum()
        total_geral = detalhado["_TOTAL_FINAL"].sum()
        report = generate_client_table_image("DEMETRA", periodo.strip() or "-", detalhado[["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]], total_geral, rebate_total, "-5% total")
        st.image(report, caption="Pronto para print", use_container_width=True)
        st.download_button("Baixar relatório em PNG", data=to_png_bytes(report), file_name="demetra_fechamento.png", mime="image/png")


def page_oscar():
    st.subheader("Oscar")
    periodo = st.text_input("Período do fechamento", key="periodo_oscar", placeholder="06/04/2026 a 12/04/2026")
    pdf = st.file_uploader("Envie o PDF", type=["pdf"], key="oscar_pdf")
    rows = []
    if pdf is not None:
        df_pdf = process_pdf_by_client(pdf, "Oscar")
        for _, row in df_pdf.iterrows():
            total_base, rebate, total_final = calc_row(float(row["ganhos"]), float(row["rake"]), float(row["rb_percentual"]), REBATE_OSCAR)
            rows.append({"AGENTE": row["agente"], "GANHOS": float(row["ganhos"]), "RAKE": float(row["rake"]), "RB": f"{int(float(row['rb_percentual']))}%", "TOTAL": total_base, "_REBATE": rebate, "_TOTAL_FINAL": total_final})
    if st.button("Gerar fechamento Oscar", type="primary", key="btn_oscar"):
        if not rows:
            st.warning("Envie o PDF.")
            return
        დეტ = pd.DataFrame(rows)
        rebate_total = დეტ["_REBATE"].sum()
        total_geral = დეტ["_TOTAL_FINAL"].sum()
        report = generate_client_table_image("OSCAR", periodo.strip() or "-", დეტ[["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]], total_geral, rebate_total, "-10% total")
        st.image(report, caption="Pronto para print", use_container_width=True)
        st.download_button("Baixar relatório em PNG", data=to_png_bytes(report), file_name="oscar_fechamento.png", mime="image/png")


def page_alex():
    st.subheader("Alex")
    periodo = st.text_input("Período do fechamento", key="periodo_alex", placeholder="06/04/2026 a 12/04/2026")
    pdf = st.file_uploader("Envie o PDF do Alex", type=["pdf"], key="alex_pdf")
    imagem = st.file_uploader("Envie a imagem Gustavo | Casarica", type=["png", "jpg", "jpeg", "webp"], key="alex_img")

    rows = []

    if pdf is not None:
        df_pdf = process_pdf_by_client(pdf, "Alex")
        for _, row in df_pdf.iterrows():
            total_base, rebate, total_final = calc_alex_row(float(row["ganhos"]), float(row["rake"]), float(row["rb_percentual"]))
            rows.append({
                "AGENTE": row["agente"],
                "GANHOS": float(row["ganhos"]),
                "RAKE": float(row["rake"]),
                "RB": f"{int(float(row['rb_percentual']))}%",
                "TOTAL": total_base,
                "_REBATE": rebate,
                "_TOTAL_FINAL": total_final,
            })

    if imagem is not None:
        img = Image.open(imagem)
        if detect_alex_casarica_image(img):
            dados = extract_alex_casarica_values(img)
            with st.expander("Diagnóstico OCR - Casarica", expanded=False):
                st.write(f"Formato detectado: {dados['formato']}")
                if dados["ocr_taxa_crop"]:
                    st.code("OCR recorte Taxa:\n" + dados["ocr_taxa_crop"])
                if dados["ocr_ganhos_crop"]:
                    st.code("OCR recorte Ganhos:\n" + dados["ocr_ganhos_crop"])
            if dados["formato"] == "taxa_ganhos_falhou":
                st.warning("Não consegui ler Taxa/Ganhos com segurança nessa imagem.")
            else:
                rows.append({
                    "AGENTE": dados["agente"],
                    "GANHOS": dados["ganhos"],
                    "RAKE": dados["rake"],
                    "RB": f"{int(RB_ALEX_IMAGEM_REAL)}%",
                    "TOTAL": dados["total_base"],
                    "_REBATE": dados["rebate"],
                    "_TOTAL_FINAL": dados["total_final"],
                })
        else:
            st.warning("Não identifiquei a imagem Gustavo | Casarica com segurança.")

    if st.button("Gerar fechamento Alex", type="primary", key="btn_alex"):
        if not rows:
            st.warning("Envie o PDF e/ou a imagem do Casarica.")
            return
        დეტ = pd.DataFrame(rows)
        rebate_total = დეტ["_REBATE"].sum()
        total_geral = დეტ["_TOTAL_FINAL"].sum()
        report = generate_client_table_image("ALEX", periodo.strip() or "-", დეტ[["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]], total_geral, rebate_total, "Ajuste Alex")
        st.image(report, caption="Pronto para print", use_container_width=True)
        st.download_button("Baixar relatório em PNG", data=to_png_bytes(report), file_name="alex_fechamento.png", mime="image/png")


st.title("Fechamentos Premier")
cliente = st.selectbox("Escolha o cliente", ["Harnefer", "Demetra", "Oscar", "Alex"])

if cliente == "Harnefer":
    page_harnefer()
elif cliente == "Demetra":
    page_demetra()
elif cliente == "Oscar":
    page_oscar()
else:
    page_alex()
