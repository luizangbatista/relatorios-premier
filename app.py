
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

RB_DEMETRA_PLANILHA = 70.0
AGENTE_PLANILHA_DEMETRA = "TheShark_ (ID11719117)"


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


def preprocess_for_ocr(img: Image.Image) -> Image.Image:
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    g = ImageEnhance.Contrast(g).enhance(2.7)
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


def normalize_id(v) -> str:
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[^\d]", "", s)


def detect_harnefer_image(img: Image.Image) -> bool:
    text = (ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)).upper()
    return ("TOTAL FEE" in text and "WINNINGS" in text) or ("GAMES" in text and "ADMIN FEE" in text)


def crop_harnefer_summary(img: Image.Image) -> Image.Image:
    w, h = img.size
    return img.crop((int(w * 0.08), int(h * 0.32), int(w * 0.93), int(h * 0.64)))


def first_money(text: str) -> float:
    for pat in [r"([0-9]{1,3}(?:[,\.][0-9]{3})*[,\.][0-9]{2})", r"([0-9]+[,\.][0-9]{2})"]:
        m = re.search(pat, text)
        if m:
            return parse_money(m.group(1))
    return 0.0


def extract_harnefer_values(img: Image.Image) -> dict:
    crop = crop_harnefer_summary(img)
    w, h = crop.size
    col_w = w // 4
    cols = [crop.crop((i * col_w, 0, (i + 1) * col_w if i < 3 else w, h)) for i in range(4)]
    texts = [ocr_image(c, psm=6) + "\n" + ocr_image(c, psm=11) for c in cols]
    rake = first_money(texts[1])
    ganhos = first_money(texts[3])
    rb = rake * 0.76
    total = ganhos + rb
    return {"ganhos": ganhos, "rake": rake, "rakeback": rb, "total_final": total}


def process_demetra_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, usecols=[5, 6, 7, 8, 9, 29])
    df.columns = ["origem", "id_conta", "nick", "codigo", "ganhos", "rake"]
    df["id_conta"] = df["id_conta"].apply(normalize_id)
    df["codigo"] = df["codigo"].apply(normalize_id)
    df["origem"] = df["origem"].astype(str).str.strip()
    df["ganhos"] = pd.to_numeric(df["ganhos"], errors="coerce").fillna(0.0)
    df["rake"] = pd.to_numeric(df["rake"], errors="coerce").fillna(0.0)
    df = df[df["codigo"] == "802606"].copy()
    if df.empty:
        return pd.DataFrame()
    mask = (df["id_conta"] == "11719117") & (df["origem"] == "TheShark_")
    return df[mask].copy()


def extract_pdf_lines(uploaded_file):
    lines = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend([ln.strip() for ln in text.splitlines() if ln.strip()])
    return lines


def process_demetra_pdf(uploaded_file):
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
        if not info or info["cliente"] != "Demetra":
            continue
        ganhos = parse_money(money_matches[0])
        rake = parse_money(money_matches[1])
        agente = re.sub(r"\s{2,}", " ", line.split(id_agente)[0].strip())
        rows.append({"agente": agente, "ganhos": ganhos, "rake": rake, "rb_percentual": float(info["rb"])})
    return pd.DataFrame(rows)


def calc_row(ganhos: float, rake: float, rb_percentual: float):
    rb_valor = rake * (rb_percentual / 100.0)
    total_base = ganhos + rb_valor
    rebate = total_base * (-0.05) if total_base > 0 else 0.0
    total_final = total_base + rebate
    return total_base, rebate, total_final


def generate_harnefer_report(periodo: str, ganhos: float, rake: float) -> Image.Image:
    return generate_summary_report(
        titulo="HARNEFER",
        periodo=periodo,
        headers=["GANHOS", "RAKE", "RB (76%)", "TOTAL"],
        values=[ganhos, rake, rake * 0.76, ganhos + rake * 0.76],
        rebate=0.0,
        total_final=ganhos + rake * 0.76,
    )


def generate_summary_report(titulo: str, periodo: str, headers, values, rebate: float, total_final: float) -> Image.Image:
    navy, gold, green, red = (7, 29, 69), (199, 143, 43), (0, 102, 45), (170, 30, 30)
    gray, light_bg, yellow = (248, 248, 248), (234, 241, 235), (248, 238, 27)
    W, H = 1400, 980
    img = Image.new("RGB", (W, H), gray)
    draw = ImageDraw.Draw(img)

    title_font = get_font(45, bold=True)
    subtitle_font = get_font(36, bold=False)
    subtitle_bold = get_font(36, bold=True)
    header_font = get_font(30, bold=True)
    value_font = get_font(30, bold=False)
    small_font = get_font(30, bold=False)
    total_font = get_font(30, bold=True)
    status_font = get_font(36, bold=True)
    status_value_font = get_font(36, bold=True)

    tw, _ = measure(draw, titulo, title_font)
    draw.text(((W - tw) / 2, 40), titulo, fill=navy, font=title_font)
    draw.line((70, 115, 300, 115), fill=gold, width=3)
    draw.line((1100, 115, 1330, 115), fill=gold, width=3)

    label = "Período do fechamento:"
    l_w, _ = measure(draw, label, subtitle_font)
    p_w, _ = measure(draw, periodo, subtitle_bold)
    x0 = (W - (l_w + 12 + p_w)) / 2
    draw.text((x0, 150), label, fill=navy, font=subtitle_font)
    draw.text((x0 + l_w + 12, 150), periodo, fill=navy, font=subtitle_bold)

    table_x1, table_x2 = 60, 1340
    header_y1, header_y2 = 250, 325
    values_y1, values_y2 = 325, 455
    draw.rounded_rectangle((table_x1, header_y1, table_x2, values_y2), radius=14, outline=navy, width=3, fill="white")
    draw.rounded_rectangle((table_x1, header_y1, table_x2, header_y2), radius=14, fill=navy)
    draw.rectangle((table_x1, header_y2 - 14, table_x2, header_y2), fill=navy)

    col_w = (table_x2 - table_x1) / 4
    for i in range(1, 4):
        x = table_x1 + i * col_w
        draw.line((x, header_y1, x, values_y2), fill=(160, 160, 160), width=2)
    for i, col in enumerate(headers):
        cx = table_x1 + i * col_w + col_w / 2
        cw, _ = measure(draw, col, header_font)
        draw.text((cx - cw / 2, header_y1 + 22), col, fill="white", font=header_font)
    for i, val in enumerate(values):
        sval = fmt_brl(val)
        cx = table_x1 + i * col_w + col_w / 2
        vw, _ = measure(draw, sval, value_font)
        draw.text((cx - vw / 2, values_y1 + 45), sval, fill=navy, font=value_font)

    draw.rectangle((60, 490, 1340, 550), fill=(230, 230, 230))
    rebate_label = "-5% total" if rebate != 0 else "Sem rebate"
    rebate_value = fmt_brl(rebate) if rebate != 0 else "0,00"
    rl_w, _ = measure(draw, rebate_label, small_font)
    rv_w, _ = measure(draw, rebate_value, small_font)
    draw.text((1000 - rl_w, 506), rebate_label, fill=navy, font=small_font)
    draw.text((1310 - rv_w, 506), rebate_value, fill=navy, font=small_font)

    draw.rectangle((60, 550, 1340, 635), fill=yellow)
    tlabel, tvalue = "TOTAL", fmt_brl(total_final)
    tl_w, _ = measure(draw, tlabel, total_font)
    tv_w, _ = measure(draw, tvalue, total_font)
    draw.text((1040 - tl_w, 575), tlabel, fill=(0, 0, 0), font=total_font)
    draw.text((1310 - tv_w, 575), tvalue, fill=(0, 0, 0), font=total_font)

    status_text = "PREMIER TEM A PAGAR" if total_final > 0 else ("PREMIER TEM A RECEBER" if total_final < 0 else "SEM VALORES")
    status_value = f"R$ {fmt_brl(abs(total_final))}"
    draw.rounded_rectangle((60, 710, 1340, 835), radius=16, outline=navy, width=3, fill=light_bg)
    stw, _ = measure(draw, status_text, status_font)
    svw, _ = measure(draw, status_value, status_value_font)
    start_x = (W - (stw + 24 + svw)) / 2
    draw.text((start_x, 748), status_text, fill=navy, font=status_font)
    draw.text((start_x + stw + 24, 744), status_value, fill=green if total_final >= 0 else red, font=status_value_font)
    return img


def generate_demetra_table_image(periodo: str, df: pd.DataFrame, total_geral: float, rebate_total: float) -> Image.Image:
    navy, gold, green, red = (7, 29, 69), (199, 143, 43), (0, 102, 45), (170, 30, 30)
    white, gray, grid = (255, 255, 255), (248, 248, 248), (70, 70, 70)
    light_bg, yellow, light_gray = (234, 241, 235), (248, 238, 27), (230, 230, 230)

    n_rows = len(df)
    row_h = 52
    table_top = 220
    header_h = 52
    W = 1450
    H = table_top + header_h + (n_rows + 3) * row_h + 180
    img = Image.new("RGB", (W, H), gray)
    draw = ImageDraw.Draw(img)

    title_font = get_font(32, bold=True)
    subtitle_font = get_font(30, bold=False)
    subtitle_bold = get_font(30, bold=True)
    header_font = get_font(28, bold=True)
    cell_font = get_font(28, bold=False)
    total_font = get_font(30, bold=True)
    status_font = get_font(30, bold=True)
    status_value_font = get_font(32, bold=True)

    title = "DEMETRA"
    tw, _ = measure(draw, title, title_font)
    draw.text(((W - tw) / 2, 20), title, fill=navy, font=title_font)
    draw.line((60, 70, 250, 70), fill=gold, width=3)
    draw.line((1190, 70, 1380, 70), fill=gold, width=3)

    label = "Período do fechamento:"
    l_w, _ = measure(draw, label, subtitle_font)
    p_w, _ = measure(draw, periodo, subtitle_bold)
    x0 = (W - (l_w + 12 + p_w)) / 2
    draw.text((x0, 100), label, fill=navy, font=subtitle_font)
    draw.text((x0 + l_w + 12, 100), periodo, fill=navy, font=subtitle_bold)

    x1 = 50
    widths = [500, 210, 210, 130, 260]
    headers = ["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]
    xs = [x1]
    for w in widths:
        xs.append(xs[-1] + w)
    table_w = sum(widths)

    draw.rectangle((x1, table_top, x1 + table_w, table_top + header_h), fill=navy)
    for i, htxt in enumerate(headers):
        cx = (xs[i] + xs[i + 1]) / 2
        cw, _ = measure(draw, htxt, header_font)
        draw.text((cx - cw / 2, table_top + 12), htxt, fill=white, font=header_font)

    y = table_top + header_h
    for _, row in df.iterrows():
        draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=white, outline=grid, width=2)
        vals = [str(row["AGENTE"]), fmt_brl(row["GANHOS"]), fmt_brl(row["RAKE"]), str(row["RB"]), fmt_brl(row["TOTAL"])]
        for i, val in enumerate(vals):
            if i == 0:
                draw.text((xs[i] + 12, y + 12), val, fill=(0, 0, 0), font=cell_font)
            else:
                vw, _ = measure(draw, val, cell_font)
                cx = (xs[i] + xs[i + 1]) / 2
                draw.text((cx - vw / 2, y + 12), val, fill=(0, 0, 0), font=cell_font)
        y += row_h

    draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=white, outline=grid, width=2)
    draw.rectangle((x1, y, xs[4], y + row_h), fill=navy)
    total_label = "TOTAL"
    tlw, _ = measure(draw, total_label, total_font)
    draw.text((xs[4] - tlw - 16, y + 10), total_label, fill=white, font=total_font)
    total_str = fmt_brl(total_geral)
    tvw, _ = measure(draw, total_str, total_font)
    draw.text((((xs[4] + xs[5]) / 2) - tvw / 2, y + 10), total_str, fill=(0, 0, 0), font=total_font)
    y += row_h

    draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=light_gray, outline=grid, width=2)
    rlabel = "-5% total" if rebate_total != 0 else "Sem rebate"
    rvalue = fmt_brl(rebate_total) if rebate_total != 0 else "0,00"
    rlw, _ = measure(draw, rlabel, cell_font)
    rvw, _ = measure(draw, rvalue, cell_font)
    draw.text((xs[4] - rlw - 16, y + 12), rlabel, fill=(0, 0, 0), font=cell_font)
    draw.text((((xs[4] + xs[5]) / 2) - rvw / 2, y + 12), rvalue, fill=(0, 0, 0), font=cell_font)
    y += row_h

    draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=yellow, outline=grid, width=2)
    flw, _ = measure(draw, total_label, total_font)
    fvw, _ = measure(draw, total_str, total_font)
    draw.text((xs[4] - flw - 16, y + 10), total_label, fill=(0, 0, 0), font=total_font)
    draw.text((((xs[4] + xs[5]) / 2) - fvw / 2, y + 10), total_str, fill=(0, 0, 0), font=total_font)
    y += row_h

    status_text = "PREMIER TEM A PAGAR" if total_geral > 0 else ("PREMIER TEM A RECEBER" if total_geral < 0 else "SEM VALORES")
    status_value = f"R$ {fmt_brl(abs(total_geral))}"
    box_y1 = y + 35
    box_y2 = box_y1 + 90
    draw.rounded_rectangle((50, box_y1, x1 + table_w, box_y2), radius=14, outline=navy, width=3, fill=light_bg)
    stw, _ = measure(draw, status_text, status_font)
    svw, _ = measure(draw, status_value, status_value_font)
    start_x = (W - (stw + 20 + svw)) / 2
    draw.text((start_x, box_y1 + 28), status_text, fill=navy, font=status_font)
    draw.text((start_x + stw + 20, box_y1 + 24), status_value, fill=green if total_geral >= 0 else red, font=status_value_font)
    return img


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
            total_base, rebate, total_final = calc_row(ganhos_excel, rake_excel, RB_DEMETRA_PLANILHA)
            rows.append({
                "AGENTE": AGENTE_PLANILHA_DEMETRA,
                "GANHOS": ganhos_excel,
                "RAKE": rake_excel,
                "RB": f"{int(RB_DEMETRA_PLANILHA)}%",
                "TOTAL": total_base,
                "_REBATE": rebate,
                "_TOTAL_FINAL": total_final,
            })
        else:
            st.info("Nenhuma linha encontrada para TheShark_ com ID 11719117 na planilha.")

    if pdf is not None:
        df_pdf = process_demetra_pdf(pdf)
        for _, row in df_pdf.iterrows():
            total_base, rebate, total_final = calc_row(float(row["ganhos"]), float(row["rake"]), float(row["rb_percentual"]))
            rows.append({
                "AGENTE": row["agente"],
                "GANHOS": float(row["ganhos"]),
                "RAKE": float(row["rake"]),
                "RB": f"{int(float(row['rb_percentual']))}%",
                "TOTAL": total_base,
                "_REBATE": rebate,
                "_TOTAL_FINAL": total_final,
            })

    if st.button("Gerar fechamento Demetra", type="primary", key="btn_demetra"):
        if not rows:
            st.warning("Envie a planilha e/ou o PDF.")
            return
        detalhado = pd.DataFrame(rows)
        rebate_total = detalhado["_REBATE"].sum()
        total_geral = detalhado["_TOTAL_FINAL"].sum()
        report = generate_demetra_table_image(
            periodo.strip() or "-",
            detalhado[["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]],
            total_geral,
            rebate_total,
        )
        st.image(report, caption="Pronto para print", use_container_width=True)
        st.download_button("Baixar relatório em PNG", data=to_png_bytes(report), file_name="demetra_fechamento.png", mime="image/png")


st.title("Fechamentos Premier")
cliente = st.selectbox("Escolha o cliente", ["Harnefer", "Demetra"])
if cliente == "Harnefer":
    page_harnefer()
else:
    page_demetra()
