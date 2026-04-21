
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
    candidates = []
    if bold:
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        ]
    else:
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
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
    config = f"--oem 3 --psm {psm}"
    try:
        return pytesseract.image_to_string(proc, lang="eng", config=config) or ""
    except Exception:
        return ""


def normalize_id(v) -> str:
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[^\d]", "", s)


# ---------------- Harnefer ----------------
def detect_harnefer_image(img: Image.Image) -> bool:
    text = (ocr_image(img, psm=6) + "\n" + ocr_image(img, psm=11)).upper()
    return ("TOTAL FEE" in text and "WINNINGS" in text) or ("GAMES" in text and "ADMIN FEE" in text)


def crop_harnefer_summary(img: Image.Image) -> Image.Image:
    w, h = img.size
    left = int(w * 0.08)
    right = int(w * 0.93)
    top = int(h * 0.32)
    bottom = int(h * 0.64)
    return img.crop((left, top, right, bottom))


def first_money(text: str) -> float:
    patterns = [
        r"([0-9]{1,3}(?:[,\.][0-9]{3})*[,\.][0-9]{2})",
        r"([0-9]+[,\.][0-9]{2})",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return parse_money(m.group(1))
    return 0.0


def extract_harnefer_values(img: Image.Image) -> dict:
    crop = crop_harnefer_summary(img)
    w, h = crop.size
    col_w = w // 4

    cols = []
    for i in range(4):
        x1 = i * col_w
        x2 = (i + 1) * col_w if i < 3 else w
        cols.append(crop.crop((x1, 0, x2, h)))

    texts = [ocr_image(c, psm=6) + "\n" + ocr_image(c, psm=11) for c in cols]

    rake = first_money(texts[1])      # valor acima de Total Fee
    ganhos = first_money(texts[3])    # valor acima de Winnings
    rb = rake * 0.76
    total = ganhos + rb

    return {
        "ganhos": ganhos,
        "rake": rake,
        "rb_percentual": 76.0,
        "rakeback": rb,
        "total_final": total,
        "ocr_fee": texts[1],
        "ocr_winnings": texts[3],
        "crop": crop,
    }


# ---------------- Demetra ----------------
def process_demetra_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, usecols=[5, 6, 7, 8, 9, 29])
    df.columns = ["origem", "id_conta", "nick", "codigo", "ganhos", "rake"]

    df["id_conta"] = df["id_conta"].apply(normalize_id)
    df["codigo"] = df["codigo"].apply(normalize_id)
    df["origem"] = df["origem"].astype(str).str.strip()
    df["nick"] = df["nick"].astype(str).str.strip()
    df["ganhos"] = pd.to_numeric(df["ganhos"], errors="coerce").fillna(0.0)
    df["rake"] = pd.to_numeric(df["rake"], errors="coerce").fillna(0.0)

    df = df[df["codigo"] == "802606"].copy()
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    demetra_mask = (df["id_conta"] == "11719117") & (df["nick"].str.lower() == "killuminatti")
    demetra_df = df[demetra_mask].copy()
    outros_df = df[~demetra_mask].copy()
    return demetra_df, outros_df


def extract_pdf_lines(uploaded_file):
    lines = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend([ln.strip() for ln in text.splitlines() if ln.strip()])
    return lines


def process_demetra_pdf(uploaded_file):
    lines = extract_pdf_lines(uploaded_file)
    rows = []
    unknown_rows = []

    for line in lines:
        if "R$" not in line:
            continue

        id_match = re.search(r"\b(\d{6,9})\b", line)
        if not id_match:
            continue

        id_agente = normalize_id(id_match.group(1))
        money_matches = re.findall(r"R\$\s*-?\d[\d\.,]*", line)
        if len(money_matches) < 2:
            continue

        ganhos = parse_money(money_matches[0])
        rake = parse_money(money_matches[1])

        info = MAPA_IDS_PDF.get(id_agente)
        if info is None:
            unknown_rows.append({
                "id_agente": id_agente,
                "linha": line,
                "ganhos": ganhos,
                "rake": rake,
            })
            continue

        if info["cliente"] == "Demetra":
            agente_raw = line.split(id_agente)[0].strip()
            agente = re.sub(r"\s{2,}", " ", agente_raw)
            rows.append({
                "agente": agente,
                "id": id_agente,
                "ganhos": ganhos,
                "rake": rake,
                "rb_percentual": float(info["rb"]),
            })

    return pd.DataFrame(rows), unknown_rows


def calc_row(ganhos: float, rake: float, rb_percentual: float):
    rb_valor = rake * (rb_percentual / 100.0)
    total_base = ganhos + rb_valor
    rebate = total_base * (-0.05) if total_base > 0 else 0.0
    total_final = total_base + rebate
    return rb_valor, total_base, rebate, total_final


# ---------------- Report images ----------------
def generate_harnefer_report(periodo: str, ganhos: float, rake: float) -> Image.Image:
    total_final = ganhos + (rake * 0.76)
    return generate_summary_report(
        titulo="HARNEFER",
        periodo=periodo,
        headers=["GANHOS", "RAKE", "RB (76%)", "TOTAL"],
        values=[ganhos, rake, rake * 0.76, total_final],
        rebate=0.0,
        total_final=total_final,
    )


def generate_summary_report(titulo: str, periodo: str, headers, values, rebate: float, total_final: float) -> Image.Image:
    navy = (7, 29, 69)
    gold = (199, 143, 43)
    green = (0, 102, 45)
    red = (170, 30, 30)
    gray = (248, 248, 248)
    light_bg = (234, 241, 235)
    yellow = (248, 238, 27)

    W, H = 1800, 1280
    img = Image.new("RGB", (W, H), gray)
    draw = ImageDraw.Draw(img)

    title_font = get_font(118, bold=True)
    subtitle_font = get_font(40, bold=False)
    subtitle_bold = get_font(40, bold=True)
    header_font = get_font(38, bold=True)
    value_font = get_font(46, bold=False)
    small_font = get_font(28, bold=False)
    total_font = get_font(52, bold=True)
    status_font = get_font(54, bold=True)
    status_value_font = get_font(64, bold=True)

    tw, _ = measure(draw, titulo, title_font)
    draw.text(((W - tw) / 2, 60), titulo, fill=navy, font=title_font)

    draw.line((80, 250, 450, 250), fill=gold, width=4)
    draw.line((1350, 250, 1720, 250), fill=gold, width=4)
    draw.polygon([(450, 235), (468, 250), (450, 265), (432, 250)], fill=gold)
    draw.polygon([(1350, 235), (1368, 250), (1350, 265), (1332, 250)], fill=gold)

    label = "Período do fechamento:"
    l_w, _ = measure(draw, label, subtitle_font)
    p_w, _ = measure(draw, periodo, subtitle_bold)
    total_sub_w = l_w + 18 + p_w
    x0 = (W - total_sub_w) / 2
    draw.text((x0, 320), label, fill=navy, font=subtitle_font)
    draw.text((x0 + l_w + 18, 320), periodo, fill=navy, font=subtitle_bold)

    table_x1, table_x2 = 70, 1730
    header_y1, header_y2 = 470, 610
    values_y1, values_y2 = 610, 825

    draw.rounded_rectangle((table_x1, header_y1, table_x2, values_y2), radius=20, outline=navy, width=5, fill="white")
    draw.rounded_rectangle((table_x1, header_y1, table_x2, header_y2), radius=20, fill=navy)
    draw.rectangle((table_x1, header_y2 - 20, table_x2, header_y2), fill=navy)

    col_w = (table_x2 - table_x1) / 4
    for i in range(1, 4):
        x = table_x1 + i * col_w
        draw.line((x, header_y1, x, values_y2), fill=(160, 160, 160), width=2)

    for i, col in enumerate(headers):
        cx = table_x1 + i * col_w + col_w / 2
        cw, _ = measure(draw, col, header_font)
        draw.text((cx - cw / 2, header_y1 + 42), col, fill="white", font=header_font)

    for i, val in enumerate(values):
        sval = fmt_brl(val)
        cx = table_x1 + i * col_w + col_w / 2
        vw, _ = measure(draw, sval, value_font)
        draw.text((cx - vw / 2, values_y1 + 78), sval, fill=navy, font=value_font)

    mid_y1, mid_y2 = 865, 935
    draw.rectangle((70, mid_y1, 1730, mid_y2), fill=(230, 230, 230))
    rebate_label = "-5% total" if rebate != 0 else "Sem rebate"
    rebate_value = fmt_brl(rebate) if rebate != 0 else "0,00"
    rl_w, _ = measure(draw, rebate_label, small_font)
    rv_w, _ = measure(draw, rebate_value, small_font)
    draw.text((1320 - rl_w, 885), rebate_label, fill=navy, font=small_font)
    draw.text((1720 - rv_w, 885), rebate_value, fill=navy, font=small_font)

    total_y1, total_y2 = 935, 1060
    draw.rectangle((70, total_y1, 1730, total_y2), fill=yellow)
    tlabel = "TOTAL"
    tvalue = fmt_brl(total_final)
    tl_w, _ = measure(draw, tlabel, total_font)
    tv_w, _ = measure(draw, tvalue, total_font)
    draw.text((1380 - tl_w, 968), tlabel, fill=(0, 0, 0), font=total_font)
    draw.text((1715 - tv_w, 968), tvalue, fill=(0, 0, 0), font=total_font)

    status_text = "PREMIER TEM A PAGAR" if total_final > 0 else ("PREMIER TEM A RECEBER" if total_final < 0 else "SEM VALORES")
    status_value = f"R$ {fmt_brl(abs(total_final))}"

    box_y1, box_y2 = 1100, 1240
    draw.rounded_rectangle((70, box_y1, 1730, box_y2), radius=20, outline=navy, width=4, fill=light_bg)

    stw, _ = measure(draw, status_text, status_font)
    svw, _ = measure(draw, status_value, status_value_font)
    start_x = (W - (stw + 40 + svw)) / 2
    draw.text((start_x, 1132), status_text, fill=navy, font=status_font)
    draw.text((start_x + stw + 40, 1120), status_value, fill=green if total_final >= 0 else red, font=status_value_font)

    return img


def generate_demetra_table_image(periodo: str, df: pd.DataFrame, total_geral: float, rebate_total: float) -> Image.Image:
    navy = (7, 29, 69)
    gold = (199, 143, 43)
    green = (0, 102, 45)
    red = (170, 30, 30)
    white = (255, 255, 255)
    gray = (248, 248, 248)
    grid = (70, 70, 70)
    light_bg = (234, 241, 235)
    yellow = (248, 238, 27)
    light_gray = (230, 230, 230)

    n_rows = len(df)
    row_h = 76
    table_top = 420
    header_h = 76
    footer_rows_h = 3 * row_h
    status_h = 140
    W = 1900
    H = table_top + header_h + (n_rows + 3) * row_h + status_h + 120

    img = Image.new("RGB", (W, H), gray)
    draw = ImageDraw.Draw(img)

    title_font = get_font(110, bold=True)
    subtitle_font = get_font(38, bold=False)
    subtitle_bold = get_font(38, bold=True)
    header_font = get_font(30, bold=True)
    cell_font = get_font(28, bold=False)
    total_font = get_font(34, bold=True)
    status_font = get_font(52, bold=True)
    status_value_font = get_font(58, bold=True)

    title = "DEMETRA"
    tw, _ = measure(draw, title, title_font)
    draw.text(((W - tw) / 2, 45), title, fill=navy, font=title_font)
    draw.line((100, 215, 420, 215), fill=gold, width=4)
    draw.line((1480, 215, 1800, 215), fill=gold, width=4)
    draw.polygon([(420, 200), (438, 215), (420, 230), (402, 215)], fill=gold)
    draw.polygon([(1480, 200), (1498, 215), (1480, 230), (1462, 215)], fill=gold)

    label = "Período do fechamento:"
    l_w, _ = measure(draw, label, subtitle_font)
    p_w, _ = measure(draw, periodo, subtitle_bold)
    total_sub_w = l_w + 18 + p_w
    x0 = (W - total_sub_w) / 2
    draw.text((x0, 280), label, fill=navy, font=subtitle_font)
    draw.text((x0 + l_w + 18, 280), periodo, fill=navy, font=subtitle_bold)

    x1 = 70
    widths = [620, 250, 250, 180, 320]  # agente, ganhos, rake, rb, total
    headers = ["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]
    xs = [x1]
    for w in widths:
        xs.append(xs[-1] + w)
    table_w = sum(widths)

    # header
    draw.rectangle((x1, table_top, x1 + table_w, table_top + header_h), fill=navy)
    for i, htxt in enumerate(headers):
        cx = (xs[i] + xs[i+1]) / 2
        cw, _ = measure(draw, htxt, header_font)
        draw.text((cx - cw / 2, table_top + 22), htxt, fill=white, font=header_font)

    # data rows
    y = table_top + header_h
    for _, row in df.iterrows():
        draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=white, outline=grid, width=2)
        vals = [
            str(row["AGENTE"]),
            fmt_brl(row["GANHOS"]),
            fmt_brl(row["RAKE"]),
            str(row["RB"]),
            fmt_brl(row["TOTAL"]),
        ]
        aligns = ["left", "center", "center", "center", "center"]
        for i, val in enumerate(vals):
            if aligns[i] == "left":
                draw.text((xs[i] + 16, y + 22), val, fill=(0, 0, 0), font=cell_font)
            else:
                vw, _ = measure(draw, val, cell_font)
                cx = (xs[i] + xs[i+1]) / 2
                draw.text((cx - vw / 2, y + 22), val, fill=(0, 0, 0), font=cell_font)
        y += row_h

    # total row
    draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=white, outline=grid, width=2)
    draw.rectangle((x1, y, xs[4], y + row_h), fill=navy)
    total_label = "TOTAL"
    tlw, _ = measure(draw, total_label, total_font)
    draw.text((xs[4] - tlw - 20, y + 18), total_label, fill=white, font=total_font)
    total_str = fmt_brl(total_geral)
    tvw, _ = measure(draw, total_str, total_font)
    draw.text((((xs[4] + xs[5]) / 2) - tvw / 2, y + 18), total_str, fill=(0, 0, 0), font=total_font)
    y += row_h

    # rebate row
    draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=light_gray, outline=grid, width=2)
    rlabel = "-5% total" if rebate_total != 0 else "Sem rebate"
    rvalue = fmt_brl(rebate_total) if rebate_total != 0 else "0,00"
    rlw, _ = measure(draw, rlabel, cell_font)
    rvw, _ = measure(draw, rvalue, cell_font)
    draw.text((xs[4] - rlw - 20, y + 22), rlabel, fill=(0, 0, 0), font=cell_font)
    draw.text((((xs[4] + xs[5]) / 2) - rvw / 2, y + 22), rvalue, fill=(0, 0, 0), font=cell_font)
    y += row_h

    # final total band
    draw.rectangle((x1, y, x1 + table_w, y + row_h), fill=yellow, outline=grid, width=2)
    final_label = "TOTAL"
    flw, _ = measure(draw, final_label, total_font)
    final_str = fmt_brl(total_geral)
    fvw, _ = measure(draw, final_str, total_font)
    draw.text((xs[4] - flw - 20, y + 18), final_label, fill=(0, 0, 0), font=total_font)
    draw.text((((xs[4] + xs[5]) / 2) - fvw / 2, y + 18), final_str, fill=(0, 0, 0), font=total_font)
    y += row_h

    # status
    status_text = "PREMIER TEM A PAGAR" if total_geral > 0 else ("PREMIER TEM A RECEBER" if total_geral < 0 else "SEM VALORES")
    status_value = f"R$ {fmt_brl(abs(total_geral))}"

    box_y1 = y + 55
    box_y2 = box_y1 + 130
    draw.rounded_rectangle((70, box_y1, x1 + table_w, box_y2), radius=20, outline=navy, width=4, fill=light_bg)

    stw, _ = measure(draw, status_text, status_font)
    svw, _ = measure(draw, status_value, status_value_font)
    start_x = (W - (stw + 30 + svw)) / 2
    draw.text((start_x, box_y1 + 34), status_text, fill=navy, font=status_font)
    draw.text((start_x + stw + 30, box_y1 + 24), status_value, fill=green if total_geral >= 0 else red, font=status_value_font)

    return img


# ---------------- Pages ----------------
def page_harnefer():
    st.subheader("Harnefer")
    periodo = st.text_input("Período do fechamento", key="periodo_harnefer", placeholder="13/04/2026 a 19/04/2026")
    arquivo = st.file_uploader("Envie a imagem do Harnefer", type=["png", "jpg", "jpeg", "webp"], key="harnefer_img")

    if not TESSERACT_OK:
        st.error("OCR indisponível: instale `pytesseract` e o programa `tesseract-ocr` no ambiente.")
        st.code("packages.txt:\ntesseract-ocr\ntesseract-ocr-eng")
        return

    if arquivo:
        img = Image.open(arquivo)
        st.image(img, caption="Imagem enviada", width=420)

        if st.button("Ler imagem e gerar fechamento", type="primary", key="btn_harnefer"):
            if not detect_harnefer_image(img):
                st.warning("Não identifiquei a imagem do Harnefer com segurança. Ela precisa mostrar os cards com Total Fee e Winnings.")
                return

            dados = extract_harnefer_values(img)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Ganhos", f"R$ {fmt_brl(dados['ganhos'])}")
            c2.metric("Rake", f"R$ {fmt_brl(dados['rake'])}")
            c3.metric("RB (76%)", f"R$ {fmt_brl(dados['rakeback'])}")
            c4.metric("Total", f"R$ {fmt_brl(dados['total_final'])}")

            with st.expander("Diagnóstico OCR"):
                st.image(dados["crop"], caption="Recorte usado para leitura", width=700)
                st.code(dados["ocr_fee"])
                st.code(dados["ocr_winnings"])

            if periodo.strip():
                report = generate_harnefer_report(periodo.strip(), dados["ganhos"], dados["rake"])
                st.image(report, caption="Relatório final", use_container_width=True)
                st.download_button("Baixar relatório em PNG", data=to_png_bytes(report), file_name="harnefer_fechamento.png", mime="image/png")
            else:
                st.warning("Preencha o período do fechamento.")


def page_demetra():
    st.subheader("Demetra")
    periodo = st.text_input("Período do fechamento", key="periodo_demetra", placeholder="06/04/2026 a 12/04/2026")
    planilha = st.file_uploader("Envie a planilha 2101...", type=["xlsx", "xls"], key="demetra_xlsx")
    pdf = st.file_uploader("Envie o PDF", type=["pdf"], key="demetra_pdf")

    excel_rows = []
    pdf_rows = []
    unknown_pdf_rows = []

    if planilha is not None:
        demetra_df, outros_df = process_demetra_excel(planilha)

        if not demetra_df.empty:
            ganhos_excel = demetra_df["ganhos"].sum()
            rake_excel = demetra_df["rake"].sum()
            rb_excel = 76.0
            rb_valor, total_base, rebate, total_final = calc_row(ganhos_excel, rake_excel, rb_excel)
            excel_rows.append({
                "AGENTE": "Killuminatti",
                "GANHOS": ganhos_excel,
                "RAKE": rake_excel,
                "RB": f"{int(rb_excel)}%",
                "TOTAL": total_base,
                "_REBATE": rebate,
                "_TOTAL_FINAL": total_final,
            })

        if not outros_df.empty:
            st.write("### Linhas da planilha não reconhecidas automaticamente")
            st.caption("Só entram no fechamento do Demetra se você marcar Demetra.")
            for idx, row in outros_df.reset_index(drop=True).iterrows():
                with st.container(border=True):
                    st.write(f"Origem: {row['origem']} | ID: {row['id_conta']} | Nick: {row['nick']}")
                    c1, c2 = st.columns(2)
                    with c1:
                        cliente = st.selectbox(
                            f"Agência da linha {idx + 1}",
                            ["Demetra", "Alex", "Oscar", "Outro"],
                            key=f"cliente_excel_{idx}",
                        )
                    with c2:
                        rb = st.number_input(
                            f"%RB da linha {idx + 1}",
                            min_value=0.0,
                            max_value=100.0,
                            value=76.0 if cliente == "Demetra" else 0.0,
                            step=1.0,
                            key=f"rb_excel_{idx}",
                        )

                    if cliente == "Demetra":
                        rb_valor, total_base, rebate, total_final = calc_row(float(row["ganhos"]), float(row["rake"]), float(rb))
                        excel_rows.append({
                            "AGENTE": str(row["nick"]) if str(row["nick"]).strip() else "Linha planilha",
                            "GANHOS": float(row["ganhos"]),
                            "RAKE": float(row["rake"]),
                            "RB": f"{int(float(rb))}%",
                            "TOTAL": total_base,
                            "_REBATE": rebate,
                            "_TOTAL_FINAL": total_final,
                        })

    if pdf is not None:
        df_pdf, unknown_pdf_rows = process_demetra_pdf(pdf)
        if not df_pdf.empty:
            for _, row in df_pdf.iterrows():
                rb_valor, total_base, rebate, total_final = calc_row(float(row["ganhos"]), float(row["rake"]), float(row["rb_percentual"]))
                pdf_rows.append({
                    "AGENTE": row["agente"],
                    "GANHOS": float(row["ganhos"]),
                    "RAKE": float(row["rake"]),
                    "RB": f"{int(float(row['rb_percentual']))}%",
                    "TOTAL": total_base,
                    "_REBATE": rebate,
                    "_TOTAL_FINAL": total_final,
                })

    if unknown_pdf_rows:
        st.write("### IDs novos do PDF")
        st.caption("Só entram no fechamento do Demetra se você marcar Demetra.")
        for idx, row in enumerate(unknown_pdf_rows):
            with st.container(border=True):
                st.code(row["linha"])
                c1, c2 = st.columns(2)
                with c1:
                    cliente = st.selectbox(
                        f"Agência do ID {row['id_agente']}",
                        ["Demetra", "Alex", "Oscar", "Outro"],
                        key=f"cliente_pdf_{idx}",
                    )
                with c2:
                    rb = st.number_input(
                        f"%RB do ID {row['id_agente']}",
                        min_value=0.0,
                        max_value=100.0,
                        value=76.0 if cliente == "Demetra" else 0.0,
                        step=1.0,
                        key=f"rb_pdf_{idx}",
                    )
                if cliente == "Demetra":
                    rb_valor, total_base, rebate, total_final = calc_row(float(row["ganhos"]), float(row["rake"]), float(rb))
                    pdf_rows.append({
                        "AGENTE": f"ID {row['id_agente']}",
                        "GANHOS": float(row["ganhos"]),
                        "RAKE": float(row["rake"]),
                        "RB": f"{int(float(rb))}%",
                        "TOTAL": total_base,
                        "_REBATE": rebate,
                        "_TOTAL_FINAL": total_final,
                    })

    if st.button("Gerar fechamento Demetra", type="primary", key="btn_demetra"):
        all_rows = excel_rows + pdf_rows
        if not all_rows:
            st.warning("Envie a planilha e/ou o PDF.")
            return

        detalhado = pd.DataFrame(all_rows)
        st.write("### Prévia dos dados")
        preview = detalhado.copy()
        for col in ["GANHOS", "RAKE", "TOTAL"]:
            preview[col] = preview[col].apply(fmt_brl)
        st.dataframe(preview[["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]], use_container_width=True)

        rebate_total = detalhado["_REBATE"].sum()
        total_geral = detalhado["_TOTAL_FINAL"].sum()

        if not periodo.strip():
            st.warning("Preencha o período para gerar a imagem final.")
            return

        report = generate_demetra_table_image(periodo.strip(), detalhado[["AGENTE", "GANHOS", "RAKE", "RB", "TOTAL"]], total_geral, rebate_total)
        st.write("### Relatório final em imagem")
        st.image(report, caption="Pronto para print", use_container_width=True)
        st.download_button(
            "Baixar relatório em PNG",
            data=to_png_bytes(report),
            file_name="demetra_fechamento.png",
            mime="image/png",
        )


st.title("Fechamentos Premier")
cliente = st.selectbox("Escolha o cliente", ["Harnefer", "Demetra"])

if cliente == "Harnefer":
    page_harnefer()
else:
    page_demetra()
