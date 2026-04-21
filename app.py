
import io
import re
from pathlib import Path

import streamlit as st
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps

try:
    import pytesseract
    TESSERACT_OK = True
except Exception:
    pytesseract = None
    TESSERACT_OK = False


st.set_page_config(page_title="Fechamentos Premier - Harnefer", layout="wide")


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
    rakeback = rake * 0.76
    total = ganhos + rakeback

    return {
        "ganhos": ganhos,
        "rake": rake,
        "rb_percentual": 76.0,
        "rakeback": rakeback,
        "total_final": total,
        "ocr_fee": texts[1],
        "ocr_winnings": texts[3],
        "crop": crop,
    }


def generate_harnefer_report_image(periodo: str, ganhos: float, rake: float, rb_percentual: float = 76.0) -> Image.Image:
    rb = rake * (rb_percentual / 100.0)
    total = ganhos + rb

    navy = (7, 29, 69)
    gold = (199, 143, 43)
    green = (0, 102, 45)
    gray = (248, 248, 248)
    light_green = (234, 241, 235)

    W, H = 1800, 1200
    img = Image.new("RGB", (W, H), gray)
    draw = ImageDraw.Draw(img)

    title_font = get_font(120, bold=True)
    subtitle_font = get_font(42, bold=False)
    subtitle_bold = get_font(42, bold=True)
    header_font = get_font(44, bold=True)
    value_font = get_font(50, bold=False)
    status_font = get_font(56, bold=True)
    status_value_font = get_font(68, bold=True)

    title = "HARNEFER"
    tw, _ = measure(draw, title, title_font)
    draw.text(((W - tw) / 2, 60), title, fill=navy, font=title_font)

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
    header_y1, header_y2 = 470, 630
    values_y1, values_y2 = 630, 885

    draw.rounded_rectangle((table_x1, header_y1, table_x2, values_y2), radius=20, outline=navy, width=5, fill="white")
    draw.rounded_rectangle((table_x1, header_y1, table_x2, header_y2), radius=20, fill=navy)
    draw.rectangle((table_x1, header_y2 - 20, table_x2, header_y2), fill=navy)

    cols = ["GANHOS", "RAKE", "RB (76%)", "TOTAL"]
    vals = [fmt_brl(ganhos), fmt_brl(rake), fmt_brl(rb), fmt_brl(total)]

    col_w = (table_x2 - table_x1) / 4
    for i in range(1, 4):
        x = table_x1 + i * col_w
        draw.line((x, header_y1, x, values_y2), fill=(160, 160, 160), width=2)

    for i, col in enumerate(cols):
        cx = table_x1 + i * col_w + col_w / 2
        cw, _ = measure(draw, col, header_font)
        draw.text((cx - cw / 2, header_y1 + 48), col, fill="white", font=header_font)

    for i, val in enumerate(vals):
        cx = table_x1 + i * col_w + col_w / 2
        vw, _ = measure(draw, val, value_font)
        draw.text((cx - vw / 2, values_y1 + 95), val, fill=navy, font=value_font)

    status_text = "PREMIER TEM A PAGAR" if total > 0 else ("PREMIER TEM A RECEBER" if total < 0 else "SEM VALORES")
    status_value = f"R$ {fmt_brl(abs(total))}"

    box_y1, box_y2 = 960, 1135
    draw.rounded_rectangle((70, box_y1, 1730, box_y2), radius=20, outline=navy, width=4, fill=light_green)

    draw.ellipse((120, 1008, 220, 1108), fill=navy)
    if total != 0:
        draw.line((150, 1058, 170, 1078), fill="white", width=8)
        draw.line((170, 1078, 205, 1038), fill="white", width=8)

    stw, _ = measure(draw, status_text, status_font)
    svw, _ = measure(draw, status_value, status_value_font)
    start_x = (W - (stw + 40 + svw)) / 2 + 40
    draw.text((start_x, 1018), status_text, fill=navy, font=status_font)
    draw.text((start_x + stw + 40, 1008), status_value, fill=green, font=status_value_font)

    return img


st.title("Fechamento por agência")
st.subheader("Harnefer")

periodo = st.text_input("Período do fechamento", placeholder="13/04/2026 a 19/04/2026")
arquivo = st.file_uploader("Envie a imagem do Harnefer", type=["png", "jpg", "jpeg", "webp"])

if not TESSERACT_OK:
    st.error("OCR indisponível: instale `pytesseract` e o programa `tesseract-ocr` no ambiente.")
    st.info("No Streamlit Cloud, crie um arquivo packages.txt com:")
    st.code("tesseract-ocr\ntesseract-ocr-eng")

if arquivo:
    img = Image.open(arquivo)
    st.image(img, caption="Imagem enviada", width=420)

    if st.button("Ler imagem e gerar fechamento", type="primary"):
        if not TESSERACT_OK:
            st.stop()

        if not detect_harnefer_image(img):
            st.warning("Não identifiquei a imagem do Harnefer com segurança. Ela precisa mostrar os cards com Total Fee e Winnings.")
        else:
            dados = extract_harnefer_values(img)

            st.write("### Valores lidos da imagem")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Ganhos", f"R$ {fmt_brl(dados['ganhos'])}")
            c2.metric("Rake", f"R$ {fmt_brl(dados['rake'])}")
            c3.metric("RB (76%)", f"R$ {fmt_brl(dados['rakeback'])}")
            c4.metric("Total", f"R$ {fmt_brl(dados['total_final'])}")

            with st.expander("Diagnóstico OCR", expanded=False):
                st.image(dados["crop"], caption="Recorte usado para leitura", width=700)
                st.text("OCR da coluna Total Fee:")
                st.code(dados["ocr_fee"])
                st.text("OCR da coluna Winnings:")
                st.code(dados["ocr_winnings"])

            if not periodo.strip():
                st.warning("Preencha o período do fechamento para gerar a imagem final.")
            else:
                report = generate_harnefer_report_image(
                    periodo=periodo.strip(),
                    ganhos=dados["ganhos"],
                    rake=dados["rake"],
                    rb_percentual=76.0,
                )
                st.write("### Relatório final em imagem")
                st.image(report, caption="Pronto para print", use_container_width=True)

                st.download_button(
                    "Baixar relatório em PNG",
                    data=to_png_bytes(report),
                    file_name="harnefer_fechamento.png",
                    mime="image/png",
                )
