# APP COM QUEBRA DE LINHA AUTOMÁTICA (WORD WRAP)
# Inclui: Harnefer, Demetra e Oscar

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
except:
    TESSERACT_OK = False


# ================= CONFIG =================
RB_HARNEFER = 76.0
RB_DEMETRA_PLANILHA = 70.0
REBATE_DEMETRA = -5.0
REBATE_OSCAR = -10.0

AGENTE_PLANILHA_DEMETRA = "TheShark_ (ID11719117)"
ID_PLANILHA = "11719117"
CODIGO = "802606"

MAPA_IDS = {
    "13472941": {"cliente": "Oscar", "rb": 65.0},
}

# ================= UTIL =================
def fmt(v):
    s = f"{float(v):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def parse_money(v):
    return float(str(v).replace(".", "").replace(",", ".").replace("R$", "").strip() or 0)

def get_font(size, bold=False):
    path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    return ImageFont.truetype(path, size=size)

def measure(draw, text, font):
    box = draw.textbbox((0,0), text, font=font)
    return box[2]-box[0], box[3]-box[1]

def wrap_text(draw, text, font, max_width):
    words = text.split()
    lines = []
    current = ""

    for word in words:
        test = current + (" " if current else "") + word
        w = measure(draw, test, font)[0]

        if w <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines


# ================= CALC =================
def calc(ganhos, rake, rb, rebate):
    rb_val = rake * (rb/100)
    base = ganhos + rb_val
    rebate_val = base * (rebate/100) if base > 0 else 0
    final = base + rebate_val
    return base, rebate_val, final


# ================= PDF =================
def read_pdf(file, cliente):
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                if "R$" not in line:
                    continue

                id_match = re.search(r"(\d{6,9})", line)
                if not id_match:
                    continue

                id_ = id_match.group(1)
                info = MAPA_IDS.get(id_)

                if not info or info["cliente"] != cliente:
                    continue

                vals = re.findall(r"R\$\s*-?[\d\.,]+", line)
                if len(vals) < 2:
                    continue

                ganhos = parse_money(vals[0])
                rake = parse_money(vals[1])
                nome = line.split(id_)[0].strip()

                rows.append({
                    "AGENTE": nome,
                    "GANHOS": ganhos,
                    "RAKE": rake,
                    "RB": info["rb"]
                })
    return pd.DataFrame(rows)


# ================= IMAGEM =================
def gerar_tabela(titulo, periodo, df, rebate_label, rebate_total, total_final):
    W, H = 1400, 900
    img = Image.new("RGB", (W,H), (245,245,245))
    draw = ImageDraw.Draw(img)

    title_font = get_font(44, True)
    sub_font = get_font(34)
    header_font = get_font(28, True)
    cell_font = get_font(26)

    # título
    tw,_ = measure(draw, titulo, title_font)
    draw.text(((W-tw)/2, 30), titulo, font=title_font, fill=(0,0,80))

    draw.text((W/2-200,100), "Período:", font=sub_font, fill=(0,0,80))
    draw.text((W/2-20,100), periodo, font=sub_font, fill=(0,0,80))

    # tabela
    x = 50
    col_w = [500,200,200,120,200]
    headers = ["AGENTE","GANHOS","RAKE","RB","TOTAL"]

    y = 180
    for i,h in enumerate(headers):
        draw.text((x+sum(col_w[:i])+10,y),h,font=header_font,fill=(0,0,0))

    y += 40

    for _,row in df.iterrows():
        max_w = col_w[0]-20
        lines = wrap_text(draw, row["AGENTE"], cell_font, max_w)

        line_h = measure(draw,"A",cell_font)[1]+4
        row_h = max(40, len(lines)*line_h)

        # AGENTE (quebrado)
        for j,l in enumerate(lines):
            draw.text((x+10,y+j*line_h),l,font=cell_font,fill=(0,0,0))

        # outras colunas
        draw.text((x+col_w[0],y),fmt(row["GANHOS"]),font=cell_font)
        draw.text((x+col_w[0]+col_w[1],y),fmt(row["RAKE"]),font=cell_font)
        draw.text((x+col_w[0]+col_w[1]+col_w[2],y),str(int(row["RB"]))+"%",font=cell_font)

        total = row["GANHOS"] + row["RAKE"]*(row["RB"]/100)
        draw.text((x+col_w[0]+col_w[1]+col_w[2]+col_w[3],y),fmt(total),font=cell_font)

        y += row_h

    # resumo
    y += 20
    draw.text((900,y),"TOTAL",font=header_font)
    draw.text((1100,y),fmt(total_final),font=header_font)

    y += 40
    draw.text((800,y),rebate_label,font=cell_font)
    draw.text((1100,y),fmt(rebate_total),font=cell_font)

    y += 50
    draw.text((700,y),"PREMIER TEM A PAGAR" if total_final>0 else "PREMIER TEM A RECEBER",font=header_font)
    draw.text((1100,y),fmt(abs(total_final)),font=header_font,fill=(0,120,0) if total_final>0 else (180,0,0))

    return img


# ================= PAGE OSCAR =================
def page_oscar():
    st.title("Oscar")

    periodo = st.text_input("Período")
    pdf = st.file_uploader("PDF", type=["pdf"])

    if st.button("Gerar"):
        df = read_pdf(pdf,"Oscar")

        rows=[]
        for _,r in df.iterrows():
            base, rebate, final = calc(r["GANHOS"],r["RAKE"],r["RB"],REBATE_OSCAR)
            rows.append({
                "AGENTE": r["AGENTE"],
                "GANHOS": r["GANHOS"],
                "RAKE": r["RAKE"],
                "RB": r["RB"],
                "FINAL": final,
                "REBATE": rebate
            })

        df2 = pd.DataFrame(rows)

        total = df2["FINAL"].sum()
        rebate_total = df2["REBATE"].sum()

        img = gerar_tabela("OSCAR", periodo, df2, "-10% total", rebate_total, total)

        st.image(img)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        st.download_button("Download", buf.getvalue(), "oscar.png")


# ================= MAIN =================
cliente = st.selectbox("Cliente",["Oscar"])

if cliente == "Oscar":
    page_oscar()
