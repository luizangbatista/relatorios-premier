
import io
import re
from pathlib import Path

import pandas as pd
import pdfplumber
import pytesseract
import streamlit as st
from PIL import Image, ImageEnhance, ImageFilter, ImageOps


st.set_page_config(page_title="Fechamentos Premier", layout="wide")


MAPA_IDS = {
    "12970177": {"cliente": "Alex", "rb": 70.0},
    "12968708": {"cliente": "Demetra", "rb": 70.0},
    "13019559": {"cliente": "Alex", "rb": 50.0},
    "1607968": {"cliente": "Demetra", "rb": 65.0},
    "1527106": {"cliente": "Demetra", "rb": 65.0},
    "13018880": {"cliente": "Alex", "rb": 45.0},
    "13213751": {"cliente": "Alex", "rb": 70.0},
    "13265647": {"cliente": "Alex", "rb": 50.0},
    "13319248": {"cliente": "Alex", "rb": 70.0},
    "13357678": {"cliente": "Demetra", "rb": 65.0},
    "13379845": {"cliente": "Alex", "rb": 60.0},
    "13104440": {"cliente": "Alex", "rb": 50.0},
    "13472941": {"cliente": "Oscar", "rb": 65.0},
}

CLIENTES_FIXOS = ["Demetra", "Alex", "Oscar", "Harnefer", "Outro"]


def fmt_brl(value: float) -> str:
    s = f"{value:,.2f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def parse_money(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    text = text.replace("R$", "").replace(" ", "").replace(")", "").replace("(", "")
    text = text.replace("—", "-").replace("–", "-").replace("−", "-")

    # Casos BR: 1.234,56
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")

    # Casos BR simples: 1234,56
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")

    # Se sobrar mais de um ponto, tratar como milhar + decimal
    if text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]

    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return 0.0

    try:
        return float(text)
    except ValueError:
        return 0.0


def preprocess_image(image: Image.Image) -> Image.Image:
    img = image.convert("L")
    img = ImageOps.autocontrast(img)
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = img.filter(ImageFilter.SHARPEN)
    return img


def run_ocr(image: Image.Image) -> str:
    img = preprocess_image(image)
    text = pytesseract.image_to_string(img, lang="eng")
    return text or ""


def detect_image_kind(text: str) -> str:
    up = text.upper()
    if "CLUB DATA" in up and "TOTAL FEE" in up and "WINNINGS" in up:
        return "harnefer"
    if "SUPER AGENTE" in up and "KILLUMINATTI" in up:
        return "killuminatti_resumo"
    if "GUSTAVO" in up and "CASARICA" in up:
        return "alex_casarica"
    return "desconhecida"


def find_money_after_label(text: str, label: str):
    pattern = rf"{label}\s*[:\-]?\s*([\-]?\s*[R$]*\s*[0-9\.,]+)"
    m = re.search(pattern, text, flags=re.IGNORECASE)
    if m:
        return parse_money(m.group(1))
    return None


def extract_harnefer(text: str) -> dict:
    # OCR típico da imagem: "... Total Fee ... 1,004.85 ... Winnings ... 1,470.15"
    fee_match = re.search(r"TOTAL\s*FEE[^0-9\-]*([\-]?[0-9][0-9\.,]+)", text, flags=re.IGNORECASE)
    winnings_match = re.search(r"WINNINGS[^0-9\-]*([\-]?[0-9][0-9\.,]+)", text, flags=re.IGNORECASE)

    rake = parse_money(fee_match.group(1)) if fee_match else 0.0
    ganhos = parse_money(winnings_match.group(1)) if winnings_match else 0.0

    return {
        "cliente": "Harnefer",
        "origem": "Imagem Harnefer",
        "ganhos": ganhos,
        "rake": rake,
        "rb_percentual": 76.0,
        "observacao": "Rake puxado de Total Fee e ganhos de Winnings.",
    }


def extract_killuminatti_summary(text: str) -> dict:
    # Busca a linha do Killuminatti
    clean = text.replace("\n", " ")
    m = re.search(
        r"KILLUMINATTI.*?([0-9\.,]+)\s+([\-]?[0-9\.,]+)\s+([0-9\.,]+)\s+([0-9\.,]+)",
        clean,
        flags=re.IGNORECASE,
    )
    rake = ganhos = 0.0

    if m:
        rake = parse_money(m.group(1))
        ganhos = parse_money(m.group(3))
    else:
        # fallback por rótulos gerais
        rake = find_money_after_label(text, "Rake") or 0.0
        ganhos = find_money_after_label(text, "Ganhos") or 0.0

    return {
        "cliente": "Demetra",
        "origem": "Imagem Killuminatti",
        "ganhos": ganhos,
        "rake": rake,
        "rb_percentual": 0.0,  # usuário pode ajustar se quiser
        "observacao": "Resumo do super agente Killuminatti.",
    }


def extract_alex_casarica(text: str) -> dict:
    # Exemplo:
    # PROFIT/LOSS -4.277,35
    # RAKE 70% 1.225,95
    profit_match = re.search(r"PROFIT\s*/\s*LOSS[^0-9\-]*([\-]?[0-9\.,]+)", text, flags=re.IGNORECASE)
    rake70_match = re.search(r"RAKE\s*70%[^0-9\-]*([\-]?[0-9\.,]+)", text, flags=re.IGNORECASE)
    saldo_final_match = re.search(r"SALDO\s*FINAL[^0-9\-]*([\-]?[0-9\.,]+)", text, flags=re.IGNORECASE)

    ganhos = parse_money(profit_match.group(1)) if profit_match else 0.0
    rake_70 = parse_money(rake70_match.group(1)) if rake70_match else 0.0
    saldo_final_lido = parse_money(saldo_final_match.group(1)) if saldo_final_match else 0.0

    # Regra passada pela usuária:
    # imagem mostra 70% do rake; contrato real do agente = 65%
    rakeback = rake_70 * (65.0 / 70.0) if rake_70 else 0.0

    return {
        "cliente": "Alex",
        "origem": "Imagem Gustavo | Casarica",
        "ganhos": ganhos,
        "rake": rake_70,
        "rb_percentual": 65.0,
        "rakeback_forcado": rakeback,
        "saldo_final_lido": saldo_final_lido,
        "observacao": "Na imagem do Alex, o valor de RAKE 70% é convertido para 65/70.",
    }


def process_excel_2101(file_obj, nome_cliente_outro: str, rb_outro: float) -> pd.DataFrame:
    df = pd.read_excel(file_obj, usecols=[5, 6, 7, 8, 9, 29])
    df.columns = ["origem_linha", "id_conta", "nick", "codigo", "ganhos", "rake"]

    df = df[df["codigo"].astype(str).str.strip() == "802606"].copy()
    if df.empty:
        return pd.DataFrame(columns=["cliente", "origem", "ganhos", "rake", "rb_percentual", "observacao"])

    df["origem_linha"] = df["origem_linha"].astype(str).str.strip()
    df["id_conta"] = df["id_conta"].astype(str).str.strip()
    df["nick"] = df["nick"].astype(str).str.strip()
    df["ganhos"] = pd.to_numeric(df["ganhos"], errors="coerce").fillna(0.0)
    df["rake"] = pd.to_numeric(df["rake"], errors="coerce").fillna(0.0)

    rows = []

    mask_demetra = (df["id_conta"] == "11719117") & (df["nick"].str.lower() == "killuminatti")
    demetra_df = df[mask_demetra]
    if not demetra_df.empty:
        rows.append({
            "cliente": "Demetra",
            "origem": "Planilha 2101",
            "ganhos": demetra_df["ganhos"].sum(),
            "rake": demetra_df["rake"].sum(),
            "rb_percentual": 0.0,
            "observacao": "Agrupado da planilha pelo Killuminatti / 11719117.",
        })

    mask_outro = df["origem_linha"].str.upper() == "PPFICHAS"
    outro_df = df[mask_outro]
    if not outro_df.empty:
        rows.append({
            "cliente": nome_cliente_outro.strip() if nome_cliente_outro.strip() else "Outro",
            "origem": "Planilha 2101",
            "ganhos": outro_df["ganhos"].sum(),
            "rake": outro_df["rake"].sum(),
            "rb_percentual": float(rb_outro or 0.0),
            "observacao": "Linhas PPFICHAS agrupadas.",
        })

    return pd.DataFrame(rows)


def extract_pdf_rows(file_obj) -> tuple[pd.DataFrame, list[dict]]:
    text_parts = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)

    full_text = "\n".join(text_parts)
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

    resultados = []
    ids_desconhecidos = []

    for line in lines:
        if "R$" not in line:
            continue

        id_match = re.search(r"\b(\d{6,9})\b", line)
        if not id_match:
            continue

        id_agente = id_match.group(1)
        money_matches = re.findall(r"R\$\s*-?\d[\d\.,]*", line)

        # Esperado:
        # ganhos, rake, rebate, rakeback
        if len(money_matches) < 2:
            continue

        ganhos = parse_money(money_matches[0])
        rake = parse_money(money_matches[1])

        info = MAPA_IDS.get(id_agente)
        if info is None:
            ids_desconhecidos.append({"id_agente": id_agente, "linha": line, "ganhos": ganhos, "rake": rake})
            continue

        resultados.append({
            "cliente": info["cliente"],
            "origem": "PDF",
            "ganhos": ganhos,
            "rake": rake,
            "rb_percentual": float(info["rb"]),
            "id_agente": id_agente,
            "observacao": f"ID {id_agente} vindo do PDF.",
        })

    df = pd.DataFrame(resultados)
    if not df.empty:
        df = (
            df.groupby(["cliente", "origem", "rb_percentual"], as_index=False)[["ganhos", "rake"]]
            .sum()
        )
        df["observacao"] = "Agrupado por cliente a partir do PDF."

    return df, ids_desconhecidos


def calcular_rebate(cliente: str, total_base: float) -> float:
    cliente = str(cliente).strip().lower()

    if cliente == "demetra":
        return total_base * (-0.05) if total_base > 0 else 0.0

    if cliente == "alex":
        # positivo -> rebate negativo; negativo -> rebate positivo
        return total_base * (-0.05)

    if cliente == "oscar":
        return total_base * (-0.10) if total_base > 0 else 0.0

    return 0.0


def finalizar_fechamento(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    out = df.copy()

    if "rakeback_forcado" not in out.columns:
        out["rakeback_forcado"] = None

    out["ganhos"] = out["ganhos"].astype(float)
    out["rake"] = out["rake"].astype(float)
    out["rb_percentual"] = out["rb_percentual"].astype(float)

    out["rakeback"] = out.apply(
        lambda r: float(r["rakeback_forcado"]) if pd.notna(r["rakeback_forcado"]) else float(r["rake"]) * (float(r["rb_percentual"]) / 100.0),
        axis=1,
    )
    out["total_base"] = out["ganhos"] + out["rakeback"]
    out["rebate"] = out.apply(lambda r: calcular_rebate(r["cliente"], r["total_base"]), axis=1)
    out["total_final"] = out["total_base"] + out["rebate"]

    def situacao(v):
        if v > 0:
            return f"Premier tem a pagar {fmt_brl(v)}"
        if v < 0:
            return f"Premier tem a receber {fmt_brl(abs(v))}"
        return "Sem valores a pagar ou receber"

    out["situacao"] = out["total_final"].apply(situacao)
    return out


def dataframe_download_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    export = df.copy()

    for col in ["ganhos", "rake", "rb_percentual", "rakeback", "total_base", "rebate", "total_final"]:
        if col in export.columns:
            export[col] = export[col].astype(float)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, index=False, sheet_name="fechamentos")
    return output.getvalue()


st.title("Fechamentos Premier")
st.caption("Planilha + PDF + imagem com regras de Demetra, Alex, Oscar, Harnefer e Outro.")

with st.sidebar:
    st.header("Configurações")
    nome_cliente_outro = st.text_input("Nome do cliente para PPFICHAS / Outro", value="Outro")
    rb_outro = st.number_input("%RB para PPFICHAS / Outro", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
    rb_demetra_planilha = st.number_input("%RB padrão do Demetra na planilha/imagem Killuminatti", min_value=0.0, max_value=100.0, value=76.0, step=1.0)
    usar_imagem_killuminatti_se_existir = st.checkbox("Preferir imagem Killuminatti para Demetra quando enviada", value=True)

arquivos = st.file_uploader(
    "Envie planilhas, PDFs e imagens",
    type=["xlsx", "xls", "pdf", "png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
)

resultado_partes = []
pendencias_ids = []

if arquivos:
    for arquivo in arquivos:
        nome = arquivo.name
        ext = Path(nome).suffix.lower()

        st.subheader(f"Arquivo: {nome}")

        if ext in [".xlsx", ".xls"]:
            if Path(nome).name.startswith("2101"):
                try:
                    df_plan = process_excel_2101(arquivo, nome_cliente_outro, rb_outro)
                    if not df_plan.empty:
                        df_plan.loc[df_plan["cliente"] == "Demetra", "rb_percentual"] = rb_demetra_planilha
                        resultado_partes.append(df_plan)
                        st.success("Planilha 2101 processada.")
                        st.dataframe(df_plan, use_container_width=True)
                    else:
                        st.warning("Planilha lida, mas sem linhas após o filtro I = 802606.")
                except Exception as e:
                    st.error(f"Erro ao ler a planilha: {e}")
            else:
                st.info("Planilha enviada não começa com 2101. Nenhuma regra específica aplicada.")

        elif ext == ".pdf":
            try:
                df_pdf, ids_novos = extract_pdf_rows(arquivo)
                if not df_pdf.empty:
                    resultado_partes.append(df_pdf)
                    st.success("PDF processado.")
                    st.dataframe(df_pdf, use_container_width=True)
                else:
                    st.warning("PDF lido, mas nenhuma linha válida foi identificada.")

                if ids_novos:
                    st.info("Há IDs não mapeados no PDF. Preencha abaixo.")
                    for item in ids_novos:
                        pendencias_ids.append(item)
            except Exception as e:
                st.error(f"Erro ao ler o PDF: {e}")

        elif ext in [".png", ".jpg", ".jpeg", ".webp"]:
            try:
                image = Image.open(arquivo)
                st.image(image, caption=nome, width=350)
                ocr_text = run_ocr(image)
                kind = detect_image_kind(ocr_text)

                st.caption(f"Tipo detectado: {kind}")

                if kind == "harnefer":
                    row = extract_harnefer(ocr_text)
                    resultado_partes.append(pd.DataFrame([row]))
                    st.success("Imagem do Harnefer processada.")
                    st.text_area("OCR extraído", ocr_text, height=180)

                elif kind == "killuminatti_resumo":
                    row = extract_killuminatti_summary(ocr_text)
                    row["rb_percentual"] = rb_demetra_planilha
                    if usar_imagem_killuminatti_se_existir:
                        # remove outras partes da Demetra depois, na consolidação, se desejar
                        row["observacao"] += " Preferência de uso ativada."
                    resultado_partes.append(pd.DataFrame([row]))
                    st.success("Imagem do Killuminatti processada.")
                    st.text_area("OCR extraído", ocr_text, height=180)

                elif kind == "alex_casarica":
                    row = extract_alex_casarica(ocr_text)
                    resultado_partes.append(pd.DataFrame([row]))
                    st.success("Imagem do Alex (Gustavo | Casarica) processada.")
                    st.text_area("OCR extraído", ocr_text, height=180)

                else:
                    st.warning("Não consegui classificar automaticamente a imagem. Use o OCR abaixo para revisão manual.")
                    st.text_area("OCR extraído", ocr_text, height=220)

            except Exception as e:
                st.error(f"Erro ao ler a imagem: {e}")

    if pendencias_ids:
        st.header("IDs novos encontrados no PDF")
        novos_rows = []
        for item in pendencias_ids:
            st.markdown(f"**ID {item['id_agente']}** — linha: `{item['linha']}`")
            col1, col2 = st.columns(2)
            with col1:
                cliente_novo = st.selectbox(
                    f"Cliente do ID {item['id_agente']}",
                    ["Demetra", "Alex", "Oscar"],
                    key=f"cliente_{item['id_agente']}",
                )
            with col2:
                rb_novo = st.number_input(
                    f"%RB do ID {item['id_agente']}",
                    min_value=0.0,
                    max_value=100.0,
                    step=1.0,
                    key=f"rb_{item['id_agente']}",
                )

            novos_rows.append({
                "cliente": cliente_novo,
                "origem": "PDF",
                "ganhos": item["ganhos"],
                "rake": item["rake"],
                "rb_percentual": float(rb_novo),
                "observacao": f"ID novo {item['id_agente']} classificado manualmente.",
            })

        if novos_rows:
            resultado_partes.append(pd.DataFrame(novos_rows))

    if resultado_partes:
        bruto = pd.concat(resultado_partes, ignore_index=True, sort=False)

        # Preferência opcional: se houver imagem Killuminatti, remover outras linhas da Demetra
        if usar_imagem_killuminatti_se_existir:
            has_killu_img = ((bruto["cliente"] == "Demetra") & (bruto["origem"] == "Imagem Killuminatti")).any()
            if has_killu_img:
                bruto = bruto[~((bruto["cliente"] == "Demetra") & (bruto["origem"] != "Imagem Killuminatti"))].copy()

        consolidado = (
            bruto.groupby(["cliente", "origem", "rb_percentual"], as_index=False)
            .agg({
                "ganhos": "sum",
                "rake": "sum",
                "observacao": lambda x: " | ".join([str(v) for v in x if pd.notna(v) and str(v).strip()]),
                "rakeback_forcado": "max" if "rakeback_forcado" in bruto.columns else "first",
                "saldo_final_lido": "max" if "saldo_final_lido" in bruto.columns else "first",
            })
            if "rakeback_forcado" in bruto.columns or "saldo_final_lido" in bruto.columns
            else bruto.groupby(["cliente", "origem", "rb_percentual"], as_index=False)
                    .agg({"ganhos": "sum", "rake": "sum", "observacao": lambda x: " | ".join([str(v) for v in x if pd.notna(v) and str(v).strip()])})
        )

        final = finalizar_fechamento(consolidado)

        st.header("Fechamento consolidado")
        st.dataframe(final, use_container_width=True)

        st.header("Resumo por cliente")
        resumo = (
            final.groupby("cliente", as_index=False)[["ganhos", "rake", "rakeback", "total_base", "rebate", "total_final"]]
            .sum()
        )
        resumo["situacao"] = resumo["total_final"].apply(
            lambda v: f"Premier tem a pagar {fmt_brl(v)}" if v > 0 else (f"Premier tem a receber {fmt_brl(abs(v))}" if v < 0 else "Sem valores a pagar ou receber")
        )
        st.dataframe(resumo, use_container_width=True)

        excel_bytes = dataframe_download_excel(final)
        st.download_button(
            "Baixar fechamento em Excel",
            data=excel_bytes,
            file_name="fechamento_premier.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        for cliente in final["cliente"].dropna().unique():
            bloco = final[final["cliente"] == cliente].copy()
            with st.expander(f"Relatório - {cliente}", expanded=False):
                st.dataframe(bloco, use_container_width=True)
                total_cliente = bloco["total_final"].sum()
                if total_cliente > 0:
                    st.success(f"Premier tem a pagar {fmt_brl(total_cliente)}")
                elif total_cliente < 0:
                    st.warning(f"Premier tem a receber {fmt_brl(abs(total_cliente))}")
                else:
                    st.info("Sem valores a pagar ou receber.")
    else:
        st.info("Envie arquivos para gerar o fechamento.")
else:
    st.info("Aguardando uploads.")
