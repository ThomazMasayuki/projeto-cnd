import re
import time
import traceback
from datetime import datetime
from pathlib import Path
from PIL import Image
import pytesseract
import time
import cv2
import numpy as np

import pandas as pd
import fitz  # PyMuPDF
from loguru import logger
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from openpyxl import load_workbook

# === Configurações ===
PLANILHA = Path("base_certidoes.xlsx")
ABA = "CDT"
COL_CNPJ = "CNPJ"
COL_VALIDADE = "VALIDADE CERTIDÃO"
COL_RAZAO = "RAZÃO SOCIAL"
OUTPUT_DIR = Path("certidoes_baixadas") / ABA.replace(" ", "_")
URL_CDT = "https://cndt-certidao.tst.jus.br/gerarCertidao.faces"
TIMEOUT = 40_000
REGEX_VALIDADE = r"VÁLIDA ATÉ:\s*(\d{2}/\d{2}/\d{4})"

# === Funções utilitárias ===
def normalizar_cnpj(cnpj: str) -> str:
    return re.sub(r"\D", "", str(cnpj)).zfill(14)

def extrair_validade_pdf(caminho_pdf: Path) -> str:
    try:
        with fitz.open(caminho_pdf) as doc:
            texto = "\n".join(page.get_text() for page in doc)
        match = re.search(REGEX_VALIDADE, texto, re.IGNORECASE)
        if match:
            return match.group(1)
    except Exception as e:
        logger.warning(f"Erro ao ler validade do PDF {caminho_pdf.name}: {e}")
    return ""

def salvar_valor_na_planilha(cnpj: str, nova_data: str, caminho: Path, aba: str):
    wb = load_workbook(caminho)
    ws = wb[aba]
    colunas = {cell.value: idx for idx, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1)), start=1)}
    idx_cnpj = colunas.get(COL_CNPJ)
    idx_validade = colunas.get(COL_VALIDADE)

    if not idx_cnpj or not idx_validade:
        logger.error("Coluna CNPJ ou VALIDADE CERTIDAO não encontrada.")
        return

    for row in ws.iter_rows(min_row=2):
        val = str(row[idx_cnpj - 1].value)
        if normalizar_cnpj(val) == normalizar_cnpj(cnpj):
            row[idx_validade - 1].value = nova_data
            break

    wb.save(caminho)
    wb.close()

def carregar_imagem_com_checagem(caminho_imagem: Path) -> np.ndarray:
    if not caminho_imagem.exists():
        raise FileNotFoundError(f"Imagem não encontrada: {caminho_imagem}")
    
    img = cv2.imread(str(caminho_imagem), cv2.IMREAD_COLOR)

    if img is None or img.size == 0:
        raise ValueError(f"Falha ao carregar imagem: {caminho_imagem}")
    
    return img

def preprocessar_imagem_para_ocr(caminho_imagem: Path) -> Image:
    img = carregar_imagem_com_checagem(caminho_imagem)

    # Aumenta escala da imagem
    scale_percent = 250
    width = int(img.shape[1] * scale_percent / 100)
    height = int(img.shape[0] * scale_percent / 100)
    img = cv2.resize(img, (width, height), interpolation=cv2.INTER_LINEAR)

    # Converte para escala de cinza
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Aplica leve desfoque
    blur = cv2.GaussianBlur(gray, (3, 3), 0)

    # Aplica binarização adaptativa
    binarizada = cv2.adaptiveThreshold(
        blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 11, 2
    )

    # Converte para PIL e retorna
    return Image.fromarray(binarizada)

def processar_cdt():
    logger.add("execucaocdt.log", rotation="1 MB")
    logger.info(f"Iniciando automação da aba: {ABA}")

    df = pd.read_excel(PLANILHA, sheet_name=ABA, dtype=str)
    df = df[[COL_RAZAO, COL_CNPJ, COL_VALIDADE]].dropna(subset=[COL_CNPJ])
    cnpjs = df[COL_CNPJ].drop_duplicates().tolist()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        for cnpj in cnpjs:
            cnpj_limpo = normalizar_cnpj(cnpj)

            try:
                logger.info(f"Consultando CNPJ: {cnpj_limpo}")
                page.goto(URL_CDT, timeout=TIMEOUT)
                page.wait_for_load_state("domcontentloaded", timeout=10000)

                # Preenche o CNPJ
                page.get_by_role("textbox", name="Registro no Cadastro Nacional").fill(cnpj_limpo)

                time.sleep(2)  # espera para o captcha carregar corretamente

                # Captura o captcha
                captcha_img = page.get_by_role("img", name="Captcha para permitir a")
                captcha_path = OUTPUT_DIR / f"captcha_{cnpj_limpo}.png"
                captcha_img.screenshot(path=str(captcha_path))

                time.sleep(1.5)  # pequena espera para garantir que a imagem foi salva

                # Resolve o captcha com pytesseract
                imagem = preprocessar_imagem_para_ocr(captcha_path)
                texto_captcha = pytesseract.image_to_string(imagem, config='--psm 8 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789').strip()
                logger.info(f"Captcha lido: '{texto_captcha}'")

                # Preenche o captcha
                page.get_by_role("textbox", name="* Digite os caracteres").fill(texto_captcha)

                time.sleep(1.5)  # garante que o campo foi preenchido

                # Clica no botão "Emitir Certidão"
                with context.expect_page() as nova_aba_evento:
                    page.get_by_role("button", name="Emitir Certidão").click()

                nova_aba = nova_aba_evento.value
                nova_aba.wait_for_load_state("networkidle", timeout=15000)

                # Salva o PDF da nova aba
                temp_path = OUTPUT_DIR / f"temp_{cnpj_limpo}.pdf"
                nova_aba.pdf(path=str(temp_path), format="A4")

                # Salva print da nova aba
                screenshot_path = OUTPUT_DIR / f"screenshot_{cnpj_limpo}.png"
                nova_aba.screenshot(path=str(screenshot_path), full_page=True)

                validade = extrair_validade_pdf(temp_path)

                if validade:
                    salvar_valor_na_planilha(cnpj_limpo, validade, PLANILHA, ABA)
                    validade_formatada = datetime.strptime(validade, "%d/%m/%Y").strftime("%Y%m%d")
                    nome_arquivo = f"cdt_{cnpj_limpo}_{validade_formatada}.pdf"
                    destino_pdf = OUTPUT_DIR / nome_arquivo
                    temp_path.rename(destino_pdf)
                    logger.success(f"{cnpj_limpo} → Sucesso: validade {validade}")
                else:
                    destino_pdf = OUTPUT_DIR / f"erro_{cnpj_limpo}.pdf"
                    temp_path.rename(destino_pdf)
                    logger.warning(f"{cnpj_limpo} → Não foi possível extrair validade.")

                nova_aba.close()
                page.bring_to_front()

            except Exception as e:
                motivo = f"{type(e).__name__}: {e}"
                logger.error(f"{cnpj_limpo} → ERRO: {motivo}")
                traceback.print_exc()

        context.close()
        browser.close()

    logger.info("Processo concluído.")

# === Execução ===
if __name__ == "__main__":
    processar_cdt()