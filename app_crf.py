import re
import time
import traceback
from datetime import datetime
from pathlib import Path
from PIL import Image
import pytesseract

import pandas as pd
import fitz  # PyMuPDF
from loguru import logger
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from openpyxl import load_workbook

# === Configurações ===
PLANILHA = Path("base_certidoes.xlsx")
ABA = "CRF"
COL_CNPJ = "CNPJ"
COL_VALIDADE = "VALIDADE CERTIDÃO"
COL_RAZAO = "RAZÃO SOCIAL"
OUTPUT_DIR = Path("certidoes_baixadas") / ABA.replace(" ", "_")
URL_CRF = "https://consulta-crf.caixa.gov.br/consultacrf/pages/consultaEmpregador.jsf"
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

# === Função principal ===
def processar_pmm():
    logger.add("execucaopmm.log", rotation="1 MB")
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
                page.goto(URL_CRF, timeout=TIMEOUT)
                page.get_by_role("radio", name="CNPJ").click()
                page.wait_for_load_state("networkidle", timeout=15000)
                
                # Preencher o número do CNPJ
                page.get_by_label("Insira o Número").fill(cnpj_limpo)

                # Captura imagem do captcha
                captcha_element = page.locator("xpath=//img[contains(@src, 'captcha')]")
                captcha_path = OUTPUT_DIR / f"captcha_{cnpj_limpo}.png"
                captcha_element.screenshot(path=str(captcha_path))

                # Leitura do captcha com pytesseract
                imagem = Image.open(captcha_path)
                texto_captcha = pytesseract.image_to_string(imagem, config='--psm 8').strip()
                logger.info(f"Captcha lido: '{texto_captcha}'")

                # Preencher captcha
                page.get_by_label("Insira o código").fill(texto_captcha)

                # === Clicar em "Consultar" e capturar nova aba com o PDF ===
                with context.expect_page() as nova_pagina_evento:
                    page.get_by_role("button", name="Consultar").click()

                nova_aba = nova_pagina_evento.value
                nova_aba.wait_for_load_state("networkidle", timeout=15000)

                # Salvar PDF renderizado da nova aba
                temp_path = OUTPUT_DIR / f"temp_{cnpj_limpo}.pdf"
                nova_aba.pdf(path=str(temp_path), format="A4")

                # Extrair validade e salvar com nome correto
                validade = extrair_validade_pdf(temp_path)

                if validade:
                    salvar_valor_na_planilha(cnpj_limpo, validade, PLANILHA, ABA)
                    validade_formatada = datetime.strptime(validade, "%d/%m/%Y").strftime("%Y%m%d")
                    nome_arquivo = f"sefaz_n_contribuinte_{cnpj_limpo}_{validade_formatada}.pdf"
                    destino_pdf = OUTPUT_DIR / nome_arquivo
                    temp_path.rename(destino_pdf)
                    logger.success(f"{cnpj_limpo} → Sucesso: validade {validade}")
                else:
                    destino_pdf = OUTPUT_DIR / f"erro_{cnpj_limpo}.pdf"
                    temp_path.rename(destino_pdf)
                    logger.warning(f"{cnpj_limpo} → Não foi possível extrair validade.")

                nova_aba.close()

            except Exception as e:
                motivo = f"{type(e).__name__}: {e}"
                logger.error(f"{cnpj_limpo} → ERRO: {motivo}")
                traceback.print_exc()

        context.close()
        browser.close()

    logger.info("Processo concluído.")

# === Execução ===
if __name__ == "__main__":
    processar_pmm()