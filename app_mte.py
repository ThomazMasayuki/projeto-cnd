import re
import time
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import fitz  # PyMuPDF
from loguru import logger
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from openpyxl import load_workbook

# === Configurações ===
PLANILHA = Path("base_certidoes.xlsx")
ABA = "MTE"
COL_CNPJ = "CNPJ"
COL_VALIDADE = "VALIDADE CERTIDÃO"
COL_RAZAO = "RAZÃO SOCIAL"
OUTPUT_DIR = Path("certidoes_baixadas") / ABA.replace(" ", "_")
URL_MTE = "https://eprocesso.sit.trabalho.gov.br/Entrar?ReturnUrl=%2FCertidao%2FEmitir"
TIMEOUT = 40_000
REGEX_VALIDADE = r"Válida até:\s*(\d{2}/\d{2}/\d{4})"

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
    
def processar_mte():
    logger.add("execucao_mte.log", rotation="1 MB")
    logger.info(f"Iniciando automação GOV.BR no MTE")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        try:
            logger.info(f"Acessando a página inicial do MTE...")
            page.goto(URL_MTE, timeout=TIMEOUT)

            # Clica no botão "Entrar com GOV.BR"
            logger.info("Clicando em 'Entrar com GOV.BR'")
            page.locator('#janela-login-gov-br').get_by_role("link", name="Entrar com GOV.BR").click()

            # Preenche o CPF e clica em "Continuar"
            logger.info("Preenchendo CPF...")
            page.get_by_role("textbox", name="Digite seu CPF").fill("02448982198")
            page.get_by_role("button", name="Continuar").click()

            # Aguarda a tela da senha ou token MFA manual
            logger.info("Aguardando usuário inserir a senha ou fazer autenticação manual...")

            # Espera indefinidamente ou pode usar input para continuar manualmente
            input("Insira a senha e complete o login no navegador, depois pressione ENTER para encerrar.")

        except Exception as e:
            motivo = f"{type(e).__name__}: {e}"
            logger.error(f"Erro durante processo de login: {motivo}")
            traceback.print_exc()

        context.close()
        browser.close()

    logger.info("Processo concluído.")

# === Execução ===
if __name__ == "__main__":
    processar_mte()