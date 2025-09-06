from playwright.sync_api import sync_playwright
import time
import requests
import re
import time
import traceback
from datetime import datetime
from pathlib import Path
from PIL import Image
import pytesseract
import cv2
import numpy as np

import pandas as pd
import fitz  # PyMuPDF
from loguru import logger
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from openpyxl import load_workbook

# === Configurações ===
PLANILHA = Path("base_certidoes.xlsx")
ABA = "FALÊNCIA"
COL_CNPJ = "CNPJ"
COL_VALIDADE = "VALIDADE CERTIDÃO"
COL_RAZAO = "RAZÃO SOCIAL"
OUTPUT_DIR = Path("certidoes_baixadas") / ABA.replace(" ", "_")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

URL_SITE = 'https://consultasaj.tjam.jus.br/sco/abrirCadastro.do'
SITEKEY = '6LcnC3cdAAAAABWUEy-SzR8kMrk3FA9llI6hU934'
EMAIL_PADRAO = "adm@joyceassessoria.com"
API_KEY_2CAPTCHA = '30924a201f06a3b554d7479c487fee8e'
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

# === 1. Função para solicitar resolução do reCAPTCHA ===
def solicitar_captcha(api_key, sitekey, url):
    payload = {
        'key': api_key,
        'method': 'userrecaptcha',
        'googlekey': sitekey,
        'pageurl': url,
        'json': 1
    }
    resposta = requests.post('http://2captcha.com/in.php', data=payload)
    return resposta.json().get('request')

# === 2. Função para buscar resultado do captcha ===
def obter_resultado(api_key, captcha_id, tentativas=30, intervalo=7):
    for tentativa in range(tentativas):
        time.sleep(intervalo)
        payload = {
            'key': api_key,
            'action': 'get',
            'id': captcha_id,
            'json': 0  # ← desativa JSON para evitar erro
        }
        resposta = requests.get('http://2captcha.com/res.php', params=payload)
        print(f"[DEBUG] Tentativa {tentativa+1} - Retorno da API: {resposta.text}")

        if 'OK|' in resposta.text:
            return resposta.text.split('|')[1]

        elif 'CAPCHA_NOT_READY' in resposta.text:
            continue  # ainda está sendo resolvido

        else:
            raise ValueError(f"[2Captcha] Resposta inesperada: {resposta.text}")

    raise TimeoutError("Captcha não resolvido após várias tentativas.")

# === 3. Navegar e injetar resposta ===
def automatizar_com_token(token_resolvido, cnpj: str, razao_social: str):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto(URL_SITE)
        page.wait_for_load_state('networkidle')

        # Aguarda campos principais
        page.wait_for_selector("select[name='entity.cdComarca']")

        # a) Selecionar "Manaus" na Comarca (value="1")
        page.select_option("select[name='entity.cdComarca']", value="1")

        # b) Selecionar "Falência e Recuperação de Crédito" (value="31")
        page.select_option("select[name='entity.cdModelo']", value="31")

        # c) Selecionar tipo Jurídica (value="J")
        page.check("input[type='radio'][value='J']")

        # d) Preencher campos de Razão Social e CNPJ
        page.get_by_label("Razão Social", exact=False).fill(razao_social)
        page.get_by_label("CNPJ", exact=False).fill(cnpj)
        
        # e) Preencher campo do e-mail
        page.fill("input[name='entity.solicitante.deEmail']", EMAIL_PADRAO)

        # Injetar token do 2Captcha
        page.evaluate(f'''
            document.getElementById("g-recaptcha-response").style.display = 'block';
            document.getElementById("g-recaptcha-response").innerHTML = "{token_resolvido}";
        ''')
        
        # Clicar no check do recaptcha
        page.check("input[type='checkbox'][value='true']")

        # Enviar
        page.wait_for_load_state("networkidle")
        page.click("input[name='pbEnviar']")

        # Tirar screenshot do resultado
        time.sleep(4)
        filename = f"{normalizar_cnpj(cnpj)}.png"
        page.screenshot(path=str(OUTPUT_DIR / filename), full_page=True)

        # Reiniciar processo
        page.click("input[name='pbNovo']")  # Cadastrar outro pedido
        browser.close()

        time.sleep(5)
        browser.close()


# === Execução principal ===
if __name__ == '__main__':
    df = pd.read_excel(PLANILHA, sheet_name=ABA)
    df[COL_CNPJ] = df[COL_CNPJ].astype(str).apply(normalizar_cnpj)

    for _, row in df.iterrows():
        cnpj = row[COL_CNPJ]
        razao = row[COL_RAZAO]
        
        print(f"[1] Solicitando resolução do reCAPTCHA para CNPJ {cnpj}...")
        captcha_id = solicitar_captcha(API_KEY_2CAPTCHA, SITEKEY, URL_SITE)

        print("[2] Aguardando resposta do 2Captcha...")
        token = obter_resultado(API_KEY_2CAPTCHA, captcha_id)

        print("[3] Token recebido. Automatizando para:", razao)
        automatizar_com_token(token, cnpj, razao)
