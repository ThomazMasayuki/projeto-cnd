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
ABA = "PMM"
COL_CNPJ = "CNPJ"
COL_VALIDADE = "VALIDADE CERTIDÃO"
COL_RAZAO = "RAZÃO SOCIAL"
OUTPUT_DIR = Path("certidoes_baixadas") / ABA.replace(" ", "_")
URL_PMM = "https://semefatende.manaus.am.gov.br/servicoJanela.php?servico=1412"
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
    
# === Helpers robustos para GeneXus / radios / iframes ===
from playwright.sync_api import TimeoutError as PWTimeout
import re

def _log_frames(page):
    # opcional: ajuda a diagnosticar se há iframes
    try:
        frames = page.frames
        logger.debug(f"[diag] frames={len(frames)}")
        for i, fr in enumerate(frames):
            logger.debug(f"  #{i}: url={getattr(fr,'url',None)!r}, name={getattr(fr,'name',None)!r}")
    except Exception:
        pass

def _first_frame_with(page, css):
    # procura o seletor em todos os frames (raiz + iframes)
    try:
        if page.locator(css).count():
            return page
    except Exception:
        pass
    for fr in page.frames:
        try:
            if fr.locator(css).count():
                return fr
        except Exception:
            continue
    return None

def selecionar_radio_cnpj(page):
    """
    Seleciona o radio 'CNPJ' usando várias estratégias:
    id do input, label[for=...], role-accessible name e, por fim, click via JS.
    Retorna o frame (page ou iframe) onde o radio foi encontrado.
    """
    page.wait_for_load_state("domcontentloaded", timeout=15000)
    _log_frames(page)

    # id real visto no seu HTML: #VTIPOFILTRO3
    fr = _first_frame_with(page, "#VTIPOFILTRO3")
    if fr:
        radio = fr.locator("#VTIPOFILTRO3")
        radio.wait_for(state="attached", timeout=8000)

        # 1) check nativo
        try:
            radio.scroll_into_view_if_needed()
            radio.check()
            if radio.is_checked():
                return fr
        except Exception:
            pass

        # 2) clicar no label vinculado
        try:
            fr.locator("label[for='VTIPOFILTRO3']").scroll_into_view_if_needed()
            fr.locator("label[for='VTIPOFILTRO3']").click()
            if radio.is_checked():
                return fr
        except Exception:
            pass

        # 3) role + accessible name
        try:
            fr.get_by_role("radio", name=re.compile(r"\bCNPJ\b", re.I)).check()
            if radio.is_checked():
                return fr
        except Exception:
            pass

        # 4) JS direto (contorna overlay/handlers)
        try:
            fr.eval_on_selector("#VTIPOFILTRO3", "el => el.click()")
            if radio.is_checked():
                return fr
        except Exception:
            pass

    # fallback por texto do label (caso id mude)
    fr = _first_frame_with(page, "label:has-text('CNPJ')")
    if fr:
        fr.locator("label:has-text('CNPJ')").click()
        try:
            r2 = fr.get_by_role("radio", name=re.compile(r"\bCNPJ\b", re.I))
            if r2 and r2.is_checked():
                return fr
        except Exception:
            pass

    raise RuntimeError("Não foi possível selecionar o radio 'CNPJ'. Verifique se há iframe/overlay.")

def preencher_cnpj_no_campo(fr, cnpj_limpo):
    """
    Preenche o campo de CNPJ (id real visto: #VNRFILTRO) simulando digitação,
    e dá TAB para disparar onblur/validações do GeneXus.
    """
    campo = fr.locator("#vNRFILTRO")
    campo.wait_for(state="visible", timeout=10000)
    campo.scroll_into_view_if_needed()
    campo.click()
    campo.fill("")          # limpa
    campo.type(cnpj_limpo)  # digitação real lida melhor com máscaras
    campo.press("Tab")      # dispara onblur/validação
    time.sleep(0.3)

# seletores padrão (ajuste se necessário)
SEL_CAPTCHA_IMG = "img[src*='Captcha']"
SEL_CAPTCHA_INPUTS = [
    "#_cfield",                # id direto (prioritário)
    "input[name='_cfield']",   # redundância útil
]

def _prep_img_for_ocr(np_img: np.ndarray) -> Image.Image:
    # cinza
    if len(np_img.shape) == 3:
        gray = cv2.cvtColor(np_img, cv2.COLOR_BGR2GRAY)
    else:
        gray = np_img.copy()
    # upscale
    gray = cv2.resize(gray, None, fx=2.8, fy=2.8, interpolation=cv2.INTER_CUBIC)
    # limpeza leve + binarização
    gray = cv2.medianBlur(gray, 3)
    thr = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                cv2.THRESH_BINARY, 31, 10)
    # garantir texto escuro em fundo claro
    if (thr == 0).sum() < (thr == 255).sum():
        thr = 255 - thr
    return Image.fromarray(thr)

def _ocr_try(pil_img: Image.Image, psm: int) -> str:
    # whitelist alfanumérica; ajuste se o captcha tiver minúsculas
    cfg = f"--oem 3 --psm {psm} -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    txt = pytesseract.image_to_string(pil_img, lang="eng", config=cfg)
    txt = re.sub(r"[^A-Za-z0-9]", "", txt).upper()
    return txt

def ler_captcha(fr, out_path: Path = None) -> str:
    """
    Captura o captcha (no frame fr), pré-processa e lê com Tesseract.
    Salva screenshot se out_path for fornecido.
    """
    el = fr.locator(SEL_CAPTCHA_IMG).first
    el.wait_for(state="visible", timeout=8000)
    if out_path:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        el.screenshot(path=str(out_path))

    # também pega a imagem como bytes para trabalhar em memória
    png_bytes = el.screenshot()
    np_img = cv2.imdecode(np.frombuffer(png_bytes, np.uint8), cv2.IMREAD_COLOR)
    pil = _prep_img_for_ocr(np_img)

    # tenta alguns PSMs comuns para captcha
    for psm in (8, 7, 13):
        txt = _ocr_try(pil, psm=psm)
        if txt and 4 <= len(txt) <= 8:  # ajuste do tamanho esperado se quiser
            return txt
    return ""  # deixou vazio se nada legível

def preencher_captcha(fr, texto_captcha: str) -> None:
        # tenta pelos seletores conhecidos (id e name)
    for sel in SEL_CAPTCHA_INPUTS:
        try:
            loc = fr.locator(sel)
            if loc.count() and loc.is_visible():
                loc.scroll_into_view_if_needed()
                loc.click()
                time.sleep(0.2)
                loc.press("Control+A")
                loc.press("Backspace")
                time.sleep(0.1)
                loc.type(texto_captcha, delay=50)  # digita como humano
                time.sleep(0.2)
                loc.press("Tab")  # dispara onblur
                logger.info(f"[Captcha] Preenchido com '{texto_captcha}' via seletor '{sel}'")
                return
        except Exception as e:
            logger.warning(f"[Captcha] Falha ao tentar preencher com seletor {sel}: {e}")

    # fallback: tenta encontrar por label (quase nunca funciona com GeneXus, mas incluímos)
    try:
        loc = fr.get_by_label("Insira o código", exact=False)
        loc.fill(texto_captcha)
        loc.press("Tab")
        return
    except Exception:
        try:
            loc = fr.locator("label:has-text('Insira o código') ~ input").first
            loc.fill(texto_captcha)
            loc.press("Tab")
        except Exception as e:
            raise RuntimeError(f"[Captcha] Não foi possível preencher com '{texto_captcha}': {e}")

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
                page.goto(URL_PMM, timeout=TIMEOUT)
                page.wait_for_load_state("domcontentloaded", timeout=15000)

                # Seleciona o radio CNPJ (retorna o frame correto)
                fr = selecionar_radio_cnpj(page)

                # Preenche o campo do CNPJ no mesmo frame
                preencher_cnpj_no_campo(fr, cnpj_limpo)

                # --- Captura + OCR do captcha (no frame correto) ---
                # Verifica se frame ainda está vivo
                try:
                    fr.title()
                except Exception:
                    logger.warning(f"Frame anterior foi fechado. Recarregando página e tentando de novo.")
                    page.goto(URL_PMM, timeout=TIMEOUT)
                    page.wait_for_load_state("domcontentloaded", timeout=10000)
                    fr = selecionar_radio_cnpj(page)
                    preencher_cnpj_no_campo(fr, cnpj_limpo)

                # Captura + OCR do captcha
                captcha_path = OUTPUT_DIR / f"captcha_{cnpj_limpo}.png"
                texto_captcha = ler_captcha(fr, out_path=captcha_path)

                logger.info(f"[Captcha] lido='{texto_captcha}'")

                # (opcional) se vier vazio, tentar 1 refresh simples da imagem e reler
                if not texto_captcha:
                    try:
                        # tente clicar em algo como "Recarregar" se existir
                        fr.locator("text=Recarregar, Atualizar, Novo Código").first.click()
                        time.sleep(0.8)
                        texto_captcha = ler_captcha(fr, out_path=captcha_path)
                        logger.info(f"[Captcha retry] lido='{texto_captcha}'")
                    except Exception:
                        pass

                # --- Preencher ---
                preencher_captcha(fr, texto_captcha or "")

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