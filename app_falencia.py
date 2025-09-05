from playwright.sync_api import sync_playwright
import time
import requests

# === Configurações ===
API_KEY_2CAPTCHA = '7d24493e32e9020be5f98d835724f30d'  # Substitua pela sua chave real
URL_SITE = 'https://consultasaj.tjam.jus.br/sco/abrirCadastro.do'
SITEKEY = '6LcnC3cdAAAAABWUEy-SzR8kMrk3FA9llI6hU934'  # detectado via inspeção no HTML

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
def obter_resultado(api_key, captcha_id, tentativas=20, intervalo=5):
    for _ in range(tentativas):
        time.sleep(intervalo)
        payload = {
            'key': api_key,
            'action': 'get',
            'id': captcha_id,
            'json': 1
        }
        resposta = requests.get('http://2captcha.com/res.php', params=payload).json()
        if resposta.get('status') == 1:
            return resposta.get('request')
    raise TimeoutError("Captcha não resolvido após várias tentativas.")

# === 3. Navegar e injetar resposta ===
def automatizar_com_token(token_resolvido):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto(URL_SITE)
        page.wait_for_load_state('networkidle')

        # Injetar o token resolvido no campo esperado pelo reCAPTCHA
        page.evaluate(f'''
            document.getElementById("g-recaptcha-response").style.display = 'block';
            document.getElementById("g-recaptcha-response").innerHTML = "{token_resolvido}";
        ''')

        # Simula o envio do formulário
        page.click('input[name="Enviar"]')

        time.sleep(5)
        browser.close()

# === Execução principal ===
if __name__ == '__main__':
    print("[1] Solicitando resolução do reCAPTCHA...")
    captcha_id = solicitar_captcha(API_KEY_2CAPTCHA, SITEKEY, URL_SITE)

    print("[2] Aguardando resposta do 2Captcha...")
    token = obter_resultado(API_KEY_2CAPTCHA, captcha_id)

    print("[3] Token recebido. Automatizando navegação...")
    automatizar_com_token(token)
