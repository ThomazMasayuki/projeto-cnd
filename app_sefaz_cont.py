# Será realizado via acesso remoto

logger.info(f"Consultando CNPJ: {cnpj_limpo}")
                page.goto(URL_PMM, timeout=TIMEOUT)
                page.locator("#VTIPOFILTRO3").check()
                page.wait_for_load_state("networkidle", timeout=15000)
                
                # Preencher o número do CNPJ
                page.locator("#VNRFILTRO").fill(cnpj_limpo)