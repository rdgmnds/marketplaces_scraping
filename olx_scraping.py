from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from datetime import datetime
import time

wb = Workbook()
ws = wb.active
ws.title = "Anuncios OLX"
ws.append(["Título", "Preço", "Local", "Link"])

def scraping():
    print('Executando o scraping...')
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True, slow_mo=100)
        context = navegador.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")
        pagina = context.new_page()

        def loop_scraping():
            pagina.wait_for_selector('div.AdListing_adListContainer__ALQla', timeout=120000)
            anuncios = pagina.locator('div.AdListing_adListContainer__ALQla section.olx-adcard')
            quantidade = anuncios.count()

            for i in range(quantidade):
                try:
                    if anuncios.nth(i).locator('h3').is_visible(timeout=2000):
                        link = anuncios.nth(i).locator('a.olx-adcard__link').get_attribute("href")
                        titulo = anuncios.nth(i).locator('h2').inner_text()
                        preco = anuncios.nth(i).locator('h3').inner_text()
                        local = anuncios.nth(i).locator('p.olx-adcard__location').inner_text()
                        ws.append([titulo, preco, local, link])
                        
                except Exception as e:
                    print(f"Erro ao buscar anúncio {i+1}: {e}")

        num_pagina = 1

        while True:
            pagina.goto(f"https://www.olx.com.br/celulares/estado-sp?o={num_pagina}")
            time.sleep(3)
            try:
                loop_scraping()
                num_pagina += 1
            except Exception as erro:
                print(f'O scraping foi concluído na página {num_pagina}.')
                break
        
    data_hora = datetime.now().strftime("%d-%m-%y")
    wb.save(f'anuncios_olx_{data_hora}.xlsx')

if __name__ == "__main__":
    scraping()