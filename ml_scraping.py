from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from datetime import datetime
import time

wb = Workbook()
ws = wb.active
ws.title = "Anuncios Mercado Livre"
ws.append(["Marca", "Título", "Preço", "Desconto", "Avaliação", "Link"])

def scraping():
    print('Executando o scraping...')
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True, slow_mo=100)
        context = navegador.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")
        pagina = context.new_page()
        pagina.goto("https://www.mercadolivre.com.br/ofertas?page=1")

        def loop_scraping():
            pagina.wait_for_selector('div.items-with-smart-groups', timeout=120000)
            anuncios = pagina.locator('div.items-with-smart-groups div.andes-card')
            quantidade = anuncios.count()

            for i in range(quantidade):
                try:
                    #VERIFICAR SE TEM NOME DE MARCA
                    if anuncios.nth(i).locator('span.poly-component__brand').is_visible(timeout=2000):
                        marca = anuncios.nth(i).locator('span.poly-component__brand').inner_text()
                    else:
                        marca = "Não informado"

                    link = anuncios.nth(i).locator('a.poly-component__title').get_attribute("href")
                    titulo = anuncios.nth(i).locator('a.poly-component__title').inner_text()
                    preco = anuncios.nth(i).locator('span.andes-money-amount__fraction').first.inner_text()

                    #VERIFICAR SE TEM DESCONTO
                    if anuncios.nth(i).locator('span.andes-money-amount__discount').is_visible(timeout=2000):
                        desconto = anuncios.nth(i).locator('span.andes-money-amount__discount').inner_text()
                    else:
                        desconto = "Não informado"

                    #VERIFICAR SE TEM AVALIAÇÃO
                    if anuncios.nth(i).locator('span.poly-reviews__rating').is_visible(timeout=2000):
                        avaliacao = anuncios.nth(i).locator('span.poly-reviews__rating').inner_text()
                    else:
                        avaliacao = "Não informado"

                    ws.append([marca, titulo, preco, desconto, avaliacao, link])
                    
                except Exception as e:
                    print(f"Erro ao buscar anúncio {i+1}: {e}")

        while True:
            try:
                time.sleep(3)
                loop_scraping()
                pagina.get_by_title("Siguiente").click()
            except Exception as erro:
                print(f'O scraping foi concluído.')
                break
        
    data_hora = datetime.now().strftime("%d-%m-%y")
    wb.save(f'relatorios/anuncios_mercadolivre_{data_hora}.xlsx')

if __name__ == "__main__":
    scraping()