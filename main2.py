import asyncio
from playwright.async_api import async_playwright
from datetime import datetime, timedelta
import time
import os
import shutil
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

DOWNLOAD_DIR = "/tmp"

def rename_downloaded_file(download_dir, download_path):
    try:
        current_hour = datetime.now().strftime("%H")
        new_file_name = f"QUEUE-{current_hour}.csv"
        new_file_path = os.path.join(download_dir, new_file_name)
        if os.path.exists(new_file_path):
            os.remove(new_file_path)
        shutil.move(download_path, new_file_path)
        print(f"Arquivo salvo como: {new_file_path}")
        return new_file_path
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {e}")
        return None

def update_packing_google_sheets(csv_file_path):
    try:
        if not os.path.exists(csv_file_path):
            print(f"Arquivo {csv_file_path} não encontrado.")
            return
            
        # 1. Autenticação (feita apenas uma vez)
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("hxh.json", scope)
        client = gspread.authorize(creds)
        
        # 2. Prepara os dados do CSV (feito apenas uma vez)
        df = pd.read_csv(csv_file_path)
        df = df.fillna("")
        dados_para_enviar = [df.columns.values.tolist()] + df.values.tolist()

        # 3. Lista de planilhas e abas que vão receber os dados
        planilhas_destino = [
            {
                "url": "https://docs.google.com/spreadsheets/d/1LZ8WUrgN36Hk39f7qDrsRwvvIy1tRXLVbl3-wSQn-Pc/edit?gid=1903873220#gid=1903873220",
                "aba": "Queue list"
            },
            {
                "url": "https://docs.google.com/spreadsheets/d/1nMLHR6Xp5xzQjlhwXufecG1INSQS4KrHn41kqjV9Rmk/edit?gid=0#gid=0",
                "aba": "Base"
            }
        ]

        # 4. Loop para atualizar todas as planilhas configuradas
        for config in planilhas_destino:
            print(f"Iniciando atualização da aba '{config['aba']}'...")
            sheet = client.open_by_url(config['url'])
            worksheet = sheet.worksheet(config['aba'])
            worksheet.clear()
            worksheet.update(dados_para_enviar)
            print(f"Dados enviados com sucesso para a aba '{config['aba']}'.")
            
        time.sleep(5)
    except Exception as e:
        print(f"Erro durante o processo de envio para o Sheets: {e}")

async def main():        
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    
    browser = None 
    
    async with async_playwright() as p:
        try:
            browser = await p.chromium.launch(headless=False, args=["--no-sandbox", "--disable-dev-shm-usage", "--window-size=1920,1080"])
            context = await browser.new_context(accept_downloads=True)
            page = await context.new_page()
            
            # LOGIN
            await page.goto("https://spx.shopee.com.br/")
            await page.wait_for_selector('xpath=//*[@placeholder="Ops ID"]', timeout=15000)
            await page.locator('xpath=//*[@placeholder="Ops ID"]').fill('Ops314485')
            await page.locator('xpath=//*[@placeholder="Senha"]').fill('@Shopee123')
            await page.locator('xpath=/html/body/div[1]/div/div[2]/div/div/div[1]/div[3]/form/div/div/button').click()
            await page.wait_for_timeout(15000)
            
            try:
                await page.locator('.ssc-dialog-close').click(timeout=5000)
            except:
                print("Nenhum pop-up foi encontrado.")
                await page.keyboard.press("Escape")

            # NAVEGAÇÃO E DOWNLOAD
            await page.goto("https://spx.shopee.com.br/#/queue-list")
            await page.wait_for_timeout(10000)
            await page.locator('xpath=/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div/div[2]/div[2]/div[2]/div/div[2]/span[2]/span/button').click()
            await page.wait_for_timeout(10000)
            await page.locator('xpath=//li[1]//span[1]//div[1]//div[1]//span[1]').click()
            await page.wait_for_timeout(10000)

            d3 = (datetime.now() - timedelta(days=7)).strftime("%Y/%m/%d")
            d1 = (datetime.now() + timedelta(days=1)).strftime("%Y/%m/%d")

            # Primeiro campo de data
            date_input1 = page.get_by_role("textbox", name="Data de início").nth(1)
            await date_input1.wait_for(state="visible", timeout=10000)
            await date_input1.click(force=True)
            await date_input1.fill(d3)

            # Segundo campo de data
            date_input2 = page.get_by_role("textbox", name="Data final").nth(1)
            await date_input2.wait_for(state="visible", timeout=10000)
            await date_input2.click(force=True)
            await date_input2.fill(d1)
            await page.wait_for_timeout(5000)
            
            await page.get_by_text("Exportação da Lista de Fila").click()
            await page.wait_for_timeout(5000)
            
            await page.get_by_role('button', name='Confirm').click()
            await page.wait_for_timeout(15000)

            # 👉 Botão de download
            async with page.expect_download() as download_info:
                await page.get_by_role("button", name="Baixar").nth(0).click()
            
            download = await download_info.value
            download_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
            await download.save_as(download_path)
            new_file_path = rename_downloaded_file(DOWNLOAD_DIR, download_path)
            
            # Atualizar Google Sheets
            if new_file_path:
                update_packing_google_sheets(new_file_path)
                print("Processo finalizado com sucesso.")
                
        except Exception as e:
            print(f"Erro durante o processo de extração: {e}")
        finally:
            if browser:
                await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
