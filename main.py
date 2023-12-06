import os
import time
import concurrent.futures
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import pandas as pd
from selenium.webdriver.common.by import By

class WebScraper(object):
    def __init__(self, excel_file_path):
        self.driver = webdriver.Firefox()
        self.excel_file_path = excel_file_path

    def pesquisar_latitude_longitude(self, endereco):
        driver = self.driver
        driver.get("https://maps.google.com/")
        input_element = driver.find_element(By.NAME, 'q')

        if input_element is None:
            raise Exception('Volte e tente de novo')

        input_element.send_keys(endereco)
        input_element.send_keys(Keys.ENTER)

        WebDriverWait(driver, 10).until(EC.url_contains('/@'))
        url = driver.current_url

        # Extrair latitude e longitude da URL
        latitude, longitude = re.findall(r'@(-?\d+\.\d+),(-?\d+\.\d+)', url)[0]
        return float(latitude), float(longitude)

    def buscar_coordenadas_excel(self):
        df = pd.read_excel(self.excel_file_path)

        df['Latitude'] = None
        df['Longitude'] = None

        for i, linha in df.iterrows():
            endereco = linha['Endereço Completo']

            latitude, longitude = self.pesquisar_latitude_longitude(endereco)
            df.at[i, 'Latitude'] = latitude
            df.at[i, 'Longitude'] = longitude

        return df

    def fechar_driver(self):
        self.driver.quit()

def processar_excel(excel_file):
    web_scraper = WebScraper(excel_file)
    df_resultado = web_scraper.buscar_coordenadas_excel()
    output_file_path = excel_file.replace(".xlsx", "_saida.xlsx")
    df_resultado.to_excel(output_file_path, index=False)
    web_scraper.fechar_driver()
    return output_file_path  # Retorna o caminho do arquivo de saída

# Lista de caminhos dos arquivos Excel
excel_files = [
    r"caminhodoseuarquivo\novo.xlsx",
]

# Medir o tempo de execução
start_time = time.time()

# Executar em paralelo usando threads
with concurrent.futures.ThreadPoolExecutor() as executor:
    output_paths = list(executor.map(processar_excel, excel_files))

end_time = time.time()
tempo_total = end_time - start_time
print(f"Tempo total de execução: {tempo_total} segundos")

# Juntar os resultados em um único DataFrame
dfs = [pd.read_excel(output_path) for output_path in output_paths]
df_final = pd.concat(dfs, ignore_index=True)

# Especificar o caminho completo para o arquivo Excel final
output_final_file_path = r"caminhodoseuarquivo\resultado_final.xlsx"

# Salvar o DataFrame resultante em um novo arquivo Excel
df_final.to_excel(output_final_file_path, index=False)

print(f"Arquivo final salvo em: {output_final_file_path}")
