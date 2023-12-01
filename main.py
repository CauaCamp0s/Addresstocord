import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import pandas as pd
from selenium.webdriver.common.by import By

class WebScraper(object):
    def __init__(self):
        self.driver = webdriver.Firefox()

    def pesquisar_latitude_longitude(self, endereco):
        self.driver.get("https://maps.google.com/")
        input_element = self.driver.find_element(By.NAME, 'q')

        if input_element is None:
            raise Exception('Volte e tente de novo')

        input_element.send_keys(endereco)
        input_element.send_keys(Keys.ENTER)

        WebDriverWait(self.driver, 10).until(EC.url_contains('/@'))
        url = self.driver.current_url

        # Extract latitude and longitude from the URL
        latitude, longitude = re.findall(r'@(-?\d+\.\d+),(-?\d+\.\d+)', url)[0]
        return float(latitude), float(longitude)

    def buscar_coordenadas_excel(self, caminho_arquivo):
        df = pd.read_excel(caminho_arquivo)

        df['Latitude'] = None
        df['Longitude'] = None

        for indice, linha in df.iterrows():
            endereco = linha['Endereco']  # Substitua 'Endereco' pelo nome da coluna que contém os endereços
            latitude, longitude = self.pesquisar_latitude_longitude(endereco)
            df.at[indice, 'Latitude'] = latitude
            df.at[indice, 'Longitude'] = longitude

        return df

    def fechar_driver(self):
        self.driver.quit()

# Medir o tempo de execução
start_time = time.time()

# Exemplo de uso
web_scraper = WebScraper()
caminho_arquivo_excel = "C:\\Users\\caua.campos\\Desktop\\web-scrapper\\seuarquivo_modificado.xlsx"
df_resultado = web_scraper.buscar_coordenadas_excel(caminho_arquivo_excel)

# Salvar o DataFrame resultante em um novo arquivo Excel
df_resultado.to_excel("C:\\Users\\caua.campos\\Desktop\\web-scrapper\\coordenadas_resultado.xlsx", index=False)

web_scraper.fechar_driver()

end_time = time.time()
tempo_total = end_time - start_time
print(f"Tempo total de execução: {tempo_total} segundos")
