import pandas as pd
import requests


def encontrar_geolocalizacao(endereco):
    url = f"https://nominatim.openstreetmap.org/search?format=json&limit=1&q={endereco}"
    response = requests.get(url)
    
    if response.status_code == 200:
        dados = response.json()
        if dados:
            latitude = dados[0]['lat']
            longitude = dados[0]['lon']
            return latitude, longitude
        else:
            return None, None
    else:
        return None, None


caminho_arquivo = "C:\\Users\\caua.campos\\Desktop\\novosProjetos\\teste.xlsx"
nome_planilha = 'Planilha1'  
coluna_enderecos = 'endereços'  


try:
    dados_excel = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha)
    enderecos = dados_excel[coluna_enderecos].tolist()

    for endereco in enderecos:
        latitude, longitude = encontrar_geolocalizacao(endereco)
        if latitude is not None and longitude is not None:
            print(f"Endereço: {endereco} | Latitude: {latitude} | Longitude: {longitude}")
        else:
            print(f"Endereço '{endereco}' não encontrado.")
except FileNotFoundError:
    print("Arquivo não encontrado. Verifique o caminho do arquivo.")
except KeyError:
    print("A planilha ou a coluna especificada não foi encontrada.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")