# import pandas as pd
# import requests


# def encontrar_geolocalizacao(endereco):
#     url = f"https://nominatim.openstreetmap.org/search?format=json&limit=1&q={endereco}"
#     response = requests.get(url)
    
#     if response.status_code == 200:
#         dados = response.json()
#         if dados:
#             latitude = dados[0]['lat']
#             longitude = dados[0]['lon']
#             return latitude, longitude
#         else:
#             return None, None
#     else:
#         return None, None


# caminho_arquivo = "C:\\Users\\caua.campos\\Desktop\\novosProjetos\\ClientesOnMaps.xlsx"
# nome_planilha = 'Planilha1'  
# coluna_enderecos = 'endereços'  


# try:
#     dados_excel = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha)
#     enderecos = dados_excel[coluna_enderecos].tolist()

#     for endereco in enderecos:
#         latitude, longitude = encontrar_geolocalizacao(endereco)
#         if latitude is not None and longitude is not None:
#             print(f"Endereço: {endereco} | Latitude: {latitude} | Longitude: {longitude}")
#         else:
#             print(f"Endereço '{endereco}' não encontrado.")
# except FileNotFoundError:
#     print("Arquivo não encontrado. Verifique o caminho do arquivo.")
# except KeyError:
#     print("A planilha ou a coluna especificada não foi encontrada.")
# except Exception as e:
#     print(f"Ocorreu um erro: {e}")




import pandas as pd
import requests
import time
import sys

def encontrar_geolocalizacao(endereco):
    print(f"Iniciando geolocalização para: {endereco}")
    url = f"https://nominatim.openstreetmap.org/search?format=json&limit=1&q={endereco}"
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        dados = response.json()
        if dados:
            latitude = dados[0]['lat']
            longitude = dados[0]['lon']
            return latitude, longitude
        else:
            return None, None
    except requests.Timeout:
        print(f"Timeout para o endereço: {endereco}")
        return None, None
    except requests.RequestException as e:
        print(f"Erro na requisição para o endereço '{endereco}': {e}")
        return None, None
    finally:
        print(f"Concluída geolocalização para: {endereco}")

caminho_arquivo = r"C:\Users\cauac\OneDrive\Área de Trabalho\importação\ClientesOnMaps.xlsx"
nome_planilha = 'Planilha1'
coluna_enderecos = 'endereços'
caminho_saida_excel = r"C:\Users\cauac\OneDrive\Área de Trabalho\importação\ResultadosGeolocalizacao.xlsx"
caminho_saida_excel_sem_lat_long = r"C:\Users\cauac\OneDrive\Área de Trabalho\importação\EnderecosSemGeolocalizacao.xlsx"

try:
    dados_excel = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha)
    enderecos = dados_excel[coluna_enderecos].tolist()

    resultados = []  # Lista para armazenar os resultados
    enderecos_sem_geolocalizacao = []  # Lista para armazenar endereços sem geolocalização

    for endereco in enderecos:
        try:
            latitude, longitude = encontrar_geolocalizacao(endereco)
            if latitude is not None and longitude is not None:
                resultados.append({
                    'Endereço': endereco,
                    'Latitude': latitude,
                    'Longitude': longitude
                })
            else:
                enderecos_sem_geolocalizacao.append(endereco)
            time.sleep(1)
        except Exception as e:
            print(f"Erro ao processar o endereço '{endereco}': {e}")

    resultados_df = pd.DataFrame(resultados)
    resultados_df.to_excel(caminho_saida_excel, index=False)
    print(f"Resultados salvos em: {caminho_saida_excel}")

    enderecos_sem_geolocalizacao_df = pd.DataFrame({'Endereço': enderecos_sem_geolocalizacao})
    enderecos_sem_geolocalizacao_df.to_excel(caminho_saida_excel_sem_lat_long, index=False)
    print(f"Endereços sem geolocalização salvos em: {caminho_saida_excel_sem_lat_long}")

except FileNotFoundError:
    print("Arquivo não encontrado. Verifique o caminho do arquivo.")
except KeyError:
    print("A planilha ou a coluna especificada não foi encontrada.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")
