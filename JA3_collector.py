import os
import requests
import csv
import openpyxl
from io import StringIO
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


#EXCEL_DIRECTORY = "C:\\Users...

JA3_URLS = [
    'https://sslbl.abuse.ch/blacklist/ja3_fingerprints.csv',
    # Adicionar mais URLs de listas de JA3 aqui no futuro
]

def fetch_ja3_data(urls):
 
    all_ja3_fingerprints = set()
    
    for url in urls:
        try:
            print(f"Buscando dados em: {url}")
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            csv_data = StringIO(response.text)
            reader = csv.reader(line for line in csv_data if not line.startswith('#'))
            
            for row in reader:
                if len(row) > 0:
                    ja3_fingerprint = row[0].strip()
                    if ja3_fingerprint:
                        all_ja3_fingerprints.add(ja3_fingerprint)
            
        except requests.exceptions.RequestException as e:
            print(f"Erro ao acessar a URL {url}: {e}")
        except Exception as e:
            print(f"Erro ao processar o CSV de {url}: {e}")
            
    return all_ja3_fingerprints

def create_ja3_excel(ja3_data, directory, filename=""): #define o nome de saída do arquivo com os IOCs JA3
   
    if not ja3_data:
        print("Nenhum dado JA3 para criar o arquivo Excel.")
        return
        
    filepath = os.path.join(directory, filename)
    
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "JA3 Fingerprints"
        
        sheet['A1'] = "JA3"
        
        for index, fingerprint in enumerate(sorted(list(ja3_data)), start=2):
            sheet[f'A{index}'] = fingerprint
            
        workbook.save(filepath)
        print(f"Arquivo Excel '{filepath}' criado com sucesso.")
    except Exception as e:
        print(f"Erro ao criar o arquivo Excel: {e}")

if __name__ == "__main__":
    ja3_fingerprints = fetch_ja3_data(JA3_URLS)
    
    if ja3_fingerprints:
        create_ja3_excel(ja3_fingerprints, EXCEL_DIRECTORY)
