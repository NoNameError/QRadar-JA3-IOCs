import os
import requests
import openpyxl
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

#QRADAR_URL = ''
#QRADAR_TOKEN = ''

#REFERENCE_SET_NAME = ''

#EXCEL_DIRECTORY = "C:\\Users.."
EXCEL_FILENAME = "" #Excel filename no arquivo anterior

def get_ja3_data_from_excel(directory, filename):
   
    ja3_fingerprints = set()
    filepath = os.path.join(directory, filename)

    if not os.path.exists(filepath):
        print(f"Erro: Arquivo '{filepath}' não encontrado. Execute o primeiro script para gerá-lo.")
        return ja3_fingerprints

    try:
        workbook = openpyxl.load_workbook(filepath, read_only=True)
        sheet = workbook.active
        
        
        header = [cell.value for cell in sheet[1]]
        try:
            ja3_column_index = header.index('JA3')
        except ValueError:
            print("Aviso: Coluna 'JA3' não encontrada. Verifique o arquivo Excel.")
            return ja3_fingerprints
        
        for row in sheet.iter_rows(min_row=2):
            if row[ja3_column_index].value:
                ja3_value = str(row[ja3_column_index].value).strip()
                if ja3_value:
                    ja3_fingerprints.add(ja3_value)

    except Exception as e:
        print(f"Erro ao processar o arquivo '{filename}': {e}")
        
    return ja3_fingerprints

def create_or_update_reference_set(set_name, element_type, values):
    
    headers = {
        'SEC': QRADAR_TOKEN,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    
    if not values:
        print(f"Aviso: Nenhum valor para adicionar ao reference set '{set_name}'.")
        return

    try:
        for value in values:
            add_value_url = f'{QRADAR_URL}/api/reference_data/sets/{set_name}?value={value}'
            
            add_response = requests.post(
                add_value_url,
                headers=headers,
                verify=False
            )

            if add_response.status_code in [200, 201]:
                print(f"Valor '{value}' adicionado com sucesso ao reference set '{set_name}'.")
            elif add_response.status_code == 409:
                print(f"Valor '{value}' já existe no reference set '{set_name}'.")
            else:
                print(f"Erro ao adicionar valor '{value}': {add_response.text}")
                
        print(f"Processamento de {len(values)} valores para o reference set '{set_name}' concluído.")

    except Exception as e:
        print(f"Erro ao processar reference set '{set_name}': {e}")

if __name__ == "__main__":
    ja3_fingerprints_to_add = get_ja3_data_from_excel(EXCEL_DIRECTORY, EXCEL_FILENAME)

    if ja3_fingerprints_to_add:
        create_or_update_reference_set(
            REFERENCE_SET_NAME,
            'ALN',  
            ja3_fingerprints_to_add
        )
    else:
        print("Nenhum fingerprint JA3 foi encontrado no arquivo Excel para ser adicionado.")
