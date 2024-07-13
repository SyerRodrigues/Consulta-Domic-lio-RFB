import pyautogui as pa
import openpyxl
import time

# Define uma pausa de 7 segundos entre cada comando do pyautogui
pa.PAUSE = 7

# Caminho para a planilha Excel contendo os CNPJs das empresas
caminho_planilha = r'Y:\\Automatização Exemplo\\BACKOFFICE\\Planilhas - backoffice\\empresas_ecac_cnpj.xlsx'
try:
    # Tenta carregar a planilha
    workbook = openpyxl.load_workbook(caminho_planilha)
    ecac_sheet = workbook['ecac']  # Seleciona a aba chamada 'ecac'
except FileNotFoundError:
    # Se a planilha não for encontrada, imprime uma mensagem de erro e encerra o programa
    print(f"Arquivo não encontrado: {caminho_planilha}")
    exit(1)

# Abre o navegador Chrome
pa.press('win')
pa.write('chrome')
pa.press('ENTER')

# Acessa a página de login do eCAC da Receita Federal
pa.write('https://cav.receita.fazenda.gov.br/autenticacao/login')
pa.press('ENTER')

# Clica no botão para entrar com GOV.BR
pa.click(x=895, y=481)  # coordenada entrar gov
# Clica para selecionar o certificado digital
pa.click(x=964, y=650)  # coordenada certificado digital
pa.press('ENTER')

# Navega para a caixa postal
pa.click(x=1250, y=229)  # caixa postal
# Abre o menu de impressão
pa.hotkey('ctrl', 'p')
pa.press('enter')
# Digita o nome do arquivo para salvar a impressão
pa.write('167 - Exemplo Contadores - 052024')
pa.click(x=125, y=46)
# Define o caminho para salvar o arquivo
pa.write('H:\\Drives compartilhados\\Bienio_Corrente\\2024\\ECAC - CLIENTES - CAIXA ELETRONICO\\MAIO\\ECAC4')
pa.press('enter')
pa.click(x=1198, y=693)

# Itera sobre as linhas da planilha a partir da linha 9
for linha in ecac_sheet.iter_rows(min_row=9):
    # Clica para abrir a seção de consulta por CNPJ
    pa.click(x=1083, y=225)
    pa.click(x=545, y=387)
    # Digita o CNPJ da empresa
    pa.write(str(linha[2].value))
    pa.click(x=781, y=404)
    # Navega para a seção de impressão novamente
    pa.click(x=1234, y=225)
    pa.hotkey('ctrl', 'p')
    pa.press('enter')
    # Digita o nome do arquivo com o nome da empresa anonimizados
    pa.write('Exemplo ' + linha[1].value)
    pa.click(x=1198, y=693)
