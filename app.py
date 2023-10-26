import PyPDF2
import re
from openpyxl import Workbook

# Abre o PDF para leitura
pdf_file = "Vouchers.pdf"

pdf_reader = PyPDF2.PdfReader(pdf_file)

# Cria um arquivo Excel
excel_file  = "vouchers.xlsx"
workbook = Workbook()
sheet = workbook.active
#Inicia uma lista para armazenar os padrões encontrados dentro do PDF
dados_encontrados = []

# Definir expressão para recuperar valores
padrao = r'\d{5}\s*-\s*\d{5}'

# Obter o número de páginas usando len(reader.pages)
num_pages = len(pdf_reader.pages)

# Iterar em cada página PDF
for page_num in range(num_pages):
    page = pdf_reader.pages[page_num]
    text = page.extract_text()

    # Procura o padrão na pagina atual
    matches = re.findall(padrao, text)

    # Adiciona os resultados a lista de palavras encontradas
    dados_encontrados.extend(matches)

# Adiciona os resultados encontrados na planilha
for dado in dados_encontrados:
    sheet.append([dado])

# Salva o arquivo em formato xlsx
workbook.save(excel_file)

print("Dados de acordo com o padrão encontrados! Conforme solicitado!")
