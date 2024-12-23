from tabula import read_pdf
import pdfplumber
import pandas as pd
import os
import re

# Extrair texto do PDF para informações adicionais como Inscrição Estadual e Extrato
def extract_text_from_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
            return text
    except Exception as e:
        print(f"Erro ao extrair texto de {pdf_path}: {e}")
        return ""

# Processar texto extraído para capturar Inscrição Estadual, Extrato
def extract_inscricao_estadual_and_extract(text):
    inscricao_estadual = "Não encontrado"
    extrato = "Não encontrado"
    
    # Buscar por Inscrição Estadual
    for line in text.splitlines():
        if "Inscrição Estadual" in line:
            inscricao_estadual = line.split(":")[-1].strip()

    # Tentar capturar "Extrato" usando uma expressão regular
    match = re.search(r"Extrato[:\s]*([^\n]+)", text)
    if match:
        extrato = match.group(1).strip()

    return inscricao_estadual, extrato

# Extrair tabelas do PDF
def extract_tables_from_pdf(pdf_path):
    try:
        tables = read_pdf(pdf_path, pages="all", lattice=True, multiple_tables=True)
        return tables
    except Exception as e:
        print(f"Erro ao extrair tabelas de {pdf_path}: {e}")
        return []

# Filtrar e processar apenas as colunas de interesse: ICMS Devido, Grupo de Mercadorias, etc.
def process_tables(tables, inscricao_estadual, periodo, extrato):
    all_data = []
    for table in tables:
        if isinstance(table, pd.DataFrame):
            # Verifique se as colunas de interesse estão presentes
            if "ICMS Devido" in table.columns and "Grupo de Mercadorias" in table.columns:
                # Selecionar apenas as colunas necessárias
                table = table[["Grupo de Mercadorias", "ICMS Devido"]]
                # Adicionar colunas "Inscrição Estadual", "Período" e "Extrato"
                table["Inscrição Estadual"] = inscricao_estadual
                table["Período"] = periodo
                table["Extrato"] = extrato
                all_data.append(table)
    return all_data

# Processar PDFs na pasta
def process_pdfs_in_folder(folder_path, output_file):
    all_data = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processando: {pdf_path}")

            # Extrair o período do nome do arquivo
            periodo = os.path.splitext(filename)[0]

            # Extrair texto e tabelas
            text = extract_text_from_pdf(pdf_path)
            inscricao_estadual, extrato = extract_inscricao_estadual_and_extract(text)
            tables = extract_tables_from_pdf(pdf_path)

            # Processar tabelas e adicionar informações adicionais
            processed_tables = process_tables(tables, inscricao_estadual, periodo, extrato)
            all_data.extend(processed_tables)

    # Combinar todas as tabelas em um único DataFrame
    if all_data:
        final_data = pd.concat(all_data, ignore_index=True)
        final_data.to_excel(output_file, index=False)
        print(f"Dados exportados para {output_file}")
    else:
        print("Nenhum dado foi extraído.")

# Caminho para a pasta com os PDFs e o arquivo de saída
folder_path = r"C:\Users\Rodrigo temp\Downloads\projet\EXTRATOS"
output_file = r"C:\\Users\\Rodrigo temp\\Downloads\\projet\\resumo_icms.xlsx"

# Executar o processo
process_pdfs_in_folder(folder_path, output_file)
