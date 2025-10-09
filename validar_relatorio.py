import os
import pandas as pd
import openpyxl
import warnings
import math
from datetime import datetime, timedelta


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def validar_iss(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, header=15, engine='openpyxl')
    df = df.drop_duplicates(subset=['CNPJ', 'Nº Documento'], keep='first')
    divergencias = []
    
    for _, row in df.iterrows():
        try:
            valor_iss = float(row["Valor ISS"])
            fluig = int(row["Fluig"]) if pd.notna(row["Fluig"]) else "(vazio)"
            fornecedor = str(row["Cliente / Fornec"])
            idmov = int(row["ID_Mov"])
            data_criacao = row["Dt  Criação"]
            percentual_iss = float(row["% ISS"])
            base_iss = float(row["Base Calc ISS"])
            calculado = base_iss * percentual_iss

            if abs(valor_iss - calculado) >= 0.02:
                divergencias.append(f" ❌ Valor ISS R$ {valor_iss:.2f} ≠ valor calculado R$ {calculado:.2f} → Fluig: {fluig} | Fornecedor: {fornecedor} | IDMOV: {idmov} | Data de lançamento: {data_criacao}.")
        except:
            continue
    return divergencias

def validar_inss(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, header=15, engine='openpyxl')
    df = df.drop_duplicates(subset=['CNPJ', 'Nº Documento'], keep='first')
    divergencias = []
    for _, row in df.iterrows():
        try:
            valor_inss = float(row["Valor INSS"])
            base_inss = float(row["Base Calc INSS"])
            if base_inss == 0:
                if valor_inss != 0:
                    divergencias.append(f" ❌ Base INSS = R$ {base_inss:.2f} porém valor INSS = R$ {valor_inss:.2f} → Fluig: {fluig} | Fornecedor: {fornecedor} | IDMOV: {idmov} | Data de lançamento: {data_criacao}.")
            if base_inss > 0 and valor_inss == 0:
                divergencias.append(f" ❌ Base INSS = R$ {base_inss:.2f} porém valor INSS = R$ {valor_inss:.2f} → Fluig: {fluig} | Fornecedor: {fornecedor} | IDMOV: {idmov} | Data de lançamento: {data_criacao}.")
            elif (row["% INSS"]) == '11%' or (row["% INSS"]) == '3.5%':
                percentual_inss = float(row["% INSS"])
            data_criacao = row["Dt  Criação"]
            inss_calculado = base_inss * percentual_inss
            fluig = int(row["Fluig"]) if pd.notna(row["Fluig"]) else "(vazio)"
            fornecedor = str(row["Cliente / Fornec"])
            idmov = int(row["ID_Mov"])
            if abs(valor_inss - inss_calculado) > 0.01:
                divergencias.append(f" ❌ Valor INSS R$ {valor_inss:.2f} ≠ valor calculado R$ {inss_calculado:.2f} → Fluig: {fluig} | Fornecedor: {fornecedor} | IDMOV: {idmov} | Data de lançamento: {data_criacao}.")
        except:
            continue
    return divergencias

def validar_csrf(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, header=15, engine='openpyxl')
    df = df.drop_duplicates(subset=['CNPJ', 'Nº Documento'], keep='first')
    divergencias = []
    for _, row in df.iterrows():
        try:
            valor_csrf = float(row["Valor CSRF"])
            percentual_csrf = float(row["% CSRF"])
            data_criacao = row["Dt  Criação"]
            base_csrf = float(row["Valor Bruto"])
            csrf_calculado = base_csrf * percentual_csrf
            fluig = int(row["Fluig"]) if pd.notna(row["Fluig"]) else "(vazio)"
            fornecedor = str(row["Cliente / Fornec"])
            idmov = int(row["ID_Mov"])
            if abs(valor_csrf - csrf_calculado) >= 0.02:
                divergencias.append(f" ❌ Valor CSRF R$ {valor_csrf:.2f} ≠ valor calculado R$ {csrf_calculado:.2f} → Fluig: {fluig} | Fornecedor: {fornecedor} | IDMOV: {idmov} | Data de lançamento: {data_criacao}.")
                
        except:
            continue   
    return divergencias


def validar_todos_os_impostos(pasta_das_planilhas):
    import os

    resultados = {}

    arquivos = sorted([
        arq for arq in os.listdir(pasta_das_planilhas)
        if arq.endswith('.xlsx') and not arq.startswith('~$')
    ])

    for nome_arquivo in arquivos:
        caminho_arquivo = os.path.join(pasta_das_planilhas, nome_arquivo)
        resultados[nome_arquivo] = {}

        
        try:
            divergencias_iss = validar_iss(caminho_arquivo)
            resultados[nome_arquivo]["ISS"] =  divergencias_iss

            divergencias_inss = validar_inss(caminho_arquivo)
            resultados[nome_arquivo]["INSS"] = divergencias_inss

            divergencias_csrf = validar_csrf(caminho_arquivo)
            resultados[nome_arquivo]["CSRF"] = divergencias_csrf



        except Exception as e:
            resultados[nome_arquivo]["AVISO"] =  f"⚠️  Filial sem movimento no periodo selecionado."
        

    return resultados