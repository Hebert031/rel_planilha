import openpyxl
import mysql.connector
from openpyxl import Workbook
# Script para gerar planilha com conex mysql
# autor: Hebert Riberio

#Função para executar a consulta SQL e retornar o resultado
def executar_consulta_sql(usuario):
    # Configurações de conexão com o banco de dados MySQL
    config = {
        'user': 'root',
        'password': 'm@M@vV',
        'host': 'vps-suporte-migra.m9.network',
        'database': 'supramail',
        'port': '3307',
    }

    #Conecta ao banco de dados
    conn = mysql.connector.connect(**config)
    cursor = conn.cursor()

    #Executa a consulta SQL para obter as informações relevantes
    consulta = "SELECT date, ok_as_sender FROM count_emails_v3 WHERE user = %s"
    cursor.execute(consulta, (usuario,))

    # Retorna o resultado da consulta
    return cursor.fetchall()

#Lê o arquivo contas.txt para obter as contas de e-mail
with open('contas.txt', 'r') as arquivo_contas:
    contas = arquivo_contas.read().splitlines()

#Dicionário para armazenar as informações dos relatórios
relatorios = {}

#Loop através de cada conta de e-mail
for conta in contas:
    # Executa a consulta SQL para obter as informações da conta
    resultados = executar_consulta_sql(conta)

    #Dicionário para armazenar as informações da conta
    relatorio_conta = {
        'Total de envios': 0,
        'Dia com maior quantidade de envios': None,
        'Quantidade de envios no dia': 0
    }

# Loop através dos resultados da consulta
for resultado in resultados:
    data = resultado[0]
    envios = resultado[1]

    # Verifica se a data está dentro do ano e mês desejados (maio de 2023)
    if data.year == 2023 and data.month == 5:
        # Atualiza o total de envios para a conta
        relatorio_conta['Total de envios'] += envios

        # Verifica se a quantidade de envios no dia é maior do que a registrada anteriormente
        if envios > relatorio_conta['Quantidade de envios no dia']:
            relatorio_conta['Dia com maior quantidade de envios'] = data
            relatorio_conta['Quantidade de envios no dia'] = envios

    #Armazena o relatório da conta no dicionário principal
    relatorios[conta] = relatorio_conta

#Cria uma planilha com os resultados
planilha = Workbook()
planilha_ativa = planilha.active

#Escrever cabeçalhos
planilha_ativa["A1"] = "Conta de E-mail"
planilha_ativa["B1"] = "Total de Envios"
planilha_ativa["C1"] = "Dia com Maior Quantidade de Envios"
planilha_ativa["D1"] = "Quantidade de Envios no Dia"

#Preencher os dados na planilha
linha = 2
for conta, relatorio_conta in relatorios.items():
    planilha_ativa[f"A{linha}"] = conta
    planilha_ativa[f"B{linha}"] = relatorio_conta["Total de envios"]
    planilha_ativa[f"C{linha}"] = relatorio_conta["Dia com maior quantidade de envios"]
    planilha_ativa[f"D{linha}"] = relatorio_conta["Quantidade de envios no dia"]
    linha += 1

#Salvar a planilha no disco
planilha.save("relatorio.xlsx")
