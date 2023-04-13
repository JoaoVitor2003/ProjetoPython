from sqlalchemy import create_engine
from urllib.parse import quote_plus
import pandas as pd
import datetime
from datetime import timedelta, date
import win32com.client as win32
import os
import shutil

def is_monday():
    return datetime.datetime.today().weekday() == 0

# Função para obter a data de consulta
def get_consultation_date():
    if is_monday():
        # Se for segunda-feira, consulta será de 3 dias atrás (sexta-feira)
        dataConsulta = datetime.datetime.today() - datetime.timedelta(days=3)
    else:
        # Caso contrário, consulta será do dia anterior
        dataConsulta = datetime.datetime.today() - datetime.timedelta(days=1)
    
    dataConsultaFormatada = dataConsulta.strftime("%d/%m/%Y")
    return dataConsultaFormatada
    
def verificaDataCsv():
    if is_monday():
        # Se for segunda-feira, consulta será de 3 dias atrás (sexta-feira)
        dataConsulta = datetime.datetime.today() - datetime.timedelta(days=3)
    else:
        # Caso contrário, consulta será do dia anterior
        dataConsulta = datetime.datetime.today() - datetime.timedelta(days=1)
    dataConsultaFormatadaCsv = dataConsulta.strftime("%d.%m.%Y")
    return dataConsultaFormatadaCsv

def create_database_connection():
    username = "remoto"
    password = "m3p1g5a8s9"
    host = "192.168.0.183"
    database = "cnvl1"

    url = f"mysql+pymysql://{username}:{quote_plus(password)}@{host}/{database}"
    engine = create_engine(url)
    connection = engine.connect()

    return connection

def close_database_connection(connection):
    connection.close()

def select_todos_usuarios():
    connection = create_database_connection()
    result = connection.execute("SELECT leiloeiro, dataleilao, placa, chassi, lote, datainc FROM cnvl1.leilaototal_new WHERE datainc = '%s'AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC" %get_consultation_date())
    #result = connection.execute("SELECT leiloeiro, dataleilao, placa, chassi, lote, datainc FROM cnvl1.leilaototal_new WHERE datainc = '21/03/2023'AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC")
    resultado = []
    dicionario = {}

    for row in result:
        chave = (row["leiloeiro"], row["dataleilao"])
        if chave not in dicionario:
            dicionario[chave] = True
            resultado.append((row["leiloeiro"], row["dataleilao"], row["lote"]))

    close_database_connection(connection)

    return resultado

def main():
    resultado = select_todos_usuarios()
    df = pd.DataFrame(resultado, columns=['leiloeiro', 'dataleilao', 'lote'])
    df.to_excel('Lotes.xlsx', index=False)

    tabela = pd.read_excel("planilha2.xlsx")
    tabela['dataleilao'] = tabela['nome'].str.split("@", n=1, expand=True)[1]
    tabela['nome'] = tabela['nome'].str.split("@", n=1, expand=True)[0]
    tabela['dataleilao'] = pd.to_datetime(tabela['dataleilao'], format='%d%m%Y')
    df.drop_duplicates(subset=['leiloeiro', 'dataleilao'], keep=False, inplace=True)
    df['dataleilao'] = df['leiloeiro'].str.split('-').str[0]
    
    #comparação básica para já colocar com OBS e sem OBS nos principais em que verifica se está no banco ou não.
    for i, row in df.iterrows():
        match = ((tabela['nome'] == row['leiloeiro']) & (tabela['dataleilao'] == row['dataleilao']))
        tabela.loc[match, 'OBS'] = 'com OBS'
        tabela.loc[~match, 'OBS'] = 'sem OBS'
        tabela.loc[~tabela['nome'].isin(df['leiloeiro']), 'OBS'] = 'com OBS'
    
    #Parte verificação de IDs iguais para seguir um padrão no mesmo ID
    tabela['id'] = tabela['nome'].str.extract('-(\d+)', expand=False).fillna(-1).astype(int)
    
    for i in range(tabela.shape[0]):
    # Salvando o id da linha atual
        id_atual = tabela.iloc[i]['id']
        data_atual = tabela.iloc[i]['dataleilao']

    # Verificando se o id da linha atual já existe na tabela
        duplicados = tabela[(tabela['id'] == id_atual) & (tabela['dataleilao'] == data_atual)].index
        if len(duplicados) > 1:
    # Se existir, colocar o valor da primeira ocorrência de "OBS" nas demais ocorrências
            primeiro_indice = duplicados[0]
            tabela.loc[duplicados[1:], "OBS"] = tabela.loc[primeiro_indice, "OBS"]
            
        #Venda direta
    conn = create_database_connection()   
    resultDireta = conn.execute("SELECT leiloeiro, dataleilao FROM sistema_melhorlance.leilaototal_new WHERE datainc = '%s' AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC" %get_consultation_date())
    #resultDireta = conn.execute("SELECT leiloeiro, dataleilao FROM sistema_melhorlance.leilaototal_new WHERE datainc = '21/03/2023' AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC")
    df_direta = pd.DataFrame(resultDireta, columns=['leiloeiro', 'dataleilao'])
    # Verificação se existem resultados na query
    if len(df_direta) >= 1:
        df_direta = df_direta.drop_duplicates()
    for i, row in df_direta.iterrows():
        leiloeiro = df_direta.loc[i, 'leiloeiro']
        tabela.loc[(tabela['nome'] == leiloeiro) & (tabela['dataleilao'] == row['dataleilao']), 'OBS'] = 'sem OBS'
    else:
    # Continuação do código sem realizar a busca na tabela
        pass      
    
    tabela = tabela.drop("id", axis=1)
    tabela.to_excel('teste2.xlsx', index=False)
    print(tabela)
    
    if len(df) >= 2:
        csv = pd.read_sql("SELECT leiloeiro, dataleilao, placa, chassi, lote, datainc FROM cnvl1.leilaototal_new WHERE datainc = '%s'AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC" %get_consultation_date(), create_database_connection())
        #csv = pd.read_sql("SELECT leiloeiro, dataleilao, placa, chassi, lote, datainc FROM cnvl1.leilaototal_new WHERE datainc = '14/03/2023'AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC", create_database_connection())
        csv.to_csv("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Normais/Lotes sem observacao - {}.csv".format(verificaDataCsv()), index=False)
        text = open("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Normais/Lotes sem observacao - {}.csv".format(verificaDataCsv()), 'r')
        text = ''.join([i for i in text])
        text = text.replace(',', ';')
        text = text.replace('""', '')
        x = open("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Normais/Lotes sem observacao - {}.csv".format(verificaDataCsv()),"w")
        x.writelines(text)
        x.close()
        
    if len(df_direta) >= 2:
        csvDireta = pd.read_sql("SELECT leiloeiro, dataleilao, placa, chassi, lote, datainc FROM sistema_melhorlance.leilaototal_new WHERE datainc = '%s' AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC" %get_consultation_date(), create_database_connection())
        csvDireta.to_csv("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Venda Direta/Lotes sem observacao - {} - Venda Direta.csv".format(verificaDataCsv()), index=False)
        textDireta = open("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Venda Direta/Lotes sem observacao - {} - Venda Direta.csv".format(verificaDataCsv()), 'r')
        textDireta = ''.join([i for i in textDireta])
        textDireta = textDireta.replace(',', ';')
        textDireta = textDireta.replace('""', '')
        y = open("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Venda Direta/Lotes sem observacao - {} - Venda Direta.csv".format(verificaDataCsv()), 'w')
        y.writelines(textDireta)
        y.close()
        
    else:
        pass 
    
    def is_monday():
        return datetime.datetime.today().weekday() == 0

    if is_monday():
        dataConsulta = datetime.datetime.today() - datetime.timedelta(days=3)
        dataArquivo = dataConsulta.strftime("%d.%m.%Y")
    else: 
        dataConsulta = datetime.datetime.today() - datetime.timedelta(days=1)
        dataArquivo = dataConsulta.strftime("%d.%m.%Y")


    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.To = "rosana@infocar.com.br;suporte@infocar.com.br"
    email.CC = "daniel.gomes@infocar.com.br;daniel.vieira@infocar.com.br;mauricio.goulart@infocar.com.br;lucas.lima@infocar.com.br"

    email.Subject = "Relação de CSV sem observação  - " + dataArquivo
    email.HTMLBody = f"""<p>Bom dia!</p>
    <p>Segue em anexo o arquivo csv com os lotes que foram inseridos sem informação no campo de observação.</p>
    """

    dirAnexoNormal = "//192.168.0.183//ti_leilao//Recortes//2023//Lotes sem observacao//Normais//Lotes sem observacao - " + dataArquivo + ".csv"
    dirAnexoVendaDireta = "//192.168.0.183//ti_leilao//Recortes//2023//Lotes sem observacao//Venda Direta//Lotes sem observacao - " + dataArquivo + " - Venda Direta.csv"

    if os.path.exists(dirAnexoNormal):
        email.Attachments.Add(dirAnexoNormal)
        
    if os.path.exists(dirAnexoVendaDireta):
        email.Attachments.Add(dirAnexoVendaDireta)
            
    email.Send()

    novoDirAnexoNormal = "//192.168.0.183//ti_leilao//Recortes//2023//Lotes sem observacao//Feitos-Normais//Lotes sem observacao - " + dataArquivo + ".csv"
    novoDirAnexoVendaDireta = "//192.168.0.183//ti_leilao//Recortes//2023//Lotes sem observacao//Feitos-Venda Direta//Lotes sem observacao - " + dataArquivo + " - Venda Direta.csv"

    if os.path.exists(dirAnexoNormal):
        shutil.move(dirAnexoNormal, novoDirAnexoNormal)
        
    if os.path.exists(dirAnexoVendaDireta):
        shutil.move(dirAnexoVendaDireta, novoDirAnexoVendaDireta)

    print("Email enviado")

if __name__ == "__main__":
    main()