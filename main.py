from sqlalchemy import create_engine
from urllib.parse import quote_plus
import pandas as pd
import datetime
from datetime import timedelta

dataConsulta = datetime.datetime.today() - timedelta(1)
dataConsultaFormatada = dataConsulta.strftime("%d/%m/%Y")
dataConsultaFormatadaCsv = dataConsulta.strftime("%d.%m.%Y")

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
    result = connection.execute("SELECT leiloeiro, dataleilao, placa, chassi, lote, datainc FROM cnvl1.leilaototal_new WHERE datainc = '%s'AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC" %dataConsultaFormatada)
    
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
    
    for i, row in df.iterrows():
        match = ((tabela['nome'] == row['leiloeiro']) & (tabela['dataleilao'] == row['dataleilao']))
        tabela.loc[match, 'OBS'] = 'com OBS'
        tabela.loc[~match, 'OBS'] = 'sem OBS'
        tabela.loc[~tabela['nome'].isin(df['leiloeiro']), 'OBS'] = 'com OBS'
    
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
    
    tabela = tabela.drop("id", axis=1)
    tabela.to_excel('teste2.xlsx', index=False)
    
    csv = pd.read_sql("SELECT leiloeiro, dataleilao, placa, chassi, lote, datainc FROM cnvl1.leilaototal_new WHERE datainc = '%s'AND origemTabela = 'AL' AND TRIM(observacao) = '' GROUP BY dataleilao , placa ORDER BY CONCAT_WS('-', SUBSTRING(dataleilao, 7, 4), SUBSTRING(dataleilao, 4, 2), SUBSTRING(dataleilao, 1, 2)) ASC , leiloeiro ASC" %dataConsultaFormatada,  create_database_connection())
    csv.to_csv("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Normais/Lotes sem observacao - %s.csv" %dataConsultaFormatadaCsv, index=False)
    text = open("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Normais/Lotes sem observacao - %s.csv" %dataConsultaFormatadaCsv, 'r')
    text = ''.join([i for i in text])
    text = text.replace(',', ';')
    text = text.replace('""', '')
    x = open("//192.168.0.183/ti_leilao/Recortes/2023/Lotes sem observacao/Normais/Lotes sem observacao - %s.csv" %dataConsultaFormatadaCsv,"w")
    x.writelines(text)
    x.close()

    print(tabela)

if __name__ == "__main__":
    main()