from conexao import criar_conexao, fechar_conexao
import pandas as pd
from flask import Flask
from urllib.parse import quote_plus
from sqlalchemy import create_engine

def select_todos_usuarios(con):
    cursor = con.cursor()
    sql = "SELECT * FROM leilao"
    cursor.execute(sql)

#remover duplicatas

    resultado = []
    dicionario = {}

    for (id, nome, dataleilao) in cursor:
        chave = (nome, dataleilao)
        if chave not in dicionario:
            dicionario[chave] = True
            resultado.append((id, nome, dataleilao))

    cursor.close()
    return resultado


def main():
    con = criar_conexao("localhost", "root", "", "leilao")

    resultado = select_todos_usuarios(con)

    fechar_conexao(con)

    df = pd.DataFrame(resultado, columns=['id', 'nome', 'dataleilao'])
    df.to_excel('Lotes.xlsx', index=False)


    tabela = pd.read_excel("planilha2.xlsx")
    tabela['dataleilao'] = tabela['nome'].str.split("@", n=1, expand=True)[1]
    tabela['nome'] = tabela['nome'].str.split("@", n=1, expand=True)[0]
    tabela['dataleilao'] = pd.to_datetime(tabela['dataleilao'], format='%d%m%Y')
    df.drop_duplicates(subset=['nome', 'dataleilao'], keep=False, inplace=True)
    count = 0
    
    for nomes in df:
            count +=1
    if nomes in df:
            tabela["OBS"] = tabela["nome"].isin(df["nome"]).apply(lambda y: "sem OBS" if y else "com OBS")
            df["OBS"] = df["dataleilao"].isin(tabela["dataleilao"]).apply(lambda y: "sem OBS" if y else "com OBS")
            

    mauricio = df.merge(tabela, right_index=True, left_index=True, how='outer')
    print(mauricio)



if __name__ == "__main__":
    main()