from conexao import criar_conexao, fechar_conexao
import pandas as pd
from flask import Flask
from urllib.parse import quote_plus
from sqlalchemy import create_engine

def select_todos_usuarios(con):
    cursor = con.cursor()
    sql = "SELECT * FROM leilao"
    cursor.execute(sql)

    for (id, nome) in cursor:
        print(id, nome)
    
    cursor.close()


def main():
    con = criar_conexao("localhost", "root", "", "leiloes")

    select_todos_usuarios(con)

    fechar_conexao(con)


if __name__ == "__main__":
    main()


engine = create_engine("mysql+pymysql://root:@localhost:3306/leiloes")
conn = engine.connect()

sql_query = pd.read_sql_query ("SELECT * FROM leilao", conn)

df = pd.DataFrame(sql_query)
df.to_excel('Lotes.xlsx', index = False)


tabela = pd.read_excel("planilha2.xlsx")
print(tabela)

mauricio = df.merge(tabela, right_index=True, left_index=True, how='outer')
print(mauricio)
