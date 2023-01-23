from conexao import criar_conexao, fechar_conexao
import pandas as pd
from flask import Flask
from urllib.parse import quote_plus
from sqlalchemy import create_engine

def select_todos_usuarios(con):
    cursor = con.cursor()
    sql = "SELECT * FROM leiloes"
    cursor.execute(sql)

    for (id, nome) in cursor:
        print(id, nome)
    
    cursor.close()


def main():
    con = criar_conexao("localhost", "root", "", "leilao")

    select_todos_usuarios(con)

    fechar_conexao(con)


if __name__ == "__main__":
    main()

engine = create_engine("mysql+pymysql://root:@localhost:3306/leilao")
conn = engine.connect()

sql_query = pd.read_sql_query ("SELECT * FROM leiloes", conn)

df = pd.DataFrame(sql_query) 
df.to_csv (r'D:\Lotes.csv', index = False)