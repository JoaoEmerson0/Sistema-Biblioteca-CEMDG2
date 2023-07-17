import sqlite3

DB = sqlite3.connect("data_base/DATA_BASE.db")

cursor = DB.cursor()

cursor.execute('''
    CREATE TABLE IF NOT EXISTS reservas
    (ID INTEGER PRIMARY KEY,
    nomeAluno TEXT NOT NULL,
    nomeLivro TEXT NOT NULL,
    estatus TEXT NOT NULL,
    dataRetirada TEXT NOT NULL,
    dataEntrega TEXT NOT NULL)
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS livros
    (ID INTEGER PRIMARY KEY,
    nomeLivro TEXT NOT NULL,
    sinopce TEXT NOT NULL,
    quantidade TEXT NOT NULL,
    estatus TEXT NOT NULL)
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS usuarios
    (ID INTEGER PRIMARY KEY,
    nome TEXT NOT NULL,
    funcao TEXT NOT NULL,
    senha TEXT NOT NULL)
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS Tickets
    (numeroTicket TEXT NOT NULL,
    solicitante TEXT NOT NULL,
    dataCriacao TEXT NOT NULL,
    NumeroErro TEXT NOT NULL,
    explicacao TEXT NOT NULL,
    status TEXT NOT NULL)
''')