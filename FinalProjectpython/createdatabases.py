import sqlite3

conn = sqlite3.connect(r'mydatabaseinout.db')
c= conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS account (code integer(30) NOT NULL
          ,name varchar(30) NOT NULL
          ,age integer(30) NOT NULL
          ,depart varchar(30) NOT NULL)''')
c.execute('''CREATE TABLE IF NOT EXISTS record (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    inout VARCHAR(30) NOT NULL,
    code INTEGER NOT NULL,
    name VARCHAR(30) NOT NULL,
    depart VARCHAR(30) NOT NULL,
    time TIMESTAMP NOT NULL,
    late VARCHAR(30) NOT NULL ,
    income INTEGER NOT NULL)''')
c.execute('''CREATE TABLE IF NOT EXISTS img(code integer(30) NOT NULL
          ,data BLOB)''')

conn.commit()
conn.close()