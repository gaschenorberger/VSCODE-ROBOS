from flask import Flask, render_template
import psycopg2

app = Flask(__name__)

DB_CONFIG = {
    "dbname": "dados_coletados",
    "user": "postgres",
    "password": "1234",
    "host": "localhost",
    "port": "5432",
}

def conectar_postgres():
    return psycopg2.connect(**DB_CONFIG)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/regime-tributario")
def regime_tributario():
    conexao = conectar_postgres()
    cursor = conexao.cursor()
    query = "SELECT nome_arquivo, ultima_atualizacao, ultima_verificacao FROM regime_tributario;"
    cursor.execute(query)
    dados = cursor.fetchall()
    cursor.close()
    conexao.close()
    return render_template("regime_tributario.html", dados=dados)

@app.route("/dados-abertos")
def dados_abertos():
    conexao = conectar_postgres()
    cursor = conexao.cursor()
    query = "SELECT nome_arquivo, ultima_atualizacao, ultima_verificacao FROM dados_abertos;"
    cursor.execute(query)
    dados = cursor.fetchall()
    cursor.close()
    conexao.close()
    return render_template("dados_abertos.html", dados=dados)

@app.route("/portal-devedores")
def portal_devedores():
    conexao = conectar_postgres()
    cursor = conexao.cursor()
    query = "SELECT nome_arquivo, ultima_atualizacao, ultima_verificacao FROM portal_devedores;"
    cursor.execute(query)
    dados = cursor.fetchall()
    cursor.close()
    conexao.close()
    return render_template("portal_devedores.html", dados=dados)

if __name__ == "__main__":
    app.run(debug=True, port=8080, host='0.0.0.0')
