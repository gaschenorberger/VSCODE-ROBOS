from flask import Flask, request, jsonify, send_file
import requests
import json
import os

app = Flask(__name__)

# CONFIGURAÇÕES (coloque seu token e certificado aqui)
NS_TOKEN = "SEU_TOKEN_NS_TECNOLOGIA"
CERT_PATH = "certificado.pfx"
CERT_PASSWORD = "sua_senha"

@app.route('/api/download-xml', methods=['POST'])
def download_xml():
    data = request.get_json()
    chave = data.get('chave')

    if not chave:
        return jsonify({'erro': 'Chave de acesso não informada'}), 400

    url = "https://nfe.ns.eti.br/nfe/download"

    payload = {
        "chave": chave,
        "tpDown": "X"  # Apenas XML
    }

    headers = {
        "Content-Type": "application/json",
        "X-AUTH-TOKEN": NS_TOKEN
    }

    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 200:
        conteudo = response.json()
        if conteudo.get("status") == 200:
            xml_base64 = conteudo["xml"]
            with open("nfe.xml", "wb") as f:
                f.write(bytes.fromhex(xml_base64.encode().hex()))
            return send_file("nfe.xml", as_attachment=True)
        else:
            return jsonify({'erro': conteudo.get("motivo")}), 400
    else:
        return jsonify({'erro': 'Erro ao consultar a API da NS Tecnologia'}), 500

if __name__ == '__main__':
    app.run(debug=True)
