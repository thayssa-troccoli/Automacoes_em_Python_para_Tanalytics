# app.py
from flask import Flask, send_file, Response
import pandas as pd
import requests
from io import BytesIO

app = Flask(__name__)

# ----- LINKS DOS ARQUIVOS -----
LINKS = {
    "Link 1": "https://hera.rotatorserver.com/web/vertea/b54ff34ab17789d7578407bfec56a1dd_1/Primeira_Avaliacao.xlsx",
    "Link 2": "https://hera.rotatorserver.com/web/vertea/0aed16b8255fa03ad5b26804b836ffd9_1/Segundo_dia.xlsx",
    "Link 3": "https://hera.rotatorserver.com/web/vertea/9f9f908d6a4e9f058e96747cd2cc2213_1/Terceiro_dia.xlsx",
    "Link 4": "https://hera.rotatorserver.com/web/vertea/de9b4fb3181b698aa5667b67b7f0aa55_1/Quarto_dia.xlsx",
    "Link 5": "https://hera.rotatorserver.com/web/vertea/9ed2e507d5c78075086d3f3a2ba57f3c_1/Quinto_dia.xlsx",
    "Link 6": "https://hera.rotatorserver.com/web/vertea/a1698f4400fb99a81705fadeeb00cf02_1/Sexto_dia.xlsx",
}

@app.route('/')
def index():
    return """
    <h2>Baixar Arquivo Unificado</h2>
    <p>Clique no botão abaixo para gerar e baixar o Excel com todas as 6 abas dos links!</p>
    <form action="/baixar">
        <button type="submit">Baixar arquivo Excel</button>
    </form>
    """

@app.route('/baixar')
def baixar():
    abas = {}
    # Baixa cada arquivo e armazena como DataFrame
    for nome_aba, url in LINKS.items():
        resp = requests.get(url)
        if resp.status_code == 200:
            df = pd.read_excel(BytesIO(resp.content))
            abas[nome_aba] = df
        else:
            return Response(f"Erro ao baixar {nome_aba}: status {resp.status_code}", status=500)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nome_aba, df in abas.items():
            df.to_excel(writer, sheet_name=nome_aba[:31], index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="todos_links.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(debug=True)
