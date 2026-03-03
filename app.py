from flask import Flask, render_template, request, send_file
import pdfplumber
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
import uuid
import os

app = Flask(__name__)

def extrair_atividades(pdf_path, companhia):
    atividades = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if texto:
                linhas = texto.split("\n")
                for linha in linhas:
                    if companhia in linha or "Todos" in linha:
                        atividades.append(linha.strip())

    return atividades

def gerar_pdf(atividades):
    nome = f"qts_{uuid.uuid4()}.pdf"
    doc = SimpleDocTemplate(nome, pagesize=A4)

    if not atividades:
        atividades = ["Nenhuma atividade encontrada para essa companhia."]

    dados = [[a] for a in atividades]

    tabela = Table(dados)
    tabela.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
    ]))

    elementos = [tabela]
    doc.build(elementos)

    return nome

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        arquivo = request.files["dsi"]
        companhia = request.form["companhia"]

        nome_temp = f"temp_{uuid.uuid4()}.pdf"
        arquivo.save(nome_temp)

        atividades = extrair_atividades(nome_temp, companhia)
        arquivo_saida = gerar_pdf(atividades)

        return send_file(arquivo_saida, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
