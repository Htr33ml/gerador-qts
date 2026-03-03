from flask import Flask, render_template, request, send_file
import pdfplumber
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
import uuid
import os
from collections import defaultdict

app = Flask(__name__)

def extrair_tabela_dsi(pdf_path):
    atividades = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and len(row) >= 6:
                        atividades.append({
                            "data": row[0],
                            "hora": row[1],
                            "atividade": row[2],
                            "local": row[3],
                            "participantes": row[4],
                            "uniforme": row[5]
                        })
    return atividades


def filtrar_companhia(atividades, companhia):
    filtradas = []
    for a in atividades:
        participantes = str(a["participantes"])
        if companhia in participantes or "Todos" in participantes:
            filtradas.append(a)
    return filtradas


def organizar_por_dia(atividades):
    organizadas = defaultdict(list)
    for a in atividades:
        organizadas[a["data"]].append(a)

    # ordenar por hora dentro de cada dia
    for dia in organizadas:
        organizadas[dia].sort(key=lambda x: x["hora"] or "")
    return organizadas


def gerar_pdf_qts(atividades_por_dia):
    nome = f"qts_{uuid.uuid4()}.pdf"
    doc = SimpleDocTemplate(nome, pagesize=A4)
    elementos = []

    styles = getSampleStyleSheet()
    elementos.append(Paragraph("<b>QUADRO DE TRABALHO SEMANAL</b>", styles["Title"]))
    elementos.append(Spacer(1, 0.3 * inch))

    for dia, atividades in atividades_por_dia.items():
        elementos.append(Paragraph(f"<b>{dia}</b>", styles["Heading2"]))
        elementos.append(Spacer(1, 0.2 * inch))

        dados = [["HORA", "ATIVIDADE", "LOCAL", "PARTICIPANTES", "UNIFORME"]]

        for a in atividades:
            dados.append([
                a["hora"],
                a["atividade"],
                a["local"],
                a["participantes"],
                a["uniforme"]
            ])

        tabela = Table(dados, repeatRows=1)
        tabela.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ]))

        elementos.append(tabela)
        elementos.append(Spacer(1, 0.4 * inch))

    doc.build(elementos)
    return nome


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        arquivo = request.files["dsi"]
        companhia = request.form["companhia"]

        nome_temp = f"temp_{uuid.uuid4()}.pdf"
        arquivo.save(nome_temp)

        atividades = extrair_tabela_dsi(nome_temp)
        atividades_filtradas = filtrar_companhia(atividades, companhia)
        atividades_organizadas = organizar_por_dia(atividades_filtradas)

        arquivo_saida = gerar_pdf_qts(atividades_organizadas)

        return send_file(arquivo_saida, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
