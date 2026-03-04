from flask import Flask, render_template, request, send_file
from docx import Document
import os

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        semana = request.form.get("semana")
        periodo = request.form.get("periodo")
        companhia = request.form.get("companhia")

        doc = Document("modelo_qts.docx")

        # substituir placeholders
        for p in doc.paragraphs:
            if "{{SEMANA}}" in p.text:
                p.text = p.text.replace("{{SEMANA}}", semana)

            if "{{PERIODO}}" in p.text:
                p.text = p.text.replace("{{PERIODO}}", periodo)

            if "{{COMPANHIA}}" in p.text:
                p.text = p.text.replace("{{COMPANHIA}}", companhia)

            if "{{DIAS_SEMANA}}" in p.text:
                p.text = ""

        # criar exemplo de tabela (teste inicial)
        doc.add_heading("SEGUNDA-FEIRA", level=2)

        table = doc.add_table(rows=2, cols=3)

        table.cell(0,0).text="HORA"
        table.cell(0,1).text="ATIVIDADE"
        table.cell(0,2).text="LOCAL"

        table.cell(1,0).text="0800"
        table.cell(1,1).text="Treinamento físico"
        table.cell(1,2).text="Pátio"

        arquivo = "QTS.docx"
        doc.save(arquivo)

        return send_file(arquivo, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
