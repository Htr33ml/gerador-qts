from flask import Flask, render_template, request, send_file
from docx import Document
import io

app = Flask(__name__)

dias_semana = ["SEGUNDA","TERÇA","QUARTA","QUINTA","SEXTA"]


def gerar_linhas_dia(dia, tiragem, tfm, formatura_final, sexta_pacote, sem_exp):

    linhas = ""

    if dia in sem_exp:
        linhas += f"{dia} - SEM EXPEDIENTE\n"
        return linhas


    if tiragem:
        linhas += "08:00 - 08:15 | Tiragem de faltas\n"

    if tfm:
        linhas += "08:15 - 09:30 | Treinamento físico militar\n"


    if dia == "SEXTA" and sexta_pacote:

        linhas += "08:00 - 10:00 | Formatura\n"
        linhas += "10:00 - 11:30 | Faxina\n"
        linhas += "11:30 - 12:00 | Formatura final\n"

    else:

        linhas += "09:30 - 15:30 | Manutenção das instalações com Enc Mat\n"

        if formatura_final:
            linhas += "15:30 - 16:00 | Formatura de liberação\n"

    return linhas


@app.route("/", methods=["GET","POST"])
def index():

    if request.method == "POST":

        semana = request.form.get("semana")
        periodo = request.form.get("periodo")
        companhia = request.form.get("companhia")

        tiragem = request.form.get("tiragem_faltas")
        tfm = request.form.get("tfm")
        formatura_final = request.form.get("formatura_final")
        sexta_pacote = request.form.get("sexta_pacote")

        dias_sem_exp = request.form.getlist("sem_expediente")

        document = Document("modelo_qts.docx")


        texto_dias = ""

        for dia in dias_semana:

            texto_dias += f"\n{dia}\n"

            linhas = gerar_linhas_dia(
                dia,
                tiragem,
                tfm,
                formatura_final,
                sexta_pacote,
                dias_sem_exp
            )

            texto_dias += linhas


        for p in document.paragraphs:

            if "{{SEMANA}}" in p.text:
                p.text = p.text.replace("{{SEMANA}}", semana)

            if "{{PERIODO}}" in p.text:
                p.text = p.text.replace("{{PERIODO}}", periodo)

            if "{{COMPANHIA}}" in p.text:
                p.text = p.text.replace("{{COMPANHIA}}", companhia)

            if "{{DIAS_SEMANA}}" in p.text:
                p.text = p.text.replace("{{DIAS_SEMANA}}", texto_dias)


        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name="QTS.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


    return render_template("index.html")


if __name__ == "__main__":
    app.run()
