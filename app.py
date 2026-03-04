from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import io

app = Flask(__name__)

dias_semana = ["SEGUNDA","TERÇA","QUARTA","QUINTA","SEXTA"]

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

        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer, pagesize=A4)

        y = 800

        pdf.setFont("Helvetica-Bold",14)
        pdf.drawString(150,y,"QUADRO DE TRABALHO SEMANAL")
        y -= 40

        pdf.setFont("Helvetica",11)
        pdf.drawString(50,y,f"Companhia: {companhia}")
        y -= 20
        pdf.drawString(50,y,f"Semana: {semana}")
        y -= 20
        pdf.drawString(50,y,f"Período: {periodo}")
        y -= 40

        for dia in dias_semana:

            pdf.setFont("Helvetica-Bold",12)
            pdf.drawString(50,y,dia)
            y -= 20

            if dia in dias_sem_exp:

                pdf.setFont("Helvetica",11)
                pdf.drawString(70,y,"SEM EXPEDIENTE")
                y -= 30
                continue


            pdf.setFont("Helvetica",10)

            if tiragem:
                pdf.drawString(70,y,"08:00 - 08:15  Tiragem de faltas")
                y -= 15

            if tfm:
                pdf.drawString(70,y,"08:15 - 09:30  Treinamento físico militar")
                y -= 15


            if dia == "SEXTA" and sexta_pacote:

                pdf.drawString(70,y,"08:00 - 10:00  Formatura")
                y -= 15

                pdf.drawString(70,y,"10:00 - 11:30  Faxina")
                y -= 15

                pdf.drawString(70,y,"11:30 - 12:00  Formatura final")
                y -= 15

            else:

                pdf.drawString(70,y,"09:30 - 15:30  Manutenção das instalações com Enc Mat")
                y -= 15

                if formatura_final:
                    pdf.drawString(70,y,"15:30 - 16:00  Formatura de liberação")
                    y -= 15


            y -= 20

            if y < 100:
                pdf.showPage()
                y = 800


        pdf.save()

        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name="QTS.pdf",
            mimetype="application/pdf"
        )

    return render_template("index.html")


if __name__ == "__main__":
    app.run()
