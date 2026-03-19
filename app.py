from flask import Flask, request, send_file, jsonify
import pandas as pd
import io

app = Flask(__name__)

# -----------------------------
# Evaluación básica indicador
# -----------------------------
def evaluar_indicador(row):

    indicador = str(row.get("Indicador",""))
    formula = str(row.get("Fórmula",""))
    fuente = str(row.get("Fuente",""))

    score = 0
    observaciones = []

    if "%" in indicador.lower() or "tasa" in indicador.lower():
        score += 25
    else:
        observaciones.append("Indicador no expresa medida relativa")

    if "/" in formula:
        score += 25
    else:
        observaciones.append("Fórmula incompleta")

    if len(indicador) > 15:
        score += 25
    else:
        observaciones.append("Indicador poco descriptivo")

    if fuente.strip():
        score += 25
    else:
        observaciones.append("Sin fuente definida")

    return score, ", ".join(observaciones)


# -----------------------------
# Endpoint análisis
# -----------------------------
@app.route("/analizar", methods=["POST"])
def analizar():

    try:

        if "file" not in request.files:
            return jsonify({"error":"Archivo no recibido"}),400

        file = request.files["file"]

        df = pd.read_excel(file, engine="openpyxl")

        resultados = []

        for _, row in df.iterrows():
            score, obs = evaluar_indicador(row)

            resultados.append({
                "Indicador": row.get("Indicador"),
                "Puntaje": score,
                "Diagnóstico": obs
            })

        diagnostico = pd.DataFrame(resultados)

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            diagnostico.to_excel(
                writer,
                index=False,
                sheet_name="Diagnostico SEM"
            )

        output.seek(0)

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Diagnostico_Indicadores_SEM.xlsx"
        )

    except Exception as e:
        return jsonify({"error":str(e)}),500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
