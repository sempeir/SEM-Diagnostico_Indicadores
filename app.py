from flask import Flask, request, send_file, jsonify
import pandas as pd
import io

app = Flask(__name__)

# -----------------------------
# Detectar columnas automáticamente
# -----------------------------
def detectar_columna(df, posibles):

    for col in df.columns:
        for p in posibles:
            if p.lower() in col.lower():
                return col
    return None


# -----------------------------
# Clasificación DANE básica
# -----------------------------
def clasificar_indicador(texto):

    t = texto.lower()

    if any(x in t for x in ["porcentaje","tasa","reducción","incremento"]):
        return "Resultado"

    if any(x in t for x in ["número","cantidad","realizadas","entregados"]):
        return "Producto"

    return "Impacto"


# -----------------------------
# Evaluación metodológica
# -----------------------------
def evaluar(objetivo, indicador):

    score = 0
    obs = []

    if len(indicador) > 15:
        score += 25
    else:
        obs.append("Indicador poco específico")

    if "%" in indicador.lower() or "tasa" in indicador.lower():
        score += 25
    else:
        obs.append("No mide cambio relativo")

    if objetivo.split(" ")[0].lower() in indicador.lower():
        score += 25
    else:
        obs.append("Baja relación con objetivo")

    if "/" in indicador:
        score += 25
    else:
        obs.append("Fórmula no explícita")

    return score, ", ".join(obs)


# -----------------------------
# ENDPOINT
# -----------------------------
@app.route("/analizar", methods=["POST"])
def analizar():

    try:

        file = request.files["file"]

        df = pd.read_excel(file, engine="openpyxl")

        col_obj = detectar_columna(df,
            ["objetivo","objetivo específico","objetivo general"])

        col_ind = detectar_columna(df,
            ["indicador","indicador propuesto","nombre indicador"])

        if not col_obj or not col_ind:
            return jsonify({"error":
                "No se detectaron columnas de objetivo o indicador"}),400

        resultados = []

        for _, row in df.iterrows():

            objetivo = str(row[col_obj])
            indicador = str(row[col_ind])

            tipo_auto = clasificar_indicador(indicador)
            score, obs = evaluar(objetivo, indicador)

            sugerido = indicador

            if tipo_auto == "Producto":
                sugerido = "Porcentaje de beneficiarios que presentan mejora asociada al objetivo"

            resultados.append({
                "Objetivo": objetivo,
                "Indicador original": indicador,
                "Tipo detectado": tipo_auto,
                "Puntaje calidad": score,
                "Diagnóstico": obs,
                "Indicador sugerido SEM": sugerido
            })

        diagnostico = pd.DataFrame(resultados)

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            diagnostico.to_excel(writer,
                                 index=False,
                                 sheet_name="Diagnóstico SEM")

        output.seek(0)

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Diagnostico_SEM.xlsx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}),500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
