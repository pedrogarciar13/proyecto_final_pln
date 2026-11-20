"""Demo web del clasificador de sentimiento (Flask).

Local:   python app/app.py            → http://localhost:5000  (usa BETO si existe modelos/beto_sentimiento)
Render:  gunicorn app.app:app         → usa TF-IDF + Regresión Logística (BETO no cabe en el plan gratuito)
Variable MODELO_DEMO para forzar el modelo ("BETO", "BETO int8", "GRU", "TF-IDF + Regresión Logística").
"""
import os
import sys

from flask import Flask, jsonify, render_template, request

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))
from servicio import Servicio  # noqa: E402

app = Flask(__name__)
srv = Servicio()

EJEMPLOS = [
    "Lo usé una semana y dejó de funcionar. La calidad es pésima, el material se siente barato y no lo recomiendo para nada.",
    "Excelente producto, lo recomiendo",
    "Excelente producto, se nota la buena calidad del material, funciona tal como dice la descripción y lo recomiendo a quien lo necesite",
    "La tela es buena y el diseño es bonito, aunque la talla es más pequeña de lo esperado y el color es más claro que en la foto.",
    "Dejó de funcionar a la semana y la calidad es muy mala, escríbanme a juan.perez@correo.com",
    "El paquete llegó tarde y el vendedor nunca respondió mis mensajes, pésimo servicio de entrega",
    "El partido terminó dos a uno y el delantero fue expulsado en el segundo tiempo por una falta muy fuerte",
]


@app.get("/")
def inicio():
    return render_template("index.html", modelo=srv.nombre, umbral=srv.umbral, ejemplos=EJEMPLOS)


@app.post("/api/predecir")
def api_predecir():
    texto = (request.get_json(silent=True) or {}).get("texto", "")
    return jsonify(srv.predecir(texto[:3000]))


@app.get("/salud")
def salud():
    return {"estado": "ok", "modelo": srv.nombre}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
