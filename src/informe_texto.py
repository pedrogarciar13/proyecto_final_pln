"""Supuestos del caso de valor y datos del equipo (compartidos por informe, presentación y guion)."""

EQUIPO = [("Pedro García Romero", "Especialista en Machine Learning"),
          ("Carlos Trilleras", "Magíster en Ciencia de Datos"),
          ("Fabiola Aldana Mora", "Magíster en Gerencia en IA")]
POR_ROL = {rol: nombre for nombre, rol in EQUIPO}
ML, DS, GIA = "Especialista en Machine Learning", "Magíster en Ciencia de Datos", "Magíster en Gerencia en IA"

# Supuestos ilustrativos del caso de valor (no son cifras internas de Mercado Libre)
SUPUESTOS = {
    "volumen_mes": 1_000_000,          # comentarios al mes en un país mediano
    "pct_negativos": 0.40,
    "muestreo_actual": 0.10,           # hoy se lee el 10% por muestreo
    "comentarios_hora_agente": 60,
    "costo_hora_agente": 8.0,          # US$ cargados
    "costo_devolucion": 12.0,          # US$ por devolución
    "tasa_evitada": 0.05,              # devoluciones evitadas por negativo atendido a tiempo
    "respuesta_actual_h": 72,          # línea base del KPI: horas hasta la primera respuesta a un negativo
    "monitoreo_mes": 1500.0,           # media persona al mes para monitoreo y reentrenamiento
    "implementacion": {"Equipo de proyecto (3 personas, 2 meses)": 18000.0,
                       "Reetiquetado de 5.000 comentarios (neutros y cortos)": 500.0,
                       "Revisión legal de licencias y privacidad": 2000.0,
                       "GPU en la nube para reentrenar con toda la base (20 h)": 30.0},
}


def caso_valor(cobertura, recall_ef, usd_millon, S=SUPUESTOS, tasa=None):
    """Beneficio mensual del muestreo actual y del modelo con abstención (US$)."""
    tasa = S["tasa_evitada"] if tasa is None else tasa
    v, neg = S["volumen_mes"], S["volumen_mes"] * S["pct_negativos"]
    costo_rev = lambda leidos: leidos / S["comentarios_hora_agente"] * S["costo_hora_agente"]
    actual = {"leidos": v * S["muestreo_actual"], "detectados": neg * S["muestreo_actual"]}
    actual["revision"] = costo_rev(actual["leidos"])
    actual["evitadas"] = actual["detectados"] * tasa
    actual["beneficio"] = actual["evitadas"] * S["costo_devolucion"]
    actual["computo"] = 0.0
    actual["neto"] = actual["beneficio"] - actual["revision"]
    modelo = {"leidos": v * (1 - cobertura), "detectados": neg * recall_ef}
    modelo["revision"] = costo_rev(modelo["leidos"])
    modelo["evitadas"] = modelo["detectados"] * tasa
    modelo["beneficio"] = modelo["evitadas"] * S["costo_devolucion"]
    modelo["computo"] = usd_millon * v / 1e6 + S["monitoreo_mes"]
    modelo["neto"] = modelo["beneficio"] - modelo["revision"] - modelo["computo"]
    return actual, modelo
