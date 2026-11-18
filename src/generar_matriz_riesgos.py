"""Genera docs/Matriz_Riesgos.xlsx a partir de los resultados del notebook.  Uso: python src/generar_matriz_riesgos.py"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(__file__))
from resumen import cargar, cifras, corto, n, pct  # noqa: E402

RAIZ = sys.argv[1] if len(sys.argv) > 1 else "."
r = cargar(RAIZ)
f = cifras(r)
s4 = r["s4"]
paises = s4["segmentos"]["country"]
peor_pais = min(paises, key=lambda p: paises[p]["BETO"])

NIVEL = {"Baja": 1, "Media": 2, "Alta": 3}
riesgos = [
    ("R01", "Datos", "Etiqueta ruidosa en el neutro",
     "El sentimiento se deriva de las estrellas; un 3★ puede contener una queja o un elogio.",
     "Alta", "Alto",
     "Reetiquetado humano de una muestra mensual de 300 neutros; abstención en baja confianza; métrica por clase.",
     "Ciencia de Datos",
     f"F1 neutro {n(f['f1_neu'][f['bert']])} ({corto(f['bert'])}) frente a ~0,85 en negativo/positivo.",
     "F1 neutro mensual sobre la muestra reetiquetada", "Media"),
    ("R02", "Datos", "Fuga de información (título, estrellas, duplicados)",
     "El título y las estrellas se escriben junto con la calificación; textos repetidos entre conjuntos.",
     "Media", "Alto",
     "Excluir review_title y review_rate; deduplicar antes de partir; vocabulario/IDF/tokenizador ajustados solo con train.",
     "Ciencia de Datos",
     f"F1 solo título {n(f['audit']['f1_solo_titulo'])} > solo comentario {n(f['audit']['f1_solo_comentario'])}; aserción de no solapamiento en el notebook.",
     "Prueba automática de solapamiento train/test en cada reentrenamiento", "Baja"),
    ("R03", "Privacidad", "Exposición de PII en texto libre",
     "Comentarios con correos, teléfonos o URL que pueden llegar a logs o a terceros.",
     "Baja", "Alto",
     "Enmascaramiento <EMAIL>/<TEL>/<URL> antes de inferir y registrar; acceso por roles; retención de 90 días; sin envío a servicios externos.",
     "Gerencia de IA",
     f"PII detectada en la base: {f['audit'].get('pii')}.",
     "% de textos con PII sin enmascarar en auditoría trimestral", "Baja"),
    ("R04", "Modelo", "Falso negativo crítico (negativo leído como positivo)",
     "El cliente molesto no recibe atención y se pierde la alerta al vendedor.",
     "Media", "Alto",
     "Umbral de abstención elegido en validación para recall de negativo ≥ 0,92; revisión humana de casos dudosos.",
     "Machine Learning",
     f"Negativos leídos como positivos en prueba común: {s4['neg_como_pos'].get(f['bert'])} ({corto(f['bert'])}) vs {s4['neg_como_pos'].get(f['base'])} (base).",
     "Recall de negativo semanal en modo sombra", "Media"),
    ("R05", "Sesgo", "Desempeño desigual por país",
     "Perú y Venezuela tienen menos del 2% de los datos; variaciones regionales del español.",
     "Media", "Medio",
     "Reportar F1 por país; mínimo de F1 por país antes de activar; muestreo adicional de países minoritarios.",
     "Ciencia de Datos",
     f"País con menor F1 (BETO, prueba común): {peor_pais} = {n(paises[peor_pais]['BETO'])}.",
     "F1 macro por país en el tablero mensual", "Media"),
    ("R13", "Sesgo", "Sesgo de longitud: comentarios cortos leídos como negativos",
     "En el corpus casi todos los comentarios de ≤ 15 palabras son negativos; los modelos aprendieron el atajo “corto ⇒ negativo”.",
     "Alta", "Alto",
     "Sin alertas automáticas al vendedor para textos de ≤ 15 palabras (control en src/demo.py); reentrenar con comentarios cortos "
     "balanceados (chats etiquetados por agentes); reportar métricas por longitud.",
     "Ciencia de Datos",
     (lambda S: f"{pct(S['pct_negativo_le15'], 0)} de los cortos son negativos; recall de positivos cortos de BETO "
      f"{pct(next(m['recall positivo'] for m in S['modelos'] if m['modelo'] == 'BETO' and m['segmento'].startswith('≤')), 0)}."
      )(f["sesgo"]) if f["sesgo"] else "Ver sección 8.1 del notebook.",
     "Recall de positivos y precisión de alertas en textos cortos (muestra semanal)", "Media"),
    ("R06", "Operación", "Deriva del lenguaje (drift)",
     "Productos nuevos, jerga, temporadas (Hot Sale, Black Friday) y cambio de canal (chats, reclamos).",
     "Alta", "Medio",
     "Monitorear % OOV/subpalabras por palabra, distribución de confianza y tasa de abstención; reentrenar si la abstención sube 10 pp.",
     "Machine Learning",
     f"OOV en validación del vocabulario RNN: {pct(f['audit'].get('oov_val'), 2)}.",
     "Tasa de abstención, OOV y distribución de clases semanal", "Media"),
    ("R07", "Operación", "Latencia o costo de BETO en CPU",
     "Transformer de 110 M de parámetros sin GPU.",
     "Media", "Medio",
     "Cuantización int8, lotes, max_length por percentiles; procesamiento por lotes horarios; TF-IDF como respaldo.",
     "Machine Learning",
     f"p95 1 comentario: BETO {n(f['p95'].get('BETO'), 1)} ms; BETO int8 {n(f['p95'].get('BETO int8'), 1)} ms; TF-IDF {n(f['p95'][f['base']], 2)} ms.",
     "Latencia p95 y costo por millón de comentarios", "Baja"),
    ("R08", "Modelo", "Sobreconfianza (probabilidades mal calibradas)",
     "Softmax suele dar confianzas altas aun cuando se equivoca; la abstención depende de esa confianza.",
     "Media", "Medio",
     "Temperature scaling ajustado en validación; ECE en el tablero; recalibrar en cada reentrenamiento.",
     "Machine Learning",
     f"ECE {corto(f['bert'])}: {n(s4['calibracion'][f['bert']]['ece_antes'])} → {n(s4['calibracion'][f['bert']]['ece_despues'])} tras calibrar.",
     "ECE mensual", "Baja"),
    ("R09", "Modelo", "Entradas fuera de dominio o adversariales",
     "Texto sin relación con compras, spam, otro idioma o ironía.",
     "Media", "Bajo",
     "Validación de entrada (vacío/corto), abstención por baja confianza, pruebas contrafactuales como regresión automatizada.",
     "Machine Learning",
     f"Casos de robustez superados: {s4['robustez']}.",
     "Suite de contrafactuales en cada versión", "Baja"),
    ("R10", "Gobierno", "Licencia y trazabilidad de datos y checkpoint",
     "MeLiSA tiene licencia 'other'; BETO CC BY 4.0; uso comercial no verificado.",
     "Baja", "Alto",
     "Uso académico; revisión legal antes de producción; versionar checkpoint, tokenizer, datos, split, semilla y umbral (model card).",
     "Gerencia de IA",
     "LICENSE y data/README.md documentan origen y licencia.",
     "Checklist de aprobación de uso previsto", "Baja"),
    ("R11", "Adopción", "Automatización sin control humano",
     "Uso de la etiqueta para sancionar vendedores o compradores.",
     "Baja", "Alto",
     "El modelo solo prioriza y alerta; ninguna sanción automática; agente confirma antes de actuar; registro de decisiones.",
     "Gerencia de IA",
     "Política de uso en el informe (sección 6) y en la data card (usos no previstos).",
     "% de alertas confirmadas por un agente", "Baja"),
    ("R12", "Seguridad", "Dependencia de descargas externas (Hugging Face)",
     "El checkpoint o el tokenizador cambian o dejan de estar disponibles.",
     "Baja", "Medio",
     "Fijar versión/commit del checkpoint, copia interna de artefactos, sin credenciales en el código (.env.example).",
     "Machine Learning",
     "requirements.txt con versiones; artefactos en modelos/.",
     "Hash del artefacto desplegado", "Baja"),
]

wb = Workbook()
ws = wb.active
ws.title = "Matriz de riesgos"
enc = ["ID", "Categoría", "Riesgo", "Causa", "Probabilidad", "Impacto", "Nivel inherente", "Control", "Responsable",
       "Evidencia (del notebook)", "Indicador de monitoreo", "Riesgo residual"]
anchos = [6, 12, 30, 40, 13, 10, 12, 52, 17, 48, 34, 13]
azul = PatternFill("solid", fgColor="0B2E59")
borde = Border(*(Side(style="thin", color="CBD5E1"),) * 4)
colores = {"Bajo": "D1FAE5", "Baja": "D1FAE5", "Medio": "FEF3C7", "Media": "FEF3C7", "Alto": "FEE2E2", "Alta": "FEE2E2"}
ws.append(["Matriz de riesgos · Sentimiento en reseñas de Mercado Libre (MeLiSA)"])
ws["A1"].font = Font(bold=True, size=14, color="0B2E59")
ws.append([f"Modelo recomendado: {f['rec']} · Probabilidad e impacto: Baja/Media/Alta · Nivel = probabilidad × impacto (1–9)"])
ws["A2"].font = Font(italic=True, color="475569")
ws.append(enc)
for i, w in enumerate(anchos, 1):
    ws.column_dimensions[chr(64 + i)].width = w
    c = ws.cell(3, i)
    c.fill, c.font, c.alignment, c.border = azul, Font(bold=True, color="FFFFFF"), Alignment(wrap_text=True, vertical="center"), borde
for fila in riesgos:
    rid, cat, ries, causa, prob, imp, ctrl, resp, evid, ind, resid = fila
    nivel = NIVEL[prob] * NIVEL[{"Bajo": "Baja", "Medio": "Media", "Alto": "Alta"}[imp]]
    ws.append([rid, cat, ries, causa, prob, imp, nivel, ctrl, resp, evid, ind, resid])
    k = ws.max_row
    for i in range(1, len(enc) + 1):
        c = ws.cell(k, i)
        c.alignment, c.border = Alignment(wrap_text=True, vertical="top"), borde
    for i, v in [(5, prob), (6, imp), (12, resid)]:
        ws.cell(k, i).fill = PatternFill("solid", fgColor=colores[v])
    ws.cell(k, 7).fill = PatternFill("solid", fgColor="FEE2E2" if nivel >= 6 else "FEF3C7" if nivel >= 3 else "D1FAE5")
    ws.cell(k, 3).font = Font(bold=True)
dv = DataValidation(type="list", formula1='"Baja,Media,Alta"', allow_blank=False)
ws.add_data_validation(dv); dv.add(f"E4:E{ws.max_row}"); dv.add(f"L4:L{ws.max_row}")
ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:L{ws.max_row}"

ws2 = wb.create_sheet("Suspensión y rollback")
filas = [("Condición de suspensión o rollback", "Umbral", "Acción", "Responsable"),
         ("Recall de negativo en modo sombra", "< 0,88 durante 2 semanas", "Volver a TF-IDF + LR y reentrenar", "Machine Learning"),
         ("Tasa de abstención", "> 35% durante 1 semana", "Revisar drift; ampliar equipo de revisión", "Gerencia de IA"),
         ("F1 macro por país (muestra reetiquetada)", "< 0,65 en cualquier país", "Desactivar alertas automáticas en ese país", "Ciencia de Datos"),
         ("Latencia p95", "> 150 ms en línea", "Pasar a procesamiento por lotes o a la versión int8", "Machine Learning"),
         ("Incidente de privacidad", "Cualquier PII expuesta en logs", "Suspender, purgar logs, auditoría", "Gerencia de IA"),
         ("Quejas de vendedores por alertas falsas", "Precisión de alertas < 80%", "Subir el umbral y revisar casos", "Gerencia de IA")]
for fila in filas:
    ws2.append(fila)
for i, w in enumerate([42, 30, 50, 20], 1):
    ws2.column_dimensions[chr(64 + i)].width = w
    ws2.cell(1, i).fill, ws2.cell(1, i).font = azul, Font(bold=True, color="FFFFFF")

ws3 = wb.create_sheet("Roadmap")
for fila in [("Fase", "Duración", "Qué ocurre", "Criterio para avanzar"),
             ("0. Preparación", "2 semanas", "Revisión legal de licencia, model card, tablero de monitoreo", "Aprobación del uso previsto"),
             ("1. Modo sombra", "4 semanas", "El modelo predice sin actuar; se compara con estrellas y con revisión humana", "Recall negativo ≥ 0,92 y cobertura ≥ 75%"),
             ("2. Humano en el ciclo", "6 semanas", "Negativos a cola prioritaria; un agente confirma antes de actuar", "Precisión de alertas ≥ 80%; tiempo de respuesta < 24 h"),
             ("3. Alertas automáticas", "Continuo", "Alerta al vendedor cuando sube el % de negativos de un producto", "Sin incidentes 4 semanas"),
             ("4. Mejora continua", "Trimestral", "Reentrenamiento, reetiquetado del neutro, sentimiento por aspecto", "Criterio de adopción vigente")]:
    ws3.append(fila)
for i, w in enumerate([24, 14, 60, 44], 1):
    ws3.column_dimensions[chr(64 + i)].width = w
    ws3.cell(1, i).fill, ws3.cell(1, i).font = azul, Font(bold=True, color="FFFFFF")
for w in (ws2, ws3):
    for row in w.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

os.makedirs(os.path.join(RAIZ, "docs"), exist_ok=True)
salida = os.path.join(RAIZ, "docs", "Matriz_Riesgos.xlsx")
wb.save(salida)
print("Matriz de riesgos ->", salida)

