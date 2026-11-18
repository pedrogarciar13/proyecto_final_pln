"""Genera docs/Informe_Tecnico_Proyecto_PLN.docx a partir de resultados/resultados.json.

Uso: python src/generar_informe.py [raiz_del_proyecto]
"""
import os
import re
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(__file__))
from resumen import cargar, cifras, corto, entero, n, pct  # noqa: E402

RAIZ = sys.argv[1] if len(sys.argv) > 1 else "."
r = cargar(RAIZ)
f = cifras(r)
s4, A, E = r["s4"], f["audit"], f["entorno"]
BASE, S3, SUB, BERT, REC = f["base"], f["s3"], f["sub"], f["bert"], f["rec"]
FIG = os.path.join(RAIZ, "resultados", "figuras")
_cont = {"t": 0, "f": 0}
NAVY = RGBColor(0x0B, 0x2E, 0x59)
TEAL = RGBColor(0x0F, 0xA3, 0xB1)

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
sec.left_margin = sec.right_margin = Cm(2.5)
sec.top_margin = sec.bottom_margin = Cm(2.3)
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.12
for nivel, tam in [(1, 16), (2, 13), (3, 11.5)]:
    h = doc.styles[f"Heading {nivel}"]
    h.font.name, h.font.size, h.font.bold, h.font.color.rgb = "Calibri", Pt(tam), True, NAVY
    h.element.rPr.rFonts.set(qn("w:asciiTheme"), "")
    h.paragraph_format.space_before = Pt(14 if nivel == 1 else 10)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True


def p(texto="", negrita_inicio=None, italic=False, size=None, align=None):
    par = doc.add_paragraph()
    if negrita_inicio:
        run = par.add_run(negrita_inicio)
        run.bold = True
    run = par.add_run(texto)
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    if align:
        par.alignment = align
    return par


def viñetas(items):
    for it in items:
        par = doc.add_paragraph(style="List Bullet")
        if isinstance(it, tuple):
            par.add_run(it[0]).bold = True
            par.add_run(it[1])
        else:
            par.add_run(it)


def sombrear(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def tabla(filas, anchos=None, size=9.5, titulo=None):
    if titulo:
        _cont["t"] += 1
        titulo = re.sub(r"^Tabla \d+\.", f"Tabla {_cont['t']}.", titulo)
        cap = doc.add_paragraph()
        run = cap.add_run(titulo)
        run.bold, run.font.size, run.font.color.rgb = True, Pt(9.5), NAVY
        cap.paragraph_format.keep_with_next = True
        cap.paragraph_format.space_after = Pt(2)
    t = doc.add_table(rows=len(filas), cols=len(filas[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, fila in enumerate(filas):
        for j, v in enumerate(fila):
            c = t.cell(i, j)
            c.text = ""
            run = c.paragraphs[0].add_run(str(v))
            run.font.size = Pt(size)
            run.bold = i == 0
            if i == 0:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                sombrear(c, "0B2E59")
            elif i % 2 == 0:
                sombrear(c, "F1F5F9")
            c.paragraphs[0].paragraph_format.space_after = Pt(0)
            if anchos:
                c.width = Cm(anchos[j])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figura(nombre, ancho_cm, pie):
    ruta = os.path.join(FIG, nombre)
    if not os.path.exists(ruta):
        return
    if re.match(r"^Figura \d+\.", pie):
        _cont["f"] += 1
        pie = re.sub(r"^Figura \d+\.", f"Figura {_cont['f']}.", pie)
    doc.add_picture(ruta, width=Cm(ancho_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = p(pie, italic=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    cap.runs[0].font.color.rgb = RGBColor(0x64, 0x74, 0x8B)


def salto():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


ab = {k: (v or {}) for k, v in f["abst"].items()}
ab_rec = ab.get(REC, {})
comun = s4["comun"]
res3 = r["resultados_s3"]
ic_bs3 = f["ic"].get(f"BETO - {S3}", (0, 0, 0))
signif = ic_bs3[1] > 0 or ic_bs3[2] < 0
D = f["decision"]
SES = f["sesgo"] or {}
if D:
    decision = D["texto"]
elif f["cumple_bert"]:
    decision = (f"Pilotear {REC} en modo sombra durante cuatro semanas, con abstención (umbral {n(f['umbral'][REC], 2)} sobre "
                f"probabilidades calibradas) y revisión humana de los casos dudosos; TF-IDF + Regresión Logística queda como respaldo "
                f"explicable y como plan de reversión.")
else:
    decision = (f"No adoptar todavía el Transformer porque no cumple los cuatro criterios de la sesión 4. Mantener {REC} con abstención "
                f"(umbral {n(f['umbral'].get(REC), 2)}) en modo sombra y repetir el experimento de BETO con GPU y con la base completa.")

# ------------------------------------------------------------------ portada
for _ in range(5):
    doc.add_paragraph()
p("UNIVERSIDAD EAN · PROCESAMIENTO DE LENGUAJE NATURAL", size=11).runs[-1].font.color.rgb = TEAL
t = p("Análisis de sentimiento en reseñas de Mercado Libre", size=26)
t.runs[-1].bold = True
t.runs[-1].font.color.rgb = NAVY
p("Línea base TF-IDF, modelos secuenciales (RNN, LSTM, GRU) y un Transformer preentrenado en español (BETO): diseño, "
  "evaluación, gobierno y recomendación de adopción", size=13)
doc.add_paragraph()
p("Informe técnico del proyecto final", size=12).runs[-1].bold = True
p("Estudiante: Jose Romero", size=11)
p("Docente: Gerson José Moreno Morales", size=11)
p("Noviembre de 2026", size=11)
doc.add_paragraph()
p("Perfiles integrados: Especialista en Machine Learning · Magíster en Ciencia de Datos · Magíster en Gerencia en IA", italic=True, size=10)
salto()

# ------------------------------------------------------------------ 1. Resumen
doc.add_heading("1. Resumen", 1)
p(f"Mercado Libre recibe un volumen de reseñas, preguntas y reclamos imposible de leer a mano, y solo las reseñas traen estrellas. "
  f"Este proyecto construye un clasificador de sentimiento (negativo, neutro, positivo) que lee únicamente el texto del comentario, "
  f"para priorizar la atención de los compradores molestos y alertar a los vendedores. Se usaron {entero(A['filas'])} reseñas públicas "
  f"del corpus MeLiSA (7 países, 5 categorías). Tras la auditoría quedaron {entero(A['filas_limpias'])}; se excluyeron el título y las "
  f"estrellas porque filtran la respuesta.")
p(f"Se compararon ocho modelos sobre la misma partición: TF-IDF con regresión logística y SVM (línea base), RNN, LSTM y GRU (sesión 3) y "
  f"BETO, un BERT preentrenado en español ajustado con {entero(f['n_train_bert'])} comentarios, en precisión completa y cuantizado a int8 "
  f"(sesión 4). Sobre una prueba común de {entero(A['n_comun'])} comentarios, {corto(BERT)} obtuvo F1 macro {n(f['f1'][BERT])}, frente a "
  f"{n(f['f1'][S3])} del mejor recurrente ({S3}) y {n(f['f1'][BASE])} de la línea base. La diferencia con {S3} es "
  f"{'+' if f['delta_bert_s3'] >= 0 else ''}{n(f['delta_bert_s3'])} (IC 95% [{n(ic_bs3[1])}; {n(ic_bs3[2])}]). "
  f"La clase neutra es el cuello de botella de todos los modelos, en buena parte por ruido de la etiqueta.")
p(decision, negrita_inicio="Recomendación: ")

# ------------------------------------------------------------------ 2. Contexto
doc.add_heading("2. Contexto y problema", 1)
doc.add_heading("2.1 Organización, usuario y decisión", 2)
p("Mercado Libre es el principal comercio electrónico de Latinoamérica. La reputación del vendedor y del producto influye en el ranking "
  "de búsqueda, en la conversión y en la permanencia de los vendedores; un comprador insatisfecho que nadie atiende termina en "
  "devolución, reclamo o pérdida del cliente.")
tabla([["Elemento", "Definición"],
       ["Usuario", "Analista de experiencia del comprador (Atención al Cliente / Reputación de vendedores)"],
       ["Decisión apoyada", "Qué comentario se atiende primero (cola de prioridad) y a qué vendedor se envía una alerta temprana"],
       ["Proceso actual", "Lectura por muestreo; las estrellas solo existen en reseñas, no en chats, preguntas ni reclamos"],
       ["Entrada / unidad", "Un comentario de texto libre en español; flujo continuo procesado en lotes horarios"],
       ["Salida", "Etiqueta + probabilidad calibrada + decisión automática o revisión humana + evidencia"],
       ["Clase crítica", "Negativo. Error más costoso: negativo leído como positivo (se pierde la alerta)"],
       ["Hipótesis de valor", "Detectar automáticamente los negativos reduce el tiempo de primera respuesta y las devoluciones"]],
      [4, 12.5], titulo="Tabla 1. Formulación del problema")
doc.add_heading("2.2 KPI y alcance", 2)
viñetas([("KPI de negocio: ", "tiempo de primera respuesta a comentarios negativos (meta < 24 h), porcentaje de alertas confirmadas por un "
          "agente (meta ≥ 80%) y reclamos/devoluciones evitados en productos con alerta temprana."),
         ("Métrica técnica principal: ", "F1 macro, porque pesa igual las tres clases; se acompaña con recall de negativo, cobertura "
          "automática bajo abstención, calibración (ECE) y latencia."),
         ("Dentro del alcance: ", "clasificación de sentimiento de un comentario en español, abstención y evidencia."),
         ("Fuera del alcance: ", "sanciones automáticas, sentimiento por aspecto (propuesto como siguiente fase), portugués y otros canales "
          "sin una nueva evaluación.")])
p("¿Por qué predecir si ya existen las estrellas? Porque los chats, preguntas y reclamos no las tienen; porque el texto llega antes que "
  "la calificación final; y porque las estrellas sirven como etiquetas gratuitas para entrenar (supervisión débil).")

# ------------------------------------------------------------------ 3. Datos
doc.add_heading("3. Datos", 1)
doc.add_heading("3.1 Origen, permisos y taxonomía", 2)
p("MeLiSA (Mercado Libre for Sentiment Analysis) es un corpus público del Laboratorio de Procesamiento de Señales y Comunicaciones de la "
  "Universidad de Buenos Aires, disponible en Hugging Face (lpsc-fiuba/melisa) con licencia “other”; aquí se usa con fines académicos y no "
  "comerciales. Son datos reales, no sintéticos. El diccionario de datos y la data card completa están en data/README.md del repositorio.")
tabla([["Estrellas", "Sentimiento", "Proporción", "Acción de negocio"],
       ["1 – 2", "negativo", "40%", "Alerta temprana y cola prioritaria"],
       ["3", "neutro", "20%", "Oportunidad de mejora"],
       ["4 – 5", "positivo", "40%", "Testimonio y reputación"]], [3, 3, 3, 7.5], titulo="Tabla 2. Esquema de etiquetas")
doc.add_heading("3.2 Auditoría y calidad", 2)
pii = A.get("pii", {})
tabla([["Revisión", "Hallazgo", "Decisión"],
       ["Nulos", "0 en todas las columnas", "—"],
       ["Duplicados de texto", entero(A["duplicados_texto"]), "Se conserva una copia (evita fuga)"],
       ["Textos contradictorios", entero(A["textos_contradictorios"]), "Se eliminan"],
       ["Longitud (palabras)", f"P50 {n(A['palabras_p50'], 0)} · P95 {n(A['palabras_p95'], 0)} · P99 {n(A['palabras_p99'], 0)}",
        "RNN: 100 tokens"],
       ["Longitud (subpalabras BETO)", f"P50 {n(A['tokens_bert']['p50'], 0)} · P95 {n(A['tokens_bert']['p95'], 0)} · P99 {n(A['tokens_bert']['p99'], 0)}",
        f"max_length {f['max_len']}: trunca {pct(A['tokens_bert']['truncados_' + str(f['max_len'])])}"],
       ["PII", f"Correos {pii.get('EMAIL', 0)} · URL {pii.get('URL', 0)} · teléfonos {pii.get('TEL', 0)}", "Enmascarada (<EMAIL>, <URL>, <TEL>)"],
       ["Países", "MLA 43% · MLM 31% · MCO 10% · MLC 9% · MLU 5% · MLV 1% · MPE 0,7%", "Evaluación por país"],
       ["Clases", "40 / 20 / 40", "Pesos de clase balanceados"],
       ["Filas finales", entero(A["filas_limpias"]), "—"]], [4, 6.5, 6], titulo="Tabla 3. Auditoría del corpus")
figura("eda.png", 15.5, "Figura 1. Longitud de los comentarios por clase y distribución de clases por país.")
doc.add_heading("3.3 Control de fuga", 2)
p(f"Se entrenó el mismo TF-IDF + SVM con una sola columna: leyendo solo el título obtiene F1 macro {n(A['f1_solo_titulo'])}, más que "
  f"leyendo el comentario completo ({n(A['f1_solo_comentario'])}). El título se escribe junto con las estrellas y en producción no existe, "
  f"así que se excluye; review_rate también, porque es la fuente de la etiqueta. Además: deduplicación antes de partir (con aserción de "
  f"no solapamiento), vocabulario/IDF/tokenizador ajustados solo con entrenamiento, y umbral, temperatura y modelo campeón elegidos en "
  f"validación. Queda documentada una fuga menor: algunos compradores escriben su calificación en el texto (“doy tres estrellas”).")
doc.add_heading("3.4 Partición", 2)
p(f"Aleatoria y estratificada por sentimiento, semilla 42: {entero(A['n_train'])} de entrenamiento (70%), {entero(A['n_val'])} de "
  f"validación (15%) y {entero(A['n_test'])} de prueba (15%). No es posible una partición temporal ni por cliente porque el archivo no trae "
  f"fecha ni identificador de comprador; las reseñas son independientes. Como BETO en CPU no puede evaluar 73 mil comentarios en un tiempo "
  f"razonable, se extrajeron con la misma semilla una validación común y una prueba común estratificadas de {entero(A['n_comun'])} "
  f"comentarios cada una, sobre las que se comparan los ocho modelos.")

# ------------------------------------------------------------------ 4. Metodología
doc.add_heading("4. Metodología", 1)
doc.add_heading("4.1 Pipeline y reproducibilidad", 2)
p(f"Todo el flujo está en un único notebook que se ejecuta de inicio a fin (Proyecto_PLN_Sentimiento_MeLiSA.ipynb), con rutas relativas y "
  f"parámetros por variables de entorno (.env.example). Hardware: CPU de {E['nucleos_logicos']} hilos, {E['ram_gb'] / 1.0737:.0f} GB de RAM, "
  f"sin GPU; Python {E['python']}, scikit-learn {E['scikit-learn']}, TensorFlow {E['tensorflow']}, PyTorch {E['torch']}, Transformers "
  f"{E['transformers']}. Semilla 42 en Python, NumPy, TensorFlow y PyTorch. Ejecución completa: {n(r['minutos_totales'], 0)} minutos.")
doc.add_heading("4.2 Línea base: TF-IDF + clasificador lineal", 2)
busq = r["busqueda_C"]
p("TF-IDF de palabras y bigramas (min_df 3, max_df 0,9, sublinear_tf, hasta 300 mil columnas) dentro de un Pipeline de scikit-learn, con "
  "regresión logística o SVM lineal y class_weight balanceado. El valor de C se eligió con F1 macro de validación:")
tabla([["Modelo", "C", "F1 validación", "Entrenamiento (s)"]] +
      [[corto(b["modelo"]), n(b["C"], 2), n(b["f1_val"]), n(b["seg_entreno"], 0)] for b in busq], [6, 2.5, 3.5, 4],
      titulo="Tabla 4. Búsqueda de C en validación")
doc.add_heading("4.3 Modelos secuenciales (sesión 3)", 2)
p("TextVectorization con vocabulario de 30 mil tokens aprendido solo con entrenamiento (0 = PAD, 1 = OOV), secuencias de 100 tokens, "
  "embedding entrenable de 64 dimensiones con mask_zero, una capa recurrente de 64 unidades many-to-one (SimpleRNN, LSTM o GRU), dropout "
  "0,3 y softmax. Adam (1e-3), entropía cruzada con pesos de clase, lotes de 256 y early stopping con paciencia 2 sobre la pérdida de "
  f"validación. OOV en validación: {pct(A['oov_val'], 2)}.")
doc.add_heading("4.4 Transformer preentrenado: BETO (sesión 4)", 2)
p("El Transformer reemplaza la recurrencia por atención de producto punto escalado, softmax(QKᵀ/√dₖ + M)·V: cada token combina "
  "información de todos los demás en una sola capa, en paralelo, con embeddings posicionales para conservar el orden. BETO "
  "(dccuchile/bert-base-spanish-wwm-cased) es un BERT encoder-only de 12 capas, 12 cabezas y 768 dimensiones (110 M de parámetros), "
  "preentrenado con Masked Language Modeling sobre un corpus grande en español. Para clasificar se usa la representación del token [CLS] "
  "seguida de una capa lineal.")
p("Se eligió fine-tuning completo (no extracción de características ni PEFT) porque el modelo es de tamaño base y el dominio es específico "
  "(jerga de compras, errores ortográficos, ironía). El notebook incluye además la auditoría del tokenizador frente a XLM-RoBERTa, el "
  "cálculo manual de atención, las máscaras de relleno y causal, la codificación posicional y la alineación BIO de subpalabras para la "
  "futura extensión a sentimiento por aspecto.")
hist = [h for h in s4["historia_bert"] if "eval_f1_macro" in h]
tabla([["Hiperparámetro", "Valor"],
       ["Checkpoint", "dccuchile/bert-base-spanish-wwm-cased (licencia CC BY 4.0)"],
       ["Optimizador", "AdamW, learning rate 2e-5, weight decay 0,01"],
       ["Warmup y scheduler", "10% de los pasos, decaimiento lineal; gradient clipping 1,0"],
       ["Batch y épocas", f"16; {len(hist)} épocas, se conserva la mejor por F1 macro de validación"],
       ["Longitud", f"max_length {f['max_len']} (elegido por percentiles), padding dinámico y agrupación por longitud"],
       ["Pérdida", "Entropía cruzada con pesos de clase (igual que los demás modelos)"],
       ["Datos", f"{entero(f['n_train_bert'])} comentarios estratificados de entrenamiento; evaluación por época con 1.500 de validación"],
       ["Tiempo", f"{n(f['seg_bert'] / 60, 0)} minutos en CPU"],
       ["Variante de eficiencia", "Cuantización dinámica int8 de las capas lineales (solo inferencia)"]], [4.5, 12],
      titulo="Tabla 5. Configuración del fine-tuning de BETO")
p(f"Restricción de cómputo: sin GPU no es viable ajustar BETO con 340 mil comentarios. Para separar el efecto del preentrenamiento del "
  f"efecto de tener más datos se entrenó también una TF-IDF + LR con exactamente los mismos {entero(f['n_train_bert'])} comentarios.")
doc.add_heading("4.5 Calibración, abstención y criterio de adopción", 2)
p("Las probabilidades se calibran con temperature scaling ajustado en la validación común. La política de abstención envía a revisión "
  "humana los comentarios cuya confianza calibrada es menor que un umbral; el umbral es el menor que logra recall de negativo ≥ 0,92 en los "
  "casos que el modelo decide solo, elegido en validación y reportado una sola vez en prueba. El criterio de la sesión 4 exige "
  "simultáneamente: mejora de F1 macro ≥ 0,02 frente al campeón anterior, recall de negativo ≥ 0,92, cobertura automática ≥ 75% y "
  "latencia p95 ≤ 150 ms por comentario.")
doc.add_heading("4.6 Arquitectura de la solución y MLOps", 2)
p("La solución se piensa como un servicio que recibe comentarios de varios canales y devuelve una decisión trazable. El flujo de datos "
  "va del canal a una zona protegida donde se enmascara la PII; la inteligencia combina el modelo recomendado, la calibración y la regla "
  "de abstención; el producto expone una API para la cola de atención y un tablero para Gerencia.")
tabla([["Capa", "Componente", "Detalle"],
       ["Ingesta", "Lotes horarios desde reseñas, chats y reclamos", "Validación de esquema, idioma y longitud"],
       ["Protección", "Enmascaramiento de PII", "<EMAIL>, <TEL>, <URL>; acceso por roles; retención de 90 días"],
       ["Inferencia", f"{corto(REC)} + temperatura + umbral", "Etiqueta, probabilidades, decisión y evidencia (src/demo.py)"],
       ["Respaldo", "TF-IDF + Regresión Logística", "Explicable, < 1 ms; plan de reversión inmediato"],
       ["Revisión humana", "Cola de casos con baja confianza", "El agente confirma antes de alertar al vendedor"],
       ["Monitoreo", "Tablero semanal", "Abstención, confianza, clases, subpalabras/OOV, latencia, F1 en muestra reetiquetada"],
       ["Reentrenamiento", "Trimestral o por disparador", "Abstención +10 pp, recall de negativo < 0,88 o caída de F1 por país"]],
      [3, 5.5, 8], titulo="Tabla 6. Arquitectura y operación propuestas")
p("Se versionan juntos, para poder reproducir y revertir cualquier predicción: checkpoint base y tokenizador, código de tokenización, "
  "max_length, pesos ajustados, hiperparámetros, semilla, partición (índices), conjunto de datos y su hash, temperatura de calibración, "
  "umbral de abstención, métricas de validación y entorno de ejecución (requirements.txt). Cada versión candidata pasa por la suite de "
  "contrafactuales y por la evaluación por segmentos antes de salir de modo sombra.")

# ------------------------------------------------------------------ 5. Resultados
doc.add_heading("5. Resultados", 1)
doc.add_heading("5.1 Prueba completa (modelos entrenados con 340 mil comentarios)", 2)
filas = [["Modelo", "Ajuste", "F1 val.", "F1 prueba", "F1 neutro", "Recall neg.", "Entreno (min)", "p95 (ms)"]]
for k, v in sorted(res3.items(), key=lambda kv: -kv[1]["val"]["f1_macro"]):
    filas.append([corto(k), v["hiperparametro"], n(v["val"]["f1_macro"]), n(v["test"]["f1_macro"]), n(v["test"]["f1_neutro"]),
                  n(v["test"]["recall_negativo"]), n(v["seg_entreno"] / 60, 1), n(v["p95_ms_1"], 1)])
tabla(filas, [3.2, 2.8, 1.5, 1.6, 1.6, 1.7, 1.9, 1.5], size=8.5, titulo=f"Tabla 6. Resultados de la sesión 3 sobre la prueba completa ({entero(A['n_test'])} comentarios)")
p(f"El mejor recurrente según validación es {S3}; la diferencia con la línea base en la prueba completa es "
  f"{'+' if f['f1_full'][S3] - f['f1_full'][BASE] >= 0 else ''}{n(f['f1_full'][S3] - f['f1_full'][BASE])} de F1 macro, a cambio de un "
  f"entrenamiento varias veces más largo.")
# curvas de aprendizaje (se generan aquí a partir del historial guardado por el notebook)
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(1, 2, figsize=(10, 3.4))
    for (k, h), col in zip(r["historias_rnn"].items(), ["#7DD3FC", "#1F6FB2", "#0B2E59"]):
        ep = range(1, len(h["loss"]) + 1)
        ax[0].plot(ep, h["loss"], "--", color=col, label=f"{k} train")
        ax[0].plot(ep, h["val_loss"], "-o", color=col, ms=4, label=f"{k} val.")
        ax[1].plot(ep, h["val_accuracy"], "-o", color=col, ms=4, label=k)
    if hist:
        ax[1].plot([h["epoch"] for h in hist], [h["eval_accuracy"] for h in hist], "-s", color="#0FA3B1", label="BETO (val.)")
    ax[0].set_title("Pérdida: entrenamiento vs. validación"); ax[0].set_xlabel("época"); ax[0].legend(fontsize=7, ncol=2)
    ax[1].set_title("Accuracy en validación"); ax[1].set_xlabel("época"); ax[1].legend(fontsize=7)
    for a in ax:
        a.spines[["top", "right"]].set_visible(False)
    plt.tight_layout(); plt.savefig(os.path.join(FIG, "curvas.png"), dpi=150); plt.close()
    figura("curvas.png", 15.5, "Figura 2. Curvas de aprendizaje: la pérdida de validación sube tras 1–2 épocas (sobreajuste), "
           "por eso el early stopping restaura la mejor época.")
except Exception as e:  # noqa: BLE001
    print("sin curvas:", e)
doc.add_heading("5.2 Comparación justa sobre la prueba común (incluye BETO)", 2)
orden = sorted(comun, key=lambda k: -comun[k]["val"]["f1_macro"])
filas = [["Modelo", "Datos entr.", "F1 val.", "F1 prueba", "F1 neg.", "F1 neu.", "F1 pos.", "Recall neg.", "Log loss", "ECE"]]
for k in orden:
    v = comun[k]
    filas.append([corto(k), entero(s4["info"][k]["n_train"]), n(v["val"]["f1_macro"]), n(v["test"]["f1_macro"]),
                  n(v["test"]["f1_negativo"]), n(v["test"]["f1_neutro"]), n(v["test"]["f1_positivo"]),
                  n(v["test"]["recall_negativo"]), n(v["test"]["log_loss"]), n(v["ece_test"])])
tabla(filas, [3.2, 1.7, 1.3, 1.4, 1.3, 1.3, 1.3, 1.5, 1.4, 1.2], size=8, titulo=f"Tabla 7. Ocho modelos sobre la misma prueba común ({entero(A['n_comun'])} comentarios)")
p(f"Campeón por F1 macro en validación común: {corto(s4['campeon'])}. {corto(BERT)} logra {n(f['f1'][BERT])} en prueba. Con el mismo "
  f"presupuesto de datos ({entero(f['n_train_bert'])} comentarios), la TF-IDF obtiene {n(f['f1'][SUB])}: la diferencia "
  f"({'+' if f['f1'][BERT] - f['f1'][SUB] >= 0 else ''}{n(f['f1'][BERT] - f['f1'][SUB])}) es el aporte del preentrenamiento. "
  f"La versión con max_length 64 en inferencia obtiene {n(s4['sens_longitud']['f1_64'])} frente a {n(s4['sens_longitud']['f1_128'])} "
  f"con 128, lo que confirma la elección de longitud por percentiles.")
tabla([["Diferencia de F1 macro (prueba común)", "Media", "IC 95%", "Lectura"]] +
      [[k.replace("TF-IDF + Regresión Logística", "TF-IDF + LR"), n(m), f"[{n(a)}; {n(b)}]",
        "significativa" if a > 0 or b < 0 else "no concluyente"] for k, (m, a, b) in f["ic"].items()],
      [7, 2, 4, 3.5], titulo="Tabla 8. Intervalos de confianza por bootstrap (1.000 remuestreos)")
doc.add_heading("5.3 Errores y segmentos", 2)
figura("confusion.png", 16, "Figura 2. Matrices de confusión normalizadas por fila en la prueba común.")
seg = s4["segmentos"]
filas = [["Segmento", corto(BASE), S3, "BETO"]]
for col, nom in [("tramo", "Longitud"), ("negacion", "Negación"), ("contraste", "Contraste"), ("country", "País")]:
    for g, v in seg[col].items():
        etiqueta = {"True": "con", "False": "sin"}.get(g, g)
        filas.append([f"{nom}: {etiqueta}", n(v[BASE]), n(v[S3]), n(v["BETO"])])
tabla(filas, [6, 3.5, 3.5, 3.5], size=8.5, titulo="Tabla 9. F1 macro por segmento (prueba común)")
causas = s4["errores_causas"]
tot = sum(causas.values()) or 1
p(f"Se revisaron todos los errores de BETO en la prueba común ({entero(tot)}) y se etiquetó su causa probable con reglas: "
  + "; ".join(f"{c} {pct(v / tot, 0)}" for c, v in causas.items()) + ". "
  f"La mayoría está en la frontera con el neutro: comentarios de tres estrellas que suenan positivos o negativos, cuya etiqueta proviene "
  f"de las estrellas y no de una lectura humana. Los errores graves (negativo leído como positivo) fueron "
  f"{s4['neg_como_pos'].get('BETO')} con BETO, {s4['neg_como_pos'].get(S3)} con {S3} y {s4['neg_como_pos'].get(BASE)} con la línea base.")
mu = s4["errores_muestra"][:12]
tabla([["Comentario (recortado)", "★", "Real", "BETO", "Conf.", "Causa probable"]] +
      [[e["texto"][:95] + ("…" if len(e["texto"]) > 95 else ""), e["review_rate"], e["sentimiento"], e["BETO"], n(e["conf_BETO"], 2), e["causa"]]
       for e in mu], [7.2, 0.7, 1.6, 1.6, 1.2, 4.2], size=7.5, titulo="Tabla 10. Muestra de errores representativos de BETO (20 en el notebook)")
doc.add_heading("5.4 Robustez y contrafactuales", 2)
casos = s4["casos_robustez"]
mods = [BASE, S3, "BETO"]
tabla([["Caso", "Texto", "Esperado"] + [corto(k) for k in mods]] +
      [[c["caso"], c["texto"][:60] + ("…" if len(c["texto"]) > 60 else ""), c["esperado"]] + [c[k] for k in mods] for c in casos],
      [2.2, 5.6, 1.8, 2.4, 2.4, 2.4], size=7.5, titulo="Tabla 11. Pruebas contrafactuales (→ revisión = se abstiene con su umbral)")
p("Casos superados de " + str(len(casos)) + ": " + ", ".join(f"{corto(k)} {s4['robustez'][k]}" for k in mods) +
  ". Las entradas fuera de dominio deben terminar en abstención; las entradas vacías se rechazan antes de llegar al modelo.")
if SES:
    doc.add_heading("5.5 Hallazgo: sesgo de longitud del corpus", 2)
    p(f"Las pruebas contrafactuales revelaron que elogios cortos se clasifican como negativos. La causa está en los datos: el "
      f"{pct(SES['pct_corpus_le15'])} del corpus son comentarios de 15 palabras o menos y, de ellos, el {pct(SES['pct_negativo_le15'])} "
      f"es negativo; los positivos casi siempre son largos. Los modelos aprendieron el atajo “corto ⇒ negativo”, un sesgo de muestreo del "
      f"corpus que no ocurriría con la misma intensidad en chats o reclamos reales.")
    tabla([["Palabras", "Negativo", "Neutro", "Positivo", "Comentarios"]] +
          [[k, pct(v["negativo"]), pct(v["neutro"]), pct(v["positivo"]), entero(v["comentarios"])] for k, v in SES["distribucion"].items()],
          [3, 3, 3, 3, 3.5], size=8.5, titulo="Tabla 13. Distribución de la etiqueta según la longitud del comentario")
    tabla([["Modelo", "Segmento", "n", "% real neg.", "% predicho neg.", "Recall positivo", "F1 macro"]] +
          [[corto(m["modelo"]), m["segmento"], m["n"], pct(m["% real negativo"]), pct(m["% predicho negativo"]),
            pct(m["recall positivo"]), n(m["F1 macro"])] for m in SES["modelos"]],
          [3.2, 2.6, 1.2, 2, 2.4, 2.3, 1.8], size=8.5, titulo="Tabla 14. Desempeño por longitud (prueba común)")
    cf = SES["contrafactual"]
    p("Contrafactual de longitud (mismo elogio, corto y con más contexto): " + "; ".join(
        f"“{c['texto'][:40]}…” ({c['palabras']} palabras) → BETO {c['BETO']}" for c in cf) +
      ". Control adoptado: los comentarios de 15 palabras o menos se priorizan pero no disparan alertas automáticas al vendedor, y el "
      "siguiente reentrenamiento debe incluir comentarios cortos balanceados (por ejemplo, de chats etiquetados por agentes).")

doc.add_heading("5.6 Calibración y abstención", 2)
cal = s4["calibracion"]
tabla([["Modelo", "Temperatura", "ECE antes", "ECE después", "Umbral", "Cobertura", "Recall neg. auto", "Revisión por 1.000"]] +
      [[corto(k), n(cal[k]["T"], 2), n(cal[k]["ece_antes"]), n(cal[k]["ece_despues"]), n(f["umbral"].get(k), 2),
        pct(ab.get(k, {}).get("cobertura")), n(ab.get(k, {}).get("recall_neg_auto")),
        entero(1000 * ab[k]["abstencion"]) if ab.get(k, {}).get("abstencion") is not None else "—"] for k in cal],
      [3.4, 1.8, 1.6, 1.8, 1.4, 1.8, 2.1, 2.3], size=8, titulo="Tabla 12. Calibración y abstención (umbral elegido en validación, reportado en prueba)")
figura("abstencion.png", 15.5, "Figura 3. Calidad frente a cobertura y recall de negativo según el umbral (prueba común, probabilidades calibradas).")
doc.add_heading("5.7 Costo, latencia y eficiencia", 2)
ef = s4["eficiencia"]
tabla([["Modelo", "F1 prueba", "p50 ms", "p95 ms", "Coment./s", "Parámetros (M)", "Tamaño (MB)", "US$/millón", "Dominado"]] +
      [[corto(e["modelo"]), n(e["F1 test común"]), n(e["p50 ms"], 1), n(e["p95 ms"], 1), entero(e["coment./s"]),
        n(e["parámetros (M)"], 2), "—" if e["tamaño (MB)"] is None else n(e["tamaño (MB)"], 1), n(e["USD por millón"], 2),
        "sí" if e["dominado"] else "no"] for e in ef], [3.2, 1.5, 1.3, 1.3, 1.6, 1.9, 1.8, 1.7, 1.6], size=8,
      titulo="Tabla 13. Perfil de eficiencia en CPU (un comentario a la vez; lotes de 32 para el throughput)")
p("Supuesto de costo: una instancia de 4 vCPU a US$0,20 por hora procesando al throughput medido en este equipo. Un modelo está dominado "
  "si otro tiene mejor F1 y menor latencia. La cuantización int8 reduce la latencia de BETO con una pérdida de F1 pequeña.")
figura("frontera.png", 12, "Figura 4. Frontera calidad-costo (escala logarítmica de latencia).")
doc.add_heading("5.8 Interpretabilidad", 2)
figura("interpretabilidad.png", 16, "Figura 5. Atención de la última capa de BETO y oclusión de palabras para un comentario con contraste.")
top = r["top_palabras"]
p("La regresión logística es explicable por coeficientes; los términos con más peso son: " +
  "; ".join(f"{c}: {', '.join(w for w, _ in top[c][:8])}" for c in ["negativo", "neutro", "positivo"]) +
  ". En BETO la atención muestra qué tokens mezcla cada posición, pero no es una explicación causal; por eso se complementa con oclusión "
  "y con la evidencia de la regresión logística en la demo.")
doc.add_heading("5.9 Criterio de adopción", 2)
crit = s4["criterio"]
tabla([["Modelo", "Criterio", "Umbral", "Valor", "Cumple"]] +
      [[corto(c["modelo"]), c["criterio"], c["umbral"].replace("Regresión Logística", "LR"),
        "—" if c["valor"] is None or c["valor"] != c["valor"] else n(c["valor"], 3), "sí" if c["cumple"] else "no"] for c in crit],
      [3, 4.5, 4.5, 2.3, 1.7], size=8.5, titulo="Tabla 14. Criterio de adopción de la sesión 4")

doc.add_heading("5.10 Validez y limitaciones de la evidencia", 2)
viñetas([("Validez interna: ", "la partición es la misma para todos los modelos; ningún hiperparámetro, umbral o temperatura se eligió "
          "con la prueba; los duplicados se eliminaron antes de partir; la prueba común se usó una sola vez. Se reporta un intervalo de "
          "confianza por bootstrap para las diferencias de F1."),
         ("Una sola semilla: ", "cada modelo se entrenó una vez con semilla 42. Las diferencias pequeñas entre LSTM y GRU, o entre BETO y "
          "BETO int8, deben leerse con cautela; el siguiente experimento debe repetir con tres semillas."),
         ("Presupuesto desigual: ", f"BETO usó {entero(f['n_train_bert'])} comentarios y las redes recurrentes {entero(A['n_train'])}. "
          "La TF-IDF con el mismo subconjunto aísla el efecto de los datos, pero BETO con la base completa probablemente mejore más."),
         ("Validez externa: ", "las reseñas no son chats ni reclamos; el lenguaje de posventa es más corto, urgente y con más errores. Antes "
          "de usarlo en esos canales hay que medir el desempeño con una muestra etiquetada de cada canal."),
         ("Etiqueta débil: ", "las estrellas miden satisfacción global, no necesariamente el tono del texto; el techo de desempeño del neutro "
          "está limitado por esa ambigüedad."),
         ("Sesgo de longitud: ", "el 94% de los comentarios cortos del corpus son negativos; las métricas globales ocultan que los "
          "positivos cortos casi no se reconocen (sección 5.5)."),
         ("Hardware: ", f"las latencias se midieron en un portátil sin GPU ({E['nucleos_logicos']} hilos). En un servidor o con GPU serán "
          "menores; el criterio de latencia debe volver a medirse en el entorno de despliegue.")])

# ------------------------------------------------------------------ 6. Riesgos y gobierno
doc.add_heading("6. Riesgos y gobierno", 1)
p("La matriz completa (13 riesgos con probabilidad, impacto, control, responsable, evidencia, indicador y riesgo residual) está en "
  "docs/Matriz_Riesgos.xlsx. Los controles principales son:")
viñetas([("Privacidad y minimización: ", "se procesa solo el texto del comentario; correos, teléfonos y URL se enmascaran antes de inferir y "
          "de registrar; acceso por roles; retención de logs de 90 días; ningún dato se envía a servicios externos."),
         ("Sesgo: ", "F1 por país, categoría y longitud en cada reentrenamiento; Perú y Venezuela tienen pocos datos y requieren un mínimo de "
          "desempeño antes de activar alertas en esos sitios."),
         ("Seguridad y trazabilidad: ", "versionado conjunto de checkpoint, tokenizer, datos, partición, semilla, max_length, temperatura y "
          "umbral; requirements.txt con versiones; sin credenciales en el código (.env.example)."),
         ("Revisión humana: ", "el modelo solo prioriza y alerta; ninguna sanción es automática; los casos de baja confianza van a un agente."),
         ("Monitoreo y drift: ", "tasa de abstención, distribución de confianza, subpalabras por palabra/OOV, distribución de clases, latencia "
          "y F1 sobre una muestra reetiquetada mensual. Reentrenamiento trimestral o cuando la abstención suba 10 puntos."),
         ("Dueño, rollback y suspensión: ", "dueño del modelo: líder de Machine Learning; dueño del proceso: Gerencia de IA. Se suspende y se "
          "vuelve a TF-IDF + LR si el recall de negativo cae por debajo de 0,88 dos semanas, si la abstención supera 35%, si un país cae por "
          "debajo de F1 0,65, ante cualquier exposición de PII o si la precisión de alertas baja de 80%.")])

# ------------------------------------------------------------------ 7. Caso de valor
doc.add_heading("7. Caso de valor", 1)
volumen = 1_000_000
cob = ab_rec.get("cobertura") or 0
rev = volumen * (1 - cob)
thr_rec = f["thr"].get(REC, 1)
costo_comp = volumen / thr_rec / 3600 * 0.20
p("Supuestos ilustrativos (deben validarse con datos internos): un millón de comentarios al mes en un país mediano; 40% negativos; un "
  "agente revisa 60 comentarios por hora a un costo cargado de US$8 por hora; hoy se lee una muestra del 10%; cada negativo atendido a "
  "tiempo evita una devolución en el 5% de los casos, con un costo promedio de US$12 por devolución.")
tabla([["Concepto", "Sin modelo (muestreo 10%)", f"Con {corto(REC)} + abstención"],
       ["Comentarios leídos por personas", entero(volumen * 0.10), entero(rev)],
       ["Horas de agente al mes", entero(volumen * 0.10 / 60), entero(rev / 60)],
       ["Costo de revisión (US$)", entero(volumen * 0.10 / 60 * 8), entero(rev / 60 * 8)],
       ["Negativos detectados", entero(volumen * 0.4 * 0.10),
        entero(volumen * 0.4 * ((ab_rec.get('recall_neg_efectivo') or 0)))],
       ["Devoluciones evitadas (5% de negativos atendidos)", entero(volumen * 0.4 * 0.10 * 0.05),
        entero(volumen * 0.4 * (ab_rec.get('recall_neg_efectivo') or 0) * 0.05)],
       ["Costo de cómputo (US$)", "0", n(costo_comp, 2)]], [6.5, 5, 5], titulo="Tabla 15. Estimación mensual (supuestos ilustrativos)")
p(f"Sensibilidad: el beneficio depende sobre todo de la tasa de devoluciones evitadas y de la cobertura automática. Con la mitad del efecto "
  f"supuesto (2,5%) el modelo sigue detectando muchos más negativos que el muestreo; si la cobertura cayera por debajo de 60%, la carga de "
  f"revisión humana crecería y habría que subir el umbral o ampliar el equipo. El costo de cómputo es marginal frente al costo humano, "
  f"incluso para BETO en CPU (US$ {n(s4['info']['BETO']['thr'] and 1e6 / s4['info']['BETO']['thr'] / 3600 * 0.20, 2)} por millón).")

# ------------------------------------------------------------------ 8. Recomendación
doc.add_heading("8. Recomendación", 1)
p(decision, negrita_inicio="Decisión: ")
if D:
    viñetas([("Por qué: ", D["por_que"]), ("Criterio estricto: ", D["criterio_estricto"])] +
            [("Condición: ", c) for c in D["condiciones"]])
viñetas(([] if D else [("Por qué: ", f"{corto(BERT)} es el mejor modelo en F1 macro sobre la misma prueba "
          f"({'diferencia significativa' if signif else 'diferencia no concluyente'} frente a {S3}), aprovecha el preentrenamiento con "
          f"muchos menos datos y, cuantizado, cabe en el presupuesto de latencia." if f["cumple_bert"] else
          f"El Transformer no cumple simultáneamente los criterios de calidad, clase crítica, cobertura y latencia en este hardware."),
         ("Condiciones: ", "modo sombra de cuatro semanas; recall de negativo ≥ 0,92 y cobertura ≥ 75% medidos sobre comentarios reales "
          "revisados por agentes; precisión de alertas ≥ 80%.")]) + [
         ("Roadmap: ", "(1) preparación y revisión legal, 2 semanas; (2) modo sombra, 4 semanas; (3) humano en el ciclo, 6 semanas; "
          "(4) alertas automáticas al vendedor; (5) mejora continua trimestral."),
         ("Siguiente experimento: ", "fine-tuning de BETO con GPU y la base completa, comparación con XLM-RoBERTa y RoBERTuito, "
          "reetiquetado humano de una muestra del neutro y sentimiento por aspecto (envío, calidad, precio, atención)."),
         ("Límites: ", "la etiqueta proviene de estrellas; los datos son reseñas, no chats ni reclamos; no hay fecha para medir drift "
          "temporal; BETO se ajustó con un subconjunto por restricción de cómputo.")])

# ------------------------------------------------------------------ Anexos
salto()
doc.add_heading("Anexos", 1)
doc.add_heading("A. Diccionario de datos", 2)
tabla([["Columna", "Tipo", "Uso"],
       ["id", "entero", "Identificador; no se modela"],
       ["country", "texto (7)", "Solo evaluación por segmento"],
       ["category", "texto (5)", "Solo evaluación por segmento"],
       ["review_content", "texto libre", "Única entrada de los modelos"],
       ["review_title", "texto corto", "Excluida por fuga"],
       ["review_rate", "entero 1–5", "Excluida por fuga; fuente de la etiqueta"],
       ["sentimiento", "categórica", "Variable objetivo"]], [4, 3.5, 9])
doc.add_heading("B. Curvas de entrenamiento", 2)
tabla([["Modelo", "Época", "Pérdida train", "Pérdida val.", "Accuracy val."]] +
      [[k, i + 1, n(h["loss"][i]), n(h["val_loss"][i]), n(h["val_accuracy"][i])]
       for k, h in r["historias_rnn"].items() for i in range(len(h["loss"]))] +
      [["BETO", n(h["epoch"], 0), "—", n(h["eval_loss"]), n(h["eval_accuracy"])] for h in hist], [3, 2, 3.5, 3.5, 3.5], size=8.5)
doc.add_heading("C. Calibración", 2)
figura("calibracion.png", 13, "Figura A1. Diagrama de confiabilidad de BETO antes y después de temperature scaling.")
doc.add_heading("D. Reproducibilidad y artefactos", 2)
viñetas(["Repositorio: github.com/pedrogarciar13/proyecto_final_pln (código, notebook ejecutado, muestra de datos, documentos).",
         "Instalación: pip install -r requirements.txt; base completa en data/ (ver data/README.md); ejecutar el notebook de inicio a fin.",
         "Artefactos generados en modelos/ (no versionados por tamaño): tfidf_lr.joblib, rnn/lstm/gru.keras, beto_sentimiento/ (~440 MB), "
         "config_demo.json.",
         "Demostración: python src/demo.py \"texto\" (etiqueta, probabilidades calibradas, decisión, alerta y evidencia).",
         "Documentos: src/generar_presentacion.py, src/generar_informe.py y src/generar_matriz_riesgos.py se regeneran desde resultados/resultados.json."])
doc.add_heading("E. Declaración de uso de IA generativa", 2)
p("Se usó un asistente de IA (Claude, de Anthropic) para estructurar el código del notebook, redactar borradores de documentación y "
  "generar los scripts de la presentación, el informe y la matriz de riesgos. Todas las cifras de este informe provienen de la ejecución "
  "del notebook (resultados/resultados.json) y fueron verificadas por el equipo; no se presenta como evidencia ningún resultado que no se "
  "haya ejecutado. No se enviaron datos confidenciales a servicios externos: el corpus es público. El equipo mantiene la responsabilidad y "
  "la autoría de las decisiones de diseño, la interpretación y la recomendación.")
doc.add_heading("F. Referencias", 2)
for ref in ["Vaswani, A. et al. (2017). Attention Is All You Need. NeurIPS.",
            "Devlin, J. et al. (2019). BERT: Pre-training of Deep Bidirectional Transformers for Language Understanding. NAACL.",
            "Cañete, J. et al. (2020). Spanish Pre-Trained BERT Model and Evaluation Data (BETO). PML4DC, ICLR.",
            "Guo, C. et al. (2017). On Calibration of Modern Neural Networks. ICML.",
            "Hochreiter, S. y Schmidhuber, J. (1997). Long Short-Term Memory. Neural Computation.",
            "Cho, K. et al. (2014). Learning Phrase Representations using RNN Encoder-Decoder (GRU). EMNLP.",
            "LPSC-FIUBA. MeLiSA: Mercado Libre for Sentiment Analysis. Hugging Face, lpsc-fiuba/melisa.",
            "Moreno Morales, G. J. (2026). Procesamiento de Lenguaje Natural, sesiones 1 a 4. Universidad Ean."]:
    par = doc.add_paragraph(ref, style="List Number")
    par.runs[0].font.size = Pt(10)

# numeración de páginas en el pie
for s in doc.sections:
    par = s.footer.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = par.add_run("Informe técnico · Sentimiento en reseñas de Mercado Libre · página ")
    run.font.size = Pt(8)
    for tipo, txt in [("begin", None), (None, "PAGE"), ("end", None)]:
        r_ = par.add_run()
        r_.font.size = Pt(8)
        if tipo:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tipo)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
        r_._r.append(el)

os.makedirs(os.path.join(RAIZ, "docs"), exist_ok=True)
salida = os.path.join(RAIZ, "docs", "Informe_Tecnico_Proyecto_PLN.docx")
doc.save(salida)
print("Informe ->", salida)
