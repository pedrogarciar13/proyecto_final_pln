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
from informe_texto import EQUIPO, POR_ROL, ML, DS, GIA, SUPUESTOS, caso_valor  # noqa: E402

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
D = f["decision"] or {}
SES = f["sesgo"] or {}
OOD = s4.get("ood") or {}
RM = s4.get("revision_manual") or []
AGG = s4.get("agregado") or {}
VAR = s4.get("variabilidad") or []
RR = s4.get("resumen_robustez") or []
rec_corto = {m["modelo"]: m["recall positivo"] for m in SES.get("modelos", []) if m["segmento"].startswith("≤")}
ood_tab = {t["método"]: t for t in OOD.get("tabla", [])}
ood_b = ood_tab.get("similitud BETO (k=5)", {})
S = SUPUESTOS
usd_millon_rec = 1e6 / s4["info"][REC]["thr"] / 3600 * 0.20
actual, modelo = caso_valor(ab_rec.get("cobertura") or 0, ab_rec.get("recall_neg_efectivo") or 0, usd_millon_rec)
impl = sum(S["implementacion"].values())
incremental = modelo["neto"] - actual["neto"]
mejor_rnn = max(["RNN", "LSTM", "GRU"], key=lambda k: f["f1"][k])

# ------------------------------------------------------------------ portada
for _ in range(4):
    doc.add_paragraph()
p("UNIVERSIDAD EAN · PROCESAMIENTO DE LENGUAJE NATURAL", size=11).runs[-1].font.color.rgb = TEAL
t = p("¿Qué siente el comprador? Análisis de sentimiento en reseñas de Mercado Libre", size=24)
t.runs[-1].bold = True
t.runs[-1].font.color.rgb = NAVY
p("Comparamos una línea base TF-IDF, redes recurrentes y un Transformer preentrenado en español (BETO), y proponemos cómo llevarlo "
  "a operación con controles", size=13)
doc.add_paragraph()
p("Informe técnico del proyecto final", size=12).runs[-1].bold = True
for nombre, rol in EQUIPO:
    p(f"{nombre} · {rol}", size=11)
p("Docente: Gerson José Moreno Morales", size=11)
p("Bogotá, noviembre de 2026", size=11)
salto()

# ------------------------------------------------------------------ 1. Resumen
doc.add_heading("1. Resumen", 1)
p(f"Mercado Libre recibe muchos más comentarios de los que una persona puede leer, y solo las reseñas traen estrellas: los chats, las "
  f"preguntas y los reclamos llegan únicamente con texto. Nos propusimos construir un clasificador de sentimiento (negativo, neutro o "
  f"positivo) que lea el comentario y ayude a decidir qué atender primero. Trabajamos con {entero(A['filas'])} reseñas públicas del "
  f"corpus MeLiSA; después de auditarlas quedaron {entero(A['filas_limpias'])}, y dejamos por fuera el título y las estrellas porque "
  f"le regalan la respuesta al modelo.")
p(f"Comparamos ocho modelos sobre la misma partición. En una prueba común de {entero(A['n_comun'])} comentarios, BETO, ajustado con "
  f"solo {entero(f['n_train_bert'])} ejemplos, llegó a un F1 macro de {n(f['f1']['BETO'])}, frente a {n(f['f1'][S3])} de la mejor red "
  f"recurrente elegida en validación ({S3}) y {n(f['f1'][BASE])} de la línea base. La ventaja sobre {S3} es de {n(ic_bs3[0])} puntos, "
  f"con un intervalo de confianza del 95% entre {n(ic_bs3[1])} y {n(ic_bs3[2])}. En el camino encontramos un sesgo serio en los datos: "
  f"el {pct(SES.get('pct_negativo_le15'), 0)} de los comentarios de 15 palabras o menos son negativos, y los modelos aprendieron a "
  f"confundir \"corto\" con \"negativo\".")
p("Recomendamos un piloto de BETO en modo sombra, procesando por lotes, con abstención, un control para textos que no son reseñas y "
  "sin alertas automáticas para comentarios cortos. TF-IDF con regresión logística queda como respaldo explicable.")

# ------------------------------------------------------------------ 2. Contexto
doc.add_heading("2. Contexto y problema", 1)
doc.add_heading("2.1 La organización y el usuario", 2)
p("La reputación de un vendedor y de un producto mueve el ranking de búsqueda y la conversión en Mercado Libre. Un comprador molesto "
  "al que nadie atiende termina, muchas veces, en una devolución, un reclamo formal o, simplemente, en un cliente que no vuelve. El "
  "problema no es la falta de información: es que llega en forma de texto libre y en un volumen imposible de leer.")
p("Pensamos el sistema para una persona concreta: el analista de experiencia del comprador. Su decisión de todos los días es qué "
  "comentario mirar primero y a qué vendedor avisarle a tiempo. El sistema no reemplaza esa decisión; ordena la cola y le dice cuándo "
  "no está seguro.")
doc.add_heading("2.2 Situación actual", 2)
p(f"No tenemos cifras internas de Mercado Libre, así que trabajamos con supuestos explícitos, razonables para un país mediano, y los "
  f"rotulamos como tales en todo el documento. Suponemos {entero(S['volumen_mes'])} comentarios al mes, de los que un "
  f"{pct(S['pct_negativos'], 0)} son negativos. Hoy se lee una muestra del {pct(S['muestreo_actual'], 0)}: unos "
  f"{entero(actual['leidos'])} comentarios, que a {S['comentarios_hora_agente']} por hora y US${n(S['costo_hora_agente'], 0)} la hora "
  f"cuestan cerca de US${entero(actual['revision'])} al mes. Con ese muestreo, nueve de cada diez quejas nunca se leen, y las que se "
  f"leen tardan en promedio unas {S['respuesta_actual_h']} horas en recibir respuesta. Esa es nuestra línea base.")
tabla([["Elemento", "Definición"],
       ["Usuario", "Analista de experiencia del comprador (Atención al Cliente y Reputación de vendedores)"],
       ["Decisión que apoya", "Qué comentario se atiende primero y a qué vendedor se le envía una alerta temprana"],
       ["Entrada y unidad", "Un comentario de texto libre en español; la unidad es el comentario"],
       ["Frecuencia y modo", "Flujo continuo procesado en lotes cada hora (no se requiere respuesta en tiempo real)"],
       ["Salida", "Etiqueta, probabilidad calibrada, decisión automática o revisión humana, y las palabras que más pesaron"],
       ["Clase crítica", "Negativo. El error más caro es leer como positivo a un cliente molesto"],
       ["KPI y línea base", f"Primera respuesta a negativos: hoy ~{S['respuesta_actual_h']} h (supuesto), meta < 24 h. Alertas confirmadas por un agente: meta ≥ 80%"]],
      [4, 12.5], titulo="Tabla 1. Formulación del problema")
doc.add_heading("2.3 Hipótesis de valor, alcance y criterios fijados de antemano", 2)
p("Nuestra hipótesis es sencilla: si el sistema encuentra la mayoría de los negativos y los pone al principio de la cola, el tiempo de "
  "primera respuesta baja y se evitan devoluciones. Medimos el éxito técnico con el F1 macro, que da el mismo peso a las tres clases, "
  "y lo acompañamos con el recall de negativo, la cobertura automática bajo abstención, la calibración y la latencia.")
p("Antes de entrenar dejamos fijos los criterios de adopción de la sesión 4 (mejora de F1 ≥ 0,02 frente al campeón anterior, recall de "
  "negativo ≥ 0,92, cobertura ≥ 75% y latencia p95 ≤ 150 ms). Como el proceso es por lotes, agregamos un criterio operativo: el sistema "
  "debe procesar el volumen de una hora (unos 1.400 comentarios) en menos de 10 minutos con un solo proceso. Quedan por fuera del "
  "alcance las sanciones automáticas, el sentimiento por aspecto (lo proponemos como siguiente fase), el portugués y otros canales "
  "sin una evaluación nueva.")

# ------------------------------------------------------------------ 3. Datos
doc.add_heading("3. Datos", 1)
doc.add_heading("3.1 Origen, periodo y licencia", 2)
p("MeLiSA (Mercado Libre for Sentiment Analysis) es un corpus público del Laboratorio de Procesamiento de Señales y Comunicaciones de "
  "la Universidad de Buenos Aires, publicado en Hugging Face (lpsc-fiuba/melisa) con licencia \"other\". Lo usamos con fines "
  "académicos. Son reseñas reales, recolectadas entre agosto de 2020 y enero de 2021 en siete países de habla hispana; cada una trae "
  "el comentario, el título, las estrellas, el país y la categoría. La tarjeta del corpus dice algo que después resultó clave: los "
  "autores priorizaron reseñas según su longitud y su valoración (likes menos dislikes), y equilibraron la cantidad de cada estrella. "
  "La data card completa, con diccionario de datos, población, limitaciones y usos no previstos, está en el repositorio "
  "(data/README.md) y resumida en el Anexo A.")
tabla([["Estrellas", "Sentimiento", "Proporción", "Qué hace el negocio"],
       ["1 – 2", "negativo", "40%", "Alerta temprana y cola prioritaria"],
       ["3", "neutro", "20%", "Oportunidad de mejora"],
       ["4 – 5", "positivo", "40%", "Testimonio y reputación"]], [3, 3, 3, 7.5], titulo="Tabla 2. Esquema de etiquetas")
p("La etiqueta no la puso una persona leyendo el texto: sale de las estrellas. Es una supervisión débil, barata y abundante, pero "
  "ruidosa, sobre todo en el neutro.")
doc.add_heading("3.2 Auditoría", 2)
pii = A.get("pii", {})
p(f"No hay valores nulos. Encontramos {entero(A['duplicados_texto'])} comentarios repetidos (dejamos una sola copia para que el mismo "
  f"texto no quede en entrenamiento y en prueba) y {entero(A['textos_contradictorios'])} textos idénticos con sentimientos distintos, "
  f"que eliminamos porque le enseñarían dos respuestas a la misma pregunta. La mediana es de {n(A['palabras_p50'], 0)} palabras y el "
  f"99% tiene {n(A['palabras_p99'], 0)} o menos; en subpalabras de BETO, el 95% cabe en {n(A['tokens_bert']['p95'], 0)}, así que una "
  f"longitud máxima de 128 trunca menos del 2%.")
p(f"Buscamos datos personales con expresiones regulares: no aparecieron correos, sí {pii.get('URL', 0)} enlaces y {pii.get('TEL', 0)} "
  f"secuencias con forma de teléfono. Revisándolas, muchas no son teléfonos sino años o referencias de producto (\"2017-2018\", "
  f"\"776622-001\"). Igual las enmascaramos todas: preferimos perder una referencia de producto a dejar pasar un número personal. El "
  f"corpus tiene un solo canal (reseñas), así que los cortes por segmento se hacen por país, categoría y longitud.")
p("Hay otra particularidad que solo vimos al probar la demo: en las 487 mil reseñas la palabra \"llegó\" no aparece ni una vez, y tampoco \"envío\", \"entrega\" ni \"servicio\"; \"vendedor\" aparece dos veces. Las reseñas de producto de Mercado Libre hablan del producto, no del despacho ni del vendedor, y todo indica que ese contenido se filtró en origen. Para nuestro caso de uso es un cambio de dominio importante: buena parte de los reclamos en chats son justamente sobre envíos, y ese vocabulario el modelo no lo vio al entrenar. Ni siquiera el control de fuera de dominio los separa, porque en significado se parecen a una queja de compra. La solución de fondo es reentrenar con conversaciones de posventa; mientras tanto, el piloto debe limitarse a reseñas.")
tabla([["Revisión", "Hallazgo", "Decisión"],
       ["Nulos", "0 en todas las columnas", "—"],
       ["Duplicados de texto", entero(A["duplicados_texto"]), "Una sola copia"],
       ["Textos contradictorios", entero(A["textos_contradictorios"]), "Se eliminan"],
       ["Longitud en palabras", f"P50 {n(A['palabras_p50'], 0)} · P95 {n(A['palabras_p95'], 0)} · P99 {n(A['palabras_p99'], 0)}", "RNN: 100 tokens"],
       ["Longitud en subpalabras", f"P95 {n(A['tokens_bert']['p95'], 0)} · P99 {n(A['tokens_bert']['p99'], 0)}", "BETO: 128 (trunca 1,8%)"],
       ["Posible PII", f"0 correos · {pii.get('URL', 0)} URL · {pii.get('TEL', 0)} \"teléfonos\" (muchos son años o referencias)", "Se enmascara todo"],
       ["Países", "MLA 43% · MLM 31% · MCO 10% · MLC 9% · MLU 5% · MLV 1% · MPE 0,7%", "Evaluación por país"],
       ["Comentarios ≤ 15 palabras", f"{pct(SES.get('pct_corpus_le15'))} del corpus; {pct(SES.get('pct_negativo_le15'))} de ellos negativos", "Ver sección 5.4"],
       ["Filas finales", entero(A["filas_limpias"]), "—"]], [4, 7, 5.5], titulo="Tabla 3. Auditoría del corpus")
figura("eda.png", 15, "Figura 1. Longitud de los comentarios por clase y distribución de clases por país.")
doc.add_heading("3.3 Fuga de información", 2)
p(f"Hicimos una prueba simple con el mismo TF-IDF y SVM: leyendo solo el título, el modelo obtiene F1 macro {n(A['f1_solo_titulo'])}; "
  f"leyendo el comentario completo, {n(A['f1_solo_comentario'])}. Que tres palabras le ganen al comentario significa que el título se "
  f"escribe junto con las estrellas. En un chat no existe, así que lo excluimos, igual que las estrellas. También cuidamos las fugas "
  f"menos visibles: deduplicamos antes de partir, el vocabulario, el IDF y el tokenizador se ajustan solo con entrenamiento, y el "
  f"umbral, la temperatura y el modelo campeón se eligen en validación. Queda una fuga pequeña que no podemos quitar: hay compradores "
  f"que escriben su calificación en el texto (\"le doy tres estrellas\").")
doc.add_heading("3.4 Partición", 2)
p(f"Partimos al azar y de forma estratificada, con semilla 42: {entero(A['n_train'])} comentarios para entrenar (70%), "
  f"{entero(A['n_val'])} para validar (15%) y {entero(A['n_test'])} para la prueba final (15%). Una partición temporal habría sido "
  f"mejor, pero el archivo no trae fecha por reseña ni identificador de comprador; las reseñas son independientes entre sí. Como BETO en "
  f"CPU no alcanzaba a evaluar 73 mil comentarios en un tiempo razonable, sacamos con la misma semilla una validación y una prueba comunes "
  f"de {entero(A['n_comun'])} comentarios, y comparamos los ocho modelos sobre esos mismos textos.")

# ------------------------------------------------------------------ 4. Metodología
doc.add_heading("4. Metodología", 1)
doc.add_heading("4.1 Pipeline y reproducibilidad", 2)
p(f"Todo el flujo vive en un solo notebook que corre de principio a fin, con rutas relativas y parámetros por variables de entorno. Lo "
  f"ejecutamos en un portátil con CPU de {E['nucleos_logicos']} hilos, {E['ram_gb'] / 1.0737:.0f} GB de RAM y sin GPU, con Python "
  f"{E['python']}, scikit-learn {E['scikit-learn']}, TensorFlow {E['tensorflow']}, PyTorch {E['torch']} y Transformers "
  f"{E['transformers']}. La corrida completa tomó {n(r['minutos_totales'], 0)} minutos; la mayor parte se la llevó BETO.")
doc.add_heading("4.2 Línea base: TF-IDF y un clasificador lineal", 2)
busq = r["busqueda_C"]
mejor_c = {b["modelo"]: max([x for x in busq if x["modelo"] == b["modelo"]], key=lambda x: x["f1_val"]) for b in busq}
p(f"La línea base convierte cada comentario en un vector TF-IDF de palabras y pares de palabras (así \"no llegó\" es una sola "
  f"característica), con hasta 300 mil columnas, y lo clasifica con regresión logística o SVM lineal, ambas con pesos de clase "
  f"balanceados porque el neutro es la mitad de frecuente. Probamos tres valores de C por modelo y nos quedamos con el mejor en "
  f"validación: C = {n(mejor_c[BASE]['C'], 1)} para la regresión logística y C = {n(mejor_c['TF-IDF + SVM lineal']['C'], 2)} para la "
  f"SVM (Anexo B).")
doc.add_heading("4.3 Redes recurrentes (sesión 3)", 2)
p(f"Las redes leen el comentario en orden. Usamos un vocabulario de 30 mil palabras aprendido solo con entrenamiento, secuencias de "
  f"100 tokens, un embedding de 64 dimensiones que se aprende con la tarea y una capa recurrente de 64 unidades (SimpleRNN, LSTM o GRU) "
  f"en modo many-to-one: solo el último estado decide. Entrenamos con Adam, pesos de clase y early stopping sobre la pérdida de "
  f"validación; en todas, la pérdida de validación empezó a subir después de una o dos épocas.")
doc.add_heading("4.4 Transformer preentrenado: BETO (sesión 4)", 2)
p("El Transformer cambia la lectura paso a paso por atención: cada token calcula qué tan relevante es cada uno de los demás, "
  "softmax(QKᵀ/√dₖ)·V, y combina su información en una sola capa. BETO es un BERT de 12 capas y 110 millones de parámetros preentrenado "
  "con textos en español. Para clasificar tomamos la representación del token [CLS] y le ponemos una capa lineal encima. En el notebook "
  "también hicimos los ejercicios de la sesión: el cálculo de atención a mano, las máscaras, la codificación posicional, la auditoría "
  "del tokenizador frente a XLM-RoBERTa y la alineación de etiquetas BIO.")
hist = [h for h in s4["historia_bert"] if "eval_f1_macro" in h]
tabla([["Decisión", "Valor y por qué"],
       ["Estrategia", "Fine-tuning completo: el modelo es mediano y el dominio (jerga de compras, errores) lo justifica"],
       ["Optimizador", "AdamW, learning rate 2e-5, weight decay 0,01, warmup del 10% y decaimiento lineal"],
       ["Lote y épocas", f"16 y {len(hist)} épocas; nos quedamos con la mejor por F1 macro de validación"],
       ["Longitud", f"max_length {f['max_len']}, elegido por percentiles, con padding dinámico"],
       ["Pérdida", "Entropía cruzada con pesos de clase, igual que los demás modelos"],
       ["Datos", f"{entero(f['n_train_bert'])} comentarios estratificados (restricción de cómputo)"],
       ["Tiempo", f"{n(f['seg_bert'] / 60, 0)} minutos en CPU"],
       ["Variante", "Cuantización dinámica int8 de las capas lineales, solo para inferencia"]], [3.5, 13],
      titulo="Tabla 4. Configuración del fine-tuning de BETO")
p(f"Sin GPU no era viable ajustar BETO con 340 mil comentarios. Para que la comparación fuera justa entrenamos también una TF-IDF con "
  f"exactamente los mismos {entero(f['n_train_bert'])} comentarios: así separamos lo que aporta el preentrenamiento de lo que aporta "
  f"tener más datos.")
doc.add_heading("4.5 Calibración, abstención y control de fuera de dominio", 2)
p("Las probabilidades de softmax suelen ser demasiado confiadas, así que las calibramos con temperature scaling en la validación común. "
  "Sobre esas probabilidades definimos la abstención: si la confianza queda por debajo de un umbral, el comentario lo revisa una "
  "persona. El umbral es el menor que logra recall de negativo ≥ 0,92 en lo que el modelo decide solo; se elige en validación y se "
  "reporta una vez en prueba.")
p(f"La revisión del proyecto nos mostró un hueco: un texto que no es una reseña (\"La fotosíntesis depende de la energía solar\") podía "
  f"recibir una etiqueta con 0,94 de confianza. La abstención por confianza no lo detecta, porque el modelo está seguro de algo absurdo. "
  f"Agregamos un control previo: comparamos la representación interna de BETO de cada texto con la de 3.000 reseñas de entrenamiento y, "
  f"si la similitud con sus cinco vecinos más cercanos queda por debajo del percentil 2 de las reseñas de validación, el texto va a "
  f"revisión. Con ese umbral el control marca el {pct(ood_b.get('reseñas de prueba marcadas'))} de las reseñas reales y detecta el "
  f"{pct(ood_b.get('textos fuera de dominio detectados'), 0)} de 20 textos ajenos que escribimos (deportes, clima, recetas, código, "
  f"otros idiomas).")
doc.add_heading("4.6 Arquitectura y operación", 2)
p("Pensamos la solución como un servicio que recibe comentarios de varios canales y devuelve una decisión trazable. Los datos pasan por "
  "una zona protegida donde se enmascara la información personal; luego el control de dominio, el modelo calibrado y la regla de "
  "abstención; al final, una cola de atención para los agentes y un tablero para Gerencia. Versionamos juntos el checkpoint, el "
  "tokenizador, la partición, la semilla, la temperatura y el umbral, de modo que cualquier predicción se pueda reproducir y revertir.")
tabla([["Capa", "Qué hace"],
       ["Ingesta", "Lotes horarios; valida esquema, idioma y longitud"],
       ["Protección", "Enmascara correos, teléfonos y enlaces; acceso por roles; logs con retención de 90 días"],
       ["Inferencia", f"Control de dominio → {corto(REC)} calibrado → umbral {n(f['umbral'].get(REC), 2)} → regla de comentarios cortos"],
       ["Respaldo", "TF-IDF + regresión logística: explicable y de un milisegundo; reversión inmediata"],
       ["Revisión humana", "Cola de casos dudosos; el agente confirma antes de alertar a un vendedor"],
       ["Monitoreo", "Abstención, confianza, distribución de clases, subpalabras por palabra, latencia y F1 en una muestra reetiquetada"]],
      [3.5, 13], titulo="Tabla 5. Arquitectura propuesta")

# ------------------------------------------------------------------ 5. Resultados
doc.add_heading("5. Resultados", 1)
doc.add_heading("5.1 Comparación sobre la misma prueba", 2)
orden = sorted(comun, key=lambda k: -comun[k]["val"]["f1_macro"])
filas = [["Modelo", "Datos de entr.", "F1 val.", "F1 prueba", "F1 neutro", "Recall neg.", "ECE", "p95 (ms)"]]
for k in orden:
    v = comun[k]
    filas.append([corto(k), entero(s4["info"][k]["n_train"]), n(v["val"]["f1_macro"]), n(v["test"]["f1_macro"]),
                  n(v["test"]["f1_neutro"]), n(v["test"]["recall_negativo"]), n(v["ece_test"]), n(s4["info"][k]["p95"], 1)])
tabla(filas, [3.6, 2, 1.6, 1.7, 1.7, 1.8, 1.4, 1.6], size=8.5,
      titulo=f"Tabla 6. Los ocho modelos sobre la misma prueba común ({entero(A['n_comun'])} comentarios)")
p(f"BETO es el mejor en validación y en prueba ({n(f['f1']['BETO'])}). La diferencia con {S3} no parece casualidad: remuestreamos la "
  f"prueba mil veces y el intervalo del 95% para la diferencia va de {n(ic_bs3[1])} a {n(ic_bs3[2])}, lejos de cero (Anexo B). "
  f"Elegimos {S3} como campeón recurrente porque fue el mejor en validación; en la prueba, {mejor_rnn} queda ligeramente arriba "
  f"({n(f['f1'][mejor_rnn])}), una diferencia que está dentro de la variación entre semillas. Lo más interesante es la última fila de "
  f"la comparación justa: la TF-IDF entrenada con los mismos {entero(f['n_train_bert'])} comentarios que BETO llega a "
  f"{n(f['f1'][SUB])}. Con el mismo presupuesto de datos, el preentrenamiento vale casi nueve puntos de F1.")
if VAR:
    p("Repetimos las redes recurrentes con dos semillas más para saber cuánto cambian: " +
      "; ".join(f"{v['modelo']} {n(v['media'])} ± {n(v['desv. est.'])}" for v in VAR) +
      ". La regresión logística y la SVM dan siempre el mismo resultado. BETO no lo pudimos repetir por costo y queda pendiente.")
doc.add_heading("5.2 Errores", 2)
figura("confusion.png", 15.5, "Figura 2. Matrices de confusión normalizadas por fila en la prueba común.")
causas = s4["errores_causas"]
tot = sum(causas.values()) or 1
p(f"BETO se equivoca en {entero(tot)} de los {entero(A['n_comun'])} comentarios de prueba. Casi todos los errores involucran al neutro, "
  f"que es la clase más difícil también para una persona. Lo que más le importa al negocio son los negativos leídos como positivos: "
  f"{s4['neg_como_pos']['BETO']} con BETO frente a {s4['neg_como_pos'][BASE]} con TF-IDF. Además, el "
  f"{pct(s4.get('errores_abstencion'), 0)} de los errores de BETO tiene una confianza por debajo del umbral, así que en operación los "
  f"habría revisado una persona.")
p("Una regla automática agrupa los errores por causa aproximada, pero no reemplaza leerlos. Tomamos 20 errores y los revisamos uno "
  "por uno. La Tabla 7 muestra doce; el notebook tiene los veinte.")
if RM:
    tabla([["Comentario (recortado)", "★", "Real → BETO", "Lo que vimos"]] +
          [[e["texto"][:85] + ("…" if len(e["texto"]) > 85 else ""), e["★"], f"{e['real']} → {e['BETO']}", e["comentario"]] for e in RM[:12]],
          [6.2, 0.7, 2.6, 7.0], size=7.5, titulo="Tabla 7. Revisión manual de errores de BETO")
    p("De esta lectura salen tres patrones. Primero, elogios cortos (\"Muy lindo novedoso, la niña se divierte mucho\") que salen "
      "negativos con mucha confianza: es el sesgo de longitud de la sección 5.4. Segundo, opiniones mixtas en las que las estrellas no "
      "coinciden con lo que dice el texto; ahí el error es tanto de la etiqueta como del modelo. Tercero, detalles que exigen conocer el "
      "dominio o leer con cuidado: \"no es sin TACC\" (no es apto para celíacos), \"no endurece jamás\" en un sellador, o \"Nos es muy "
      "fácil\" en lugar de \"No es muy fácil\".")
doc.add_heading("5.3 Robustez y contrafactuales", 2)
rr = {(x["modelo"], x["grupo"]): x for x in RR}
p("Probamos pares mínimos: la misma frase con y sin negación, con el contraste invertido, con errores de ortografía, en mayúsculas, "
  "con emojis, con ironía y textos que no son reseñas. Contamos cada caso en una de tres columnas: acierta, se abstiene (manda a "
  "revisión algo que sí tenía respuesta) o falla.")
if RR:
    tabla([["Modelo", "Grupo", "Casos", "Acierta", "Se abstiene", "Falla"]] +
          [[corto(x["modelo"]), x["grupo"], x["casos"], x["acierta"], x["se abstiene"], x["falla"]] for x in RR],
          [4, 4.5, 1.6, 1.8, 2.2, 1.6], size=8.5, titulo="Tabla 8. Resultado de los contrafactuales (detalle en el Anexo C)")
p("En los textos cortos, BETO y GRU responden \"negativo\" casi siempre, también ante \"El producto es excelente, lo recomiendo\". En una "
  "primera versión contábamos eso como \"7 aciertos de 13\", lo cual era engañoso: los aciertos eran justo los casos negativos. Con "
  "los pares alargados a más de 15 palabras, BETO deja de decir \"negativo\" a todo: el elogio largo sale positivo (con confianza "
  "baja, así que lo revisaría una persona) y la queja larga sale negativa. Pero no lo resuelve todo: \"Pensé que sería malo, pero es "
  "excelente\", alargado, sale neutro. Curiosamente, en esta prueba TF-IDF es el que mejor se porta (11 de 19), porque se apoya en "
  "palabras como \"excelente\" y no en la longitud. Los tres textos ajenos los manda a revisión solo BETO, gracias al control de "
  "dominio.")
doc.add_heading("5.4 El sesgo de longitud", 2)
if SES:
    p(f"El {pct(SES['pct_corpus_le15'])} del corpus son comentarios de 15 palabras o menos, y el {pct(SES['pct_negativo_le15'])} de "
      f"ellos son negativos. La tarjeta de MeLiSA da la explicación: los autores priorizaron reseñas largas y valoradas, y los elogios "
      f"cortos (\"Excelente, lo recomiendo\") quedaron casi fuera. Para los modelos, entonces, \"corto\" era casi sinónimo de \"queja\". "
      f"El resultado se ve en la prueba: con 15 palabras o menos, BETO reconoce solo el {pct(rec_corto.get('BETO'), 0)} de los positivos "
      f"y TF-IDF el {pct(rec_corto.get(BASE), 0)}. El mismo elogio con más contexto se clasifica bien.")
    p("Es la amenaza más seria para usar el modelo en chats y reclamos, que suelen ser cortos. Mientras no reentrenemos con comentarios "
      "cortos balanceados, la regla es que un comentario de 15 palabras o menos se puede priorizar, pero nunca dispara una alerta "
      "automática al vendedor.")
doc.add_heading("5.5 Resumen agregado de opinión", 2)
if AGG:
    p("Para Gerencia, el valor no está solo en cada comentario sino en el agregado: qué porcentaje de quejas tiene cada país o categoría. "
      "Comparamos el porcentaje de negativos real con el que estima BETO en la prueba común.")
    tabla([["País", "Comentarios", "% neg. real", "% neg. BETO", "Diferencia (pp)"]] +
          [[a["country"], entero(a["comentarios"]), pct(a["% negativo real"]), pct(a["% negativo BETO"]),
            n(a["error % negativo (pp)"], 1)] for a in AGG["pais"]], [3, 2.6, 2.6, 2.6, 2.8], size=8.5,
          titulo="Tabla 9. Porcentaje de negativos por país: real y estimado")
    grandes = [a for a in AGG["pais"] if a["comentarios"] >= 1000]
    maxg = max(abs(a["error % negativo (pp)"]) for a in grandes)
    peor = max((a for a in AGG["pais"] if a["comentarios"] >= 100), key=lambda a: abs(a["error % negativo (pp)"]))
    p(f"En Argentina y México, los dos países con más datos, la estimación se aleja del valor real como mucho {n(maxg, 1)} puntos; en "
      f"{peor['country']} la diferencia llega a {n(abs(peor['error % negativo (pp)']), 1)} puntos. Sirve para seguir tendencias, pero no "
      f"para comparar países que se diferencian por pocos puntos, sobre todo los de menos datos.")
doc.add_heading("5.6 Calibración, abstención y costo", 2)
figura("abstencion.png", 15, "Figura 3. Calidad frente a cobertura y recall de negativo según el umbral (prueba común, probabilidades calibradas).")
cal = s4["calibracion"]["BETO"]
p(f"La calibración bajó el error de calibración (ECE) de BETO de {n(cal['ece_antes'])} a {n(cal['ece_despues'])}. Con el umbral de "
  f"{n(f['umbral']['BETO'], 2)}, BETO decide solo el {pct(ab['BETO'].get('cobertura'))} de los comentarios, con recall de negativo "
  f"{n(ab['BETO'].get('recall_neg_auto'))} en lo que decide; unos {entero(1000 * (ab['BETO'].get('abstencion') or 0))} de cada mil van a "
  f"revisión. En costo, TF-IDF responde en un milisegundo y BETO en {n(f['p95']['BETO'], 0)} ms en el percentil 95 (en CPU); la versión "
  f"int8 baja a {n(f['p95']['BETO int8'], 0)} ms y pierde dos puntos de F1. Con una máquina de 4 vCPU a US$0,20 la hora, un millón de "
  f"comentarios cuesta centavos con TF-IDF y unos US${n(f['thr'] and 1e6 / s4['info']['BETO']['thr'] / 3600 * 0.20, 2)} con BETO "
  f"(Anexo B).")
doc.add_heading("5.7 Criterio de adopción", 2)
crit = s4["criterio"]
mods_c = list(dict.fromkeys(c["modelo"] for c in crit))
criterios = list(dict.fromkeys(c["criterio"] for c in crit))
filas = [["Criterio"] + [corto(m) for m in mods_c]]
for cr in criterios:
    fila = [cr]
    for m in mods_c:
        c = next((x for x in crit if x["modelo"] == m and x["criterio"] == cr), None)
        if c is None:
            fila.append("—")
            continue
        v = c["valor"]
        txt = "—" if v is None or v != v else (n(v, 2 if v < 1 else 1) + (" ms" if "Latencia" in cr else " min" if "lote" in cr else "") if ("Latencia" in cr or "lote" in cr)
                                               else pct(v) if "Cobertura" in cr else n(v))
        fila.append(("sí · " if c["cumple"] else "no · ") + txt)
    filas.append(fila)
tabla(filas, [5] + [11.5 / len(mods_c)] * len(mods_c), size=8.5, titulo="Tabla 10. Criterio de adopción (BETO frente a GRU; GRU frente a la línea base)")
p("Con el criterio estricto del curso ningún modelo cumple todo. BETO cumple calidad, clase crítica, cobertura y capacidad en lote, y "
  "falla la latencia en línea por 0,1 ms en un portátil sin GPU. La versión int8 cumple la latencia, pero su mejora sobre GRU se queda "
  "en 0,014. GRU no supera a la línea base por 0,02 ni llega al 75% de cobertura. Como el proceso se definió desde el principio por "
  "lotes, el criterio que importa para operar es la capacidad en lote, y ahí BETO cumple con margen.")
doc.add_heading("5.8 Validez y límites", 2)
p("El corpus tampoco habla de envíos ni de vendedores (sección 3.2), así que el modelo no está listo para reclamos de despacho sin reentrenar con ese tipo de texto.")
p("Cuidamos la validez interna: misma partición para todos, ninguna decisión tomada con la prueba, duplicados fuera antes de partir e "
  "intervalos de confianza para las diferencias. Los límites también son claros. BETO usó 16 mil comentarios y no lo repetimos con "
  "otras semillas. La etiqueta viene de las estrellas, que miden satisfacción y no necesariamente el tono del texto. Las reseñas no son "
  "chats: antes de usar el modelo en otro canal hay que medirlo con una muestra etiquetada de ese canal. Y las latencias se midieron en "
  "un portátil; en el entorno de despliegue hay que volver a medirlas.")

# ------------------------------------------------------------------ 6. Riesgos y gobierno
doc.add_heading("6. Riesgos y gobierno", 1)
p("La matriz completa tiene catorce riesgos con su causa, probabilidad, impacto, control, responsable, evidencia, indicador y riesgo "
  "residual (docs/Matriz_Riesgos.xlsx). Aquí contamos los que más nos preocupan y cómo los vamos a manejar.")
doc.add_heading("6.1 Privacidad", 2)
p("El sistema solo necesita el texto del comentario. Antes de inferir y antes de escribir cualquier log se enmascaran correos, "
  "teléfonos y enlaces; el acceso a los datos es por roles y los logs se borran a los 90 días. Ningún comentario sale a un servicio "
  "externo: el modelo corre dentro de la infraestructura. Una exposición de datos personales, por pequeña que sea, es causa de "
  "suspensión inmediata.")
doc.add_heading("6.2 Sesgo y equidad", 2)
p(f"Encontramos dos sesgos que vale la pena separar. El primero es el de longitud, que ya describimos y que tiene un control explícito. "
  f"El segundo es por país: Perú y Venezuela tienen menos del 2% de los datos. En la prueba común el F1 de BETO por país va de "
  f"{n(min(v['BETO'] for v in s4['segmentos']['country'].values()))} a {n(max(v['BETO'] for v in s4['segmentos']['country'].values()))}, "
  f"así que no vemos un país claramente perjudicado, pero con tan pocos datos en algunos sitios la conclusión es frágil. Por eso pedimos "
  f"un F1 mínimo de 0,65 por país antes de activar alertas en ese sitio.")
doc.add_heading("6.3 Revisión humana y uso previsto", 2)
p("El modelo prioriza y alerta; no decide sobre personas. Ninguna sanción a un vendedor o a un comprador puede salir de una predicción "
  "sin que un agente lo confirme. Los casos con baja confianza, los textos fuera de dominio y las alertas sobre comentarios cortos pasan "
  "siempre por una persona.")
doc.add_heading("6.4 Monitoreo, drift y reentrenamiento", 2)
p("Cada semana miramos la tasa de abstención, la distribución de la confianza y de las clases, las subpalabras por palabra (si sube, "
  "está llegando vocabulario nuevo), la latencia y la proporción de textos marcados como fuera de dominio. Cada mes, un agente "
  "reetiqueta una muestra de 300 comentarios para medir el F1 real por país y por longitud. Reentrenamos cada trimestre, o antes si la "
  "abstención sube diez puntos o cambia la mezcla de canales.")
doc.add_heading("6.5 Responsables, reversión y suspensión", 2)
tabla([["Responsable", "Rol", "Qué le toca"],
       [POR_ROL[ML], ML, "Modelo, calibración, latencia, reentrenamiento y reversión técnica"],
       [POR_ROL[DS], DS, "Calidad de datos, reetiquetado, sesgos, segmentos y análisis de errores"],
       [POR_ROL[GIA], GIA, "Uso previsto, privacidad, KPI, capacidad del equipo humano y decisión de suspender"]],
      [4, 4.6, 7.9], titulo="Tabla 11. Responsables del piloto")
p("Suspendemos y volvemos a TF-IDF con regresión logística si el recall de negativo baja de 0,88 dos semanas seguidas, si la "
  "abstención supera el 35% una semana, si un país queda por debajo de F1 0,65 en la muestra reetiquetada, si la precisión de las "
  "alertas baja del 80% o ante cualquier exposición de datos personales. La reversión es inmediata porque el respaldo siempre está "
  "corriendo en paralelo.")

# ------------------------------------------------------------------ 7. Caso de valor
doc.add_heading("7. Caso de valor", 1)
p("No tenemos cifras internas de Mercado Libre, así que construimos el caso con los supuestos de la sección 2.2 y los marcamos como "
  f"ilustrativos. Además suponemos que una devolución cuesta US${n(S['costo_devolucion'], 0)} y que el "
  f"{pct(S['tasa_evitada'], 0)} de los negativos atendidos a tiempo evita una. Con el modelo, la cobertura y el recall efectivo son los "
  f"medidos en la prueba ({pct(ab_rec.get('cobertura'))} y {pct(ab_rec.get('recall_neg_efectivo'))}); los negativos que van a revisión "
  f"también se atienden, porque los lee una persona.")
tabla([["Concepto (mensual)", "Muestreo actual", f"{corto(REC)} con abstención"],
       ["Comentarios leídos por personas", entero(actual["leidos"]), entero(modelo["leidos"])],
       ["Costo de revisión humana", f"US${entero(actual['revision'])}", f"US${entero(modelo['revision'])}"],
       ["Negativos atendidos", entero(actual["detectados"]), entero(modelo["detectados"])],
       ["Devoluciones evitadas", entero(actual["evitadas"]), entero(modelo["evitadas"])],
       ["Ahorro por devoluciones evitadas", f"US${entero(actual['beneficio'])}", f"US${entero(modelo['beneficio'])}"],
       ["Cómputo y monitoreo", "US$0", f"US${entero(modelo['computo'])}"],
       ["Beneficio neto", f"US${entero(actual['neto'])}", f"US${entero(modelo['neto'])}"]], [6.5, 4.5, 5.5],
      titulo="Tabla 12. Beneficio mensual estimado (supuestos ilustrativos)")
p(f"El beneficio neto incremental frente al muestreo actual es de unos US${entero(incremental)} al mes. Implementarlo tiene costos que "
  f"se pagan una sola vez: " + ", ".join(f"{k.lower()} (US${entero(v)})" for k, v in S["implementacion"].items()) +
  f", unos US${entero(impl)} en total. Con estos supuestos, la inversión se recupera en menos de un mes; aunque los supuestos estuvieran "
  f"muy inflados, el margen es amplio.")
filas = [["Devoluciones evitadas por negativo atendido"] + [f"cobertura {pct(c, 0)}" for c in (0.60, 0.70, ab_rec.get("cobertura") or 0.78)]]
for tasa in (0.01, 0.025, 0.05):
    fila = [pct(tasa, 1)]
    for c in (0.60, 0.70, ab_rec.get("cobertura") or 0.78):
        a0, m0 = caso_valor(c, ab_rec.get("recall_neg_efectivo") or 0, usd_millon_rec, tasa=tasa)
        fila.append(f"US${entero(m0['neto'] - a0['neto'])}")
    filas.append(fila)
tabla(filas, [6, 3.5, 3.5, 3.5], size=9, titulo="Tabla 13. Sensibilidad del beneficio neto incremental mensual")
p("La sensibilidad muestra que el resultado depende mucho más de cuántas devoluciones se evitan que de la cobertura. Con una tasa del "
  "1%, cinco veces menor que la supuesta, el piloto sigue dejando un beneficio positivo. Lo que sí sube con menos cobertura es la carga "
  "del equipo humano: con 60% de cobertura habría que revisar casi el doble de comentarios que con la cobertura medida. El piloto en modo "
  "sombra sirve justamente para medir esa tasa con datos reales antes de comprometer más recursos.")

# ------------------------------------------------------------------ 8. Recomendación
doc.add_heading("8. Recomendación", 1)
p("Recomendamos un piloto de BETO en modo sombra durante cuatro semanas, procesando por lotes horarios, con abstención (umbral "
  f"{n(f['umbral'].get(REC), 2)} sobre probabilidades calibradas), con el control de fuera de dominio y sin alertas automáticas para "
  "comentarios de 15 palabras o menos. TF-IDF con regresión logística corre en paralelo como respaldo y como explicación de cada "
  "decisión.")
p("Lo recomendamos porque BETO es el único modelo que mejora de forma significativa al campeón recurrente con el mismo test, y lo hace "
  "con veinte veces menos datos de entrenamiento. Lo recomendamos con condiciones porque no cumple el criterio estricto de latencia en "
  "línea en este hardware y porque el sesgo de longitud del corpus todavía no está corregido. Si el piloto necesitara responder en "
  "tiempo real, usaríamos la versión int8 o una GPU y mediríamos de nuevo.")
p("El piloto pasa a la fase de humano en el ciclo si durante las cuatro semanas el recall de negativo se mantiene en 0,92 o más, la "
  "cobertura en 75% o más y los agentes confirman al menos el 80% de las alertas. El siguiente experimento es ajustar BETO con GPU y "
  "la base completa, agregar comentarios cortos balanceados, compararlo con XLM-R y RoBERTuito, y empezar a trabajar el sentimiento por "
  "aspecto (envío, calidad, precio, atención) para que la alerta diga qué falló.")

# ------------------------------------------------------------------ Anexos
salto()
doc.add_heading("Anexos", 1)
doc.add_heading("A. Data card resumida", 2)
tabla([["Campo", "Detalle"],
       ["Fuente", "MeLiSA, LPSC-FIUBA (Universidad de Buenos Aires), Hugging Face lpsc-fiuba/melisa, licencia \"other\""],
       ["Periodo", "Agosto de 2020 a enero de 2021"],
       ["Población", "Compradores de 7 países hispanohablantes que escribieron una reseña; 5 categorías de producto"],
       ["Tipo de dato", "Real y público; PII enmascarada en la muestra del repositorio"],
       ["Selección", "Los autores priorizaron reseñas largas y valoradas y equilibraron las estrellas (origen del sesgo de longitud)"],
       ["Etiqueta", "Derivada de las estrellas (1–2 negativo, 3 neutro, 4–5 positivo); sin anotación humana"],
       ["Columnas usadas", "review_content como entrada; country y category solo para evaluar; review_title y review_rate excluidas por fuga"],
       ["Limitaciones", "Sin fecha por reseña; un solo canal; pocos datos de Perú y Venezuela; etiqueta ruidosa en el neutro"],
       ["Usos no previstos", "Sanciones automáticas a personas, inferir atributos personales, otros idiomas o dominios sin reevaluar"]],
      [3.5, 13], size=8.5)
doc.add_heading("B. Experimentos y parámetros", 2)
tabla([["Modelo", "C", "F1 validación", "Entrenamiento (s)"]] +
      [[corto(b["modelo"]), n(b["C"], 2), n(b["f1_val"]), n(b["seg_entreno"], 0)] for b in busq], [6, 2.5, 3.5, 4], size=8.5,
      titulo="Tabla B1. Búsqueda de C en validación")
filas = [["Modelo", "Ajuste", "F1 val.", "F1 prueba", "F1 neutro", "Recall neg.", "Entreno (min)"]]
for k, v in sorted(res3.items(), key=lambda kv: -kv[1]["val"]["f1_macro"]):
    filas.append([corto(k), v["hiperparametro"], n(v["val"]["f1_macro"]), n(v["test"]["f1_macro"]), n(v["test"]["f1_neutro"]),
                  n(v["test"]["recall_negativo"]), n(v["seg_entreno"] / 60, 1)])
tabla(filas, [3.2, 3.2, 1.7, 1.8, 1.8, 2, 2.2], size=8,
      titulo=f"Tabla B2. Modelos de la sesión 3 sobre la prueba completa ({entero(A['n_test'])} comentarios)")
tabla([["Diferencia de F1 macro", "Media", "IC 95%"]] +
      [[k.replace("TF-IDF + Regresión Logística", "TF-IDF + LR"), n(m), f"[{n(a)}; {n(b)}]"] for k, (m, a, b) in f["ic"].items()],
      [8, 3, 5], size=8.5, titulo="Tabla B3. Intervalos de confianza por bootstrap (1.000 remuestreos)")
ef = s4["eficiencia"]
tabla([["Modelo", "F1 prueba", "p50 ms", "p95 ms", "Coment./s", "Parám. (M)", "MB", "US$/millón"]] +
      [[corto(e["modelo"]), n(e["F1 test común"]), n(e["p50 ms"], 1), n(e["p95 ms"], 1), entero(e["coment./s"]), n(e["parámetros (M)"], 2),
        "—" if e["tamaño (MB)"] is None or e["tamaño (MB)"] != e["tamaño (MB)"] else n(e["tamaño (MB)"], 1), n(e["USD por millón"], 2)] for e in ef],
      [3.4, 1.7, 1.5, 1.5, 1.8, 1.9, 1.5, 2], size=8, titulo="Tabla B4. Perfil de eficiencia en CPU")
seg = s4["segmentos"]
filas = [["Segmento", corto(BASE), S3, "BETO"]]
for col, nom in [("tramo", "Longitud"), ("negacion", "Negación"), ("contraste", "Contraste"), ("country", "País")]:
    for g, v in seg[col].items():
        filas.append([f"{nom}: {({'True': 'con', 'False': 'sin'}).get(g, g)}", n(v[BASE]), n(v[S3]), n(v["BETO"])])
tabla(filas, [6, 3.5, 3.5, 3.5], size=8, titulo="Tabla B5. F1 macro por segmento (prueba común)")
doc.add_heading("C. Contrafactuales", 2)
casos = s4["casos_robustez"]
mods = [BASE, S3, "BETO"]
tabla([["Caso", "Texto", "Esperado"] + [corto(k) for k in mods]] +
      [[c["caso"], c["texto"][:58] + ("…" if len(c["texto"]) > 58 else ""), c["esperado"]] + [c[k] for k in mods] for c in casos],
      [2.4, 5.4, 1.8, 2.3, 2.3, 2.3], size=7)
doc.add_heading("D. Calibración y curvas de entrenamiento", 2)
figura("calibracion.png", 12, "Figura D1. Diagrama de confiabilidad de BETO antes y después de temperature scaling.")
figura("curvas.png", 14, "Figura D2. Curvas de aprendizaje de las redes y de BETO.")
figura("interpretabilidad.png", 15, "Figura D3. Atención de la última capa y oclusión para \"No llegó completo, pero el vendedor respondió "
       "rápido y lo solucionó\": BETO lo clasifica como negativo con 0,97, aunque la frase termina bien. Es un buen ejemplo de los límites "
       "del modelo con contrastes.")
doc.add_heading("E. Reproducibilidad", 2)
p("Repositorio: https://github.com/pedrogarciar13/proyecto_final_pln. Para reproducir: instalar requirements.txt, poner la base completa "
  "en data/ (instrucciones en data/README.md) y ejecutar el notebook de principio a fin. La demo corre con python src/demo.py \"texto\" "
  "en la terminal o con python app/app.py en el navegador (con BETO). También está publicada en línea de dos formas: con BETO en https://huggingface.co/spaces/pedrogarciar13/sentimiento-melisa-beto, donde el modelo comprimido a int8 corre en el navegador del visitante (coincide con el original en el 98% de las etiquetas de una muestra de 1.000 reseñas), y con el modelo de respaldo TF-IDF en https://sentimiento-melisa-demo.onrender.com. Los documentos se regeneran desde resultados/resultados.json con los scripts "
  "de src/.")
doc.add_heading("F. Uso de herramientas de IA generativa", 2)
p("Como pide el curso, declaramos el uso de IA generativa. La única herramienta de IA generativa que usamos fue Claude Code (Anthropic), "
  "como asistente principal de programación y redacción; GitHub, Render y Hugging Face los usamos solo como plataformas para el "
  "código, las demos y el modelo. Escribió y depuró buena parte del código del notebook, los scripts que generan los documentos y las demos; propuso "
  "análisis como el control de fuera de dominio y el caso de valor con sensibilidad, y redactó los borradores de este informe. El "
  "equipo definió el problema, revisó el código y los resultados, validó las conclusiones, leyó cada uno de los errores de la Tabla 7, revisó las referencias y tomó la decisión final. Ninguna cifra se "
  "escribió a mano: todas salen de ejecutar el notebook y los documentos se generan desde resultados/resultados.json. Las "
  "interpretaciones que no coincidieron con lo ejecutado se corrigieron. El caso de valor usa supuestos ilustrativos, no datos medidos. "
  "No usamos datos confidenciales: el corpus es público y enmascaramos los datos personales. El detalle está en DECLARACION_IA.md, en "
  "el repositorio.")
doc.add_heading("G. Referencias", 2)
for ref in ["Vaswani, A. et al. (2017). Attention Is All You Need. NeurIPS.",
            "Devlin, J. et al. (2019). BERT: Pre-training of Deep Bidirectional Transformers for Language Understanding. NAACL.",
            "Cañete, J. et al. (2020). Spanish Pre-Trained BERT Model and Evaluation Data (BETO). PML4DC, ICLR.",
            "Guo, C. et al. (2017). On Calibration of Modern Neural Networks. ICML.",
            "Hochreiter, S. y Schmidhuber, J. (1997). Long Short-Term Memory. Neural Computation.",
            "Cho, K. et al. (2014). Learning Phrase Representations using RNN Encoder-Decoder. EMNLP.",
            "LPSC-FIUBA (2021). MeLiSA: Mercado Libre for Sentiment Analysis. Hugging Face, lpsc-fiuba/melisa.",
            "Moreno Morales, G. J. (2026). Procesamiento de Lenguaje Natural, sesiones 1 a 4. Universidad Ean."]:
    par = doc.add_paragraph(ref, style="List Number")
    par.runs[0].font.size = Pt(10)

for s_ in doc.sections:
    par = s_.footer.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = par.add_run("Sentimiento en reseñas de Mercado Libre · página ")
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
import csv  # noqa: E402
with open(os.path.join(RAIZ, "resultados", "caso_valor.csv"), "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(["concepto", "muestreo_actual", "modelo"])
    for k in actual:
        w.writerow([k, round(actual[k], 2), round(modelo[k], 2)])
    w.writerow(["implementacion_unica", "", round(impl, 2)])
    w.writerow(["neto_incremental_mensual", "", round(incremental, 2)])
salida = os.path.join(RAIZ, "docs", "Informe_Tecnico_Proyecto_PLN.docx")
doc.save(salida)
print("Informe ->", salida)
