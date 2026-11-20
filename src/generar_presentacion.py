"""Genera docs/Presentacion_Proyecto_PLN.pptx (18 diapositivas) a partir de resultados/resultados.json.

Uso: python src/generar_presentacion.py [raiz_del_proyecto]
"""
import os
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(__file__))
from resumen import cargar, cifras, corto, entero, n, pct  # noqa: E402
from informe_texto import EQUIPO, POR_ROL, ML as R_ML, DS as R_DS, GIA as R_GIA, caso_valor  # noqa: E402

RAIZ = sys.argv[1] if len(sys.argv) > 1 else "."
r = cargar(RAIZ)
f = cifras(r)
s4 = r["s4"]
A = f["audit"]
BASE, S3, SUB, BERT, REC = f["base"], f["s3"], f["sub"], f["bert"], f["rec"]
FIG = os.path.join(RAIZ, "resultados", "figuras")

NAVY, TEAL, BLUE, ORANGE, GREEN, RED = "0B2E59", "0FA3B1", "1F6FB2", "F39C12", "2BB673", "E4572E"
INK, MUTED, LIGHT, LINE, WHITE = "1E293B", "64748B", "F1F5F9", "CBD5E1", "FFFFFF"
FONT_T, FONT_B = "Calibri", "Calibri"
COL_MODELO = {BASE: "94A3B8", "TF-IDF + SVM lineal": "CBD5E1", "RNN": "7DD3FC", "LSTM": BLUE, "GRU": BLUE,
              SUB: "FDBA74", "BETO": TEAL, "BETO int8": NAVY}


def rgb(h):
    return RGBColor.from_string(h)


prs = Presentation()
prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
BLANK = prs.slide_layouts[6]
TOTAL = 18
_num = [0]


def texto(slide, x, y, w, h, contenido, size=13, color=INK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
          font=FONT_B, italic=False, espacio=4):
    """contenido: str o lista de párrafos; cada párrafo es str o lista de (texto, {opciones})."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    parrafos = contenido if isinstance(contenido, list) else [contenido]
    for i, p in enumerate(parrafos):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.alignment = align
        par.space_after = Pt(espacio)
        runs = p if isinstance(p, list) else [(p, {})]
        for t, o in runs:
            run = par.add_run()
            run.text = t
            run.font.name = o.get("font", font)
            run.font.size = Pt(o.get("size", size))
            run.font.bold = o.get("bold", bold)
            run.font.italic = o.get("italic", italic)
            run.font.color.rgb = rgb(o.get("color", color))
    return tb


def caja(slide, x, y, w, h, fill=LIGHT, line=None, forma=MSO_SHAPE.ROUNDED_RECTANGLE, radio=0.08):
    s = slide.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    if forma == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radio
    return s


def nueva(kicker, titulo, notas, oscura=False):
    _num[0] += 1
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = rgb(NAVY if oscura else WHITE)
    if kicker:
        texto(s, 0.5, 0.28, 9, 0.25, kicker.upper(), size=10, bold=True, color=TEAL)
    if titulo:
        texto(s, 0.5, 0.5, 9, 0.6, titulo, size=24, bold=True, color=WHITE if oscura else NAVY, font=FONT_T)
    if not oscura:
        texto(s, 0.5, 5.28, 6, 0.2, "Sentimiento en reseñas de Mercado Libre · Proyecto final PLN · Universidad Ean",
              size=8, color=MUTED)
        texto(s, 8.9, 5.28, 0.6, 0.2, f"{_num[0]} / {TOTAL}", size=8, color=MUTED, align=PP_ALIGN.RIGHT)
    s.notes_slide.notes_text_frame.text = notas
    return s


def tarjeta(slide, x, y, w, h, titulo, cuerpo, acento=TEAL, fill=LIGHT, size=12, num=None):
    caja(slide, x, y, w, h, fill)
    dx = 0.18
    if num is not None:
        c = caja(slide, x + 0.18, y + 0.18, 0.36, 0.36, acento, forma=MSO_SHAPE.OVAL)
        c.text_frame.text = str(num)
        p = c.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.size, p.runs[0].font.bold, p.runs[0].font.color.rgb = Pt(12), True, rgb(WHITE)
        dx = 0.66
    texto(slide, x + dx, y + 0.2, w - dx - 0.18, 0.35, titulo, size=size + 1, bold=True, color=acento if num is None else NAVY)
    texto(slide, x + 0.18, y + 0.62, w - 0.36, h - 0.75, cuerpo, size=size, color=INK)


def dato(slide, x, y, w, valor, etiqueta, color=NAVY, size=34, oscura=False):
    texto(slide, x, y, w, 0.6, valor, size=size, bold=True, color=color, font=FONT_T)
    texto(slide, x, y + 0.66, w, 0.6, etiqueta, size=11, color=LIGHT if oscura else MUTED)


def tabla(slide, x, y, w, filas, anchos, size=10.5, alto_fila=0.27, resaltar=None, colores_celda=None):
    shp = slide.shapes.add_table(len(filas), len(filas[0]), Inches(x), Inches(y), Inches(w), Inches(alto_fila * len(filas)))
    t = shp.table
    for j, a in enumerate(anchos):
        t.columns[j].width = Inches(a)
    for i, fila in enumerate(filas):
        t.rows[i].height = Inches(alto_fila)
        for j, v in enumerate(fila):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.06)
            c.margin_top = c.margin_bottom = Inches(0.02)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            c.fill.fore_color.rgb = rgb(NAVY if i == 0 else ("E0F2F4" if resaltar is not None and i == resaltar else (WHITE if i % 2 else LIGHT)))
            if colores_celda and (i, j) in colores_celda:
                c.fill.fore_color.rgb = rgb(colores_celda[(i, j)])
            tf = c.text_frame
            tf.text = str(v)
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            for run in p.runs:
                run.font.size = Pt(size)
                run.font.name = FONT_B
                run.font.bold = i == 0 or (resaltar is not None and i == resaltar and j == 0)
                run.font.color.rgb = rgb(WHITE if i == 0 else INK)
    return t


def estilo_grafico(ch, leyenda=False, size=10):
    ch.font.size = Pt(size)
    ch.font.name = FONT_B
    ch.font.color.rgb = rgb(INK)
    ch.has_legend = leyenda
    if leyenda:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(size)


def ejes(ch, vmin=None, vmax=None, fmt='0.00', grid=True):
    va = ch.value_axis
    if vmin is not None:
        va.minimum_scale = vmin
    if vmax is not None:
        va.maximum_scale = vmax
    va.has_major_gridlines = grid
    if grid:
        va.major_gridlines.format.line.color.rgb = rgb("E2E8F0")
    va.format.line.fill.background()
    va.tick_labels.number_format = fmt
    va.tick_labels.number_format_is_linked = False
    va.tick_labels.font.color.rgb = rgb(MUTED)
    ca = ch.category_axis
    ca.format.line.color.rgb = rgb(LINE)
    ca.tick_labels.font.color.rgb = rgb(INK)


def imagen(slide, nombre, x, y, w=None, h=None):
    ruta = os.path.join(FIG, nombre)
    if os.path.exists(ruta):
        return slide.shapes.add_picture(ruta, Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
    texto(slide, x, y, w or 3, 0.3, f"[figura {nombre} no disponible]", size=10, color=MUTED)


ab_rec = f["abst"].get(REC) or {}
ab_bert = f["abst"].get(BERT) or {}
ok = "✔"
no = "✘"
D = f["decision"]
if D:
    decision = D["texto"]
elif f["cumple_bert"]:
    decision = (f"Pilotear {REC} en modo sombra con abstención (umbral {n(f['umbral'][REC], 2)}): en la prueba común decide solo el "
                f"{pct(ab_rec.get('cobertura'))} de los comentarios con recall de negativo {n(ab_rec.get('recall_neg_auto'))}; "
                f"el resto pasa a una persona. TF-IDF + LR queda como respaldo explicable.")
else:
    decision = (f"No adoptar todavía el Transformer: no cumple los cuatro criterios de la sesión 4. Se mantiene {REC} con abstención "
                f"(umbral {n(f['umbral'].get(REC), 2)}) en modo sombra y BETO pasa a un piloto con GPU y más datos de entrenamiento.")
decision_corta = D["texto_corto"] if D else decision
SES = f["sesgo"] or {}
ses_mod = {(m["modelo"], m["segmento"]): m for m in SES.get("modelos", [])}

# ------------------------------------------------------------------ 1. Portada
s = nueva(None, None, "Habla: Gerencia de IA\n\n\"Buenos días. Nuestra pregunta es de negocio: ¿puede Mercado Libre saber qué siente un "
          "comprador leyendo solo su comentario, y cuánta complejidad vale la pena para lograrlo? Comparamos tres familias: una línea "
          "base TF-IDF, redes recurrentes de la sesión 3 y un Transformer preentrenado en español, BETO, de la sesión 4. Hablamos "
          "desde tres roles: Gerencia de IA, Ciencia de Datos y Machine Learning.\"", oscura=True)
texto(s, 0.6, 0.55, 8.8, 0.3, "UNIVERSIDAD EAN · PROCESAMIENTO DE LENGUAJE NATURAL · PROYECTO FINAL", size=11, bold=True, color=TEAL)
texto(s, 0.6, 1.2, 8.8, 1.4, "¿Qué siente el comprador? Sentimiento en reseñas de Mercado Libre", size=34, bold=True, color=WHITE)
texto(s, 0.6, 2.75, 8.6, 0.8, "Línea base TF-IDF frente a RNN/LSTM/GRU (sesión 3) y un Transformer preentrenado, BETO (sesión 4), "
      "con abstención, calibración y gobierno", size=15, color="CADCFC")
caja(s, 0.6, 3.85, 8.8, 0.02, TEAL, forma=MSO_SHAPE.RECTANGLE)
texto(s, 0.6, 4.0, 8.8, 0.3, "   ·   ".join(nom for nom, _ in EQUIPO), size=13, bold=True, color=WHITE)
texto(s, 0.6, 3.55, 8.8, 0.3, "Machine Learning · Ciencia de Datos · Gerencia en IA", size=12, color=TEAL)
texto(s, 0.6, 4.4, 8, 0.3, "Docente: Gerson José Moreno Morales · Noviembre de 2026", size=12, color="CADCFC")

# ------------------------------------------------------------------ 2. Resumen ejecutivo
s = nueva("Resumen ejecutivo · " + f"{POR_ROL[R_GIA]} · Magíster en Gerencia en IA", "La decisión en una diapositiva",
          f"Habla: Gerencia de IA\n\n\"Empezamos por la conclusión. Sobre la misma prueba de 5.000 comentarios, {corto(BERT)} obtiene "
          f"F1 macro {n(f['f1'][BERT])}, {('+' if f['delta_bert_s3'] >= 0 else '')}{n(f['delta_bert_s3'])} frente al mejor modelo "
          f"recurrente ({S3}) y {('+' if f['delta_bert_base'] >= 0 else '')}{n(f['delta_bert_base'])} frente a TF-IDF. "
          f"{decision} En las secciones siguientes mostramos de dónde sale cada número.\"")
for i, (v, et, col) in enumerate([
        (n(f["f1"][BERT]), f"F1 macro de {corto(BERT)} en la prueba común ({entero(A['n_comun'])} comentarios)", TEAL),
        (("+" if f["delta_bert_s3"] >= 0 else "") + n(f["delta_bert_s3"]), f"frente al campeón recurrente ({S3}); el mínimo exigido es +0,02", NAVY),
        (pct(ab_rec.get("cobertura")), f"de comentarios que {corto(REC)} decide solo con recall de negativo ≥ 0,92", GREEN)]):
    dato(s, 0.5 + i * 3.05, 1.3, 2.8, v, et, col)
caja(s, 0.5, 2.85, 9, 2.2, LIGHT)
texto(s, 0.75, 3.0, 8.5, 0.35, "Decisión del comité", size=14, bold=True, color=NAVY)
texto(s, 0.75, 3.4, 8.5, 1.6, decision, size=13)

# ------------------------------------------------------------------ 3. Problema
s = nueva("Problema · " + f"{POR_ROL[R_GIA]} · Magíster en Gerencia en IA", "Quién usa el sistema y qué decisión mejora",
          "Habla: Gerencia de IA\n\n\"El usuario es el analista de experiencia del comprador. La decisión: qué comentarios se atienden "
          "primero y a qué vendedor se le envía una alerta. Hoy se lee por muestreo y las estrellas solo existen en reseñas, no en chats ni "
          "reclamos. Construimos la etiqueta con las estrellas: 1 y 2 negativo, 3 neutro, 4 y 5 positivo. El error más caro es leer como "
          "positivo a un cliente molesto: se pierde la alerta. Por eso la clase crítica es negativo.\"")
tarjeta(s, 0.5, 1.3, 2.85, 2.25, "Usuario y decisión", "Analista de experiencia del comprador: decide qué comentario atender primero "
        "y a qué vendedor alertar.", num=1, acento=NAVY)
tarjeta(s, 3.575, 1.3, 2.85, 2.25, "Entrada y salida", "Un comentario en español → etiqueta + probabilidad calibrada + decisión "
        "automática o revisión humana.", num=2, acento=NAVY)
tarjeta(s, 6.65, 1.3, 2.85, 2.25, "KPI", "Primera respuesta a negativos < 24 h · % de alertas confirmadas por un agente · reclamos "
        "evitados.", num=3, acento=NAVY)
for i, (est, cl, col, acc) in enumerate([("1★ 2★", "NEGATIVO · 40%", RED, "Alerta temprana: clase crítica"),
                                          ("3★", "NEUTRO · 20%", "94A3B8", "Oportunidad de mejora"),
                                          ("4★ 5★", "POSITIVO · 40%", GREEN, "Testimonio y reputación")]):
    x = 0.5 + i * 3.075
    caja(s, x, 3.75, 2.85, 1.3, WHITE, line=col)
    texto(s, x + 0.18, 3.85, 2.5, 0.3, est, size=12, color=MUTED)
    texto(s, x + 0.18, 4.15, 2.5, 0.35, cl, size=15, bold=True, color=col)
    texto(s, x + 0.18, 4.55, 2.5, 0.35, acc, size=11)

# ------------------------------------------------------------------ 4. Datos
s = nueva("Datos · " + f"{POR_ROL[R_DS]} · Magíster en Ciencia de Datos", "MeLiSA: 487 mil reseñas reales, auditadas antes de modelar",
          f"Habla: Ciencia de Datos\n\n\"Usamos MeLiSA, una base pública de la Universidad de Buenos Aires: {entero(A['filas'])} reseñas "
          f"de 7 países y 5 categorías. Cada hallazgo terminó en una decisión: {A['duplicados_texto']} textos repetidos se dejan una sola "
          f"vez, {A['textos_contradictorios']} textos con sentimientos distintos se eliminan, y los pocos correos, teléfonos o URL se "
          f"enmascaran. Quedan {entero(A['filas_limpias'])} reseñas. La mediana es de 30 palabras y el 99% tiene 118 o menos.\"")
for i, (v, et) in enumerate([(entero(A["filas"]), "reseñas en español"), ("7", "países · 5 categorías"),
                             (str(A["duplicados_texto"]), "duplicados (1 copia)"), (str(A["textos_contradictorios"]), "textos contradictorios (eliminados)")]):
    dato(s, 0.5 + i * 2.3, 1.25, 2.1, v, et, NAVY if i < 2 else ORANGE, size=28)
cd = CategoryChartData()
paises = sorted(A["paises"].items(), key=lambda kv: -kv[1])
cd.categories = [p for p, _ in paises]
cd.add_series("Reseñas", [v for _, v in paises])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.4), Inches(2.45), Inches(4.6), Inches(2.75), cd)
ch = gf.chart
estilo_grafico(ch)
ch.has_title = True
ch.chart_title.text_frame.text = "Reseñas por país"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
ejes(ch, fmt="#,##0", grid=False)
ch.value_axis.visible = False
pl = ch.plots[0]
pl.gap_width = 60
pl.series[0].format.fill.solid()
pl.series[0].format.fill.fore_color.rgb = rgb(BLUE)
pl.has_data_labels = True
pl.data_labels.number_format = '#,##0'
pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
pl.data_labels.font.size = Pt(9)
pii = A.get("pii", {})
tabla(s, 5.3, 2.55, 4.2, [["Revisión", "Hallazgo"],
                          ["Nulos", "0"],
                          ["PII (correo / URL / tel.*)", f"{pii.get('EMAIL', 0)} / {pii.get('URL', 0)} / {pii.get('TEL', 0)} → enmascarada"],
                          ["Palabras P50 / P95 / P99", f"{n(A['palabras_p50'], 0)} / {n(A['palabras_p95'], 0)} / {n(A['palabras_p99'], 0)}"],
                          ["Subpalabras BETO P95", f"{n(A['tokens_bert']['p95'], 0)} → max_length {f['max_len']}"],
                          ["Neutro", "20% → pesos de clase"],
                          ["Filas finales", entero(A["filas_limpias"])]], [2.0, 2.2], size=10.5)
texto(s, 5.3, 4.55, 4.2, 0.5, "*Muchos “teléfonos” son años o referencias de producto (p. ej. 2017-2018); igual se enmascaran.", size=9, color=MUTED)

# ------------------------------------------------------------------ 5. Fuga y partición
s = nueva("Datos · " + f"{POR_ROL[R_DS]} · Magíster en Ciencia de Datos", "Sin fuga: el título regala la respuesta y queda por fuera",
          f"Habla: Ciencia de Datos\n\n\"Un modelo que lee solo el título saca F1 {n(A['f1_solo_titulo'])}; leyendo el comentario saca "
          f"{n(A['f1_solo_comentario'])}. Que tres palabras le ganen al comentario completo demuestra que el título se escribe junto con "
          f"las estrellas: se excluye, igual que las estrellas. Partimos 70/15/15, estratificado y con semilla 42. Como BETO en CPU no puede "
          f"evaluar 73 mil comentarios, todos los modelos se comparan además sobre la misma validación y prueba comunes de 5.000.\"")
cd = CategoryChartData()
cd.categories = ["Solo título", "Solo comentario"]
cd.add_series("F1 macro", [A["f1_solo_titulo"], A["f1_solo_comentario"]])
ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.4), Inches(1.3), Inches(4.4), Inches(1.9), cd).chart
estilo_grafico(ch)
ejes(ch, 0, 1, grid=False)
ch.value_axis.visible = False
pl = ch.plots[0]
pl.gap_width = 50
pl.has_data_labels = True
pl.data_labels.number_format = '0.000'
pl.data_labels.number_format_is_linked = False
pl.data_labels.font.size = Pt(11)
for i, col in enumerate([RED, BLUE]):
    pt = pl.series[0].points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = rgb(col)
texto(s, 0.5, 3.25, 4.3, 0.9, "Mismo TF-IDF + SVM, misma partición. Excluidas: review_title y review_rate. Vocabulario, IDF y "
      "tokenizador se ajustan solo con entrenamiento.", size=11, color=MUTED)
for i, (nom, num, uso, col) in enumerate([("ENTRENAMIENTO · 70%", entero(A["n_train"]), "Aprende pesos, vocabulario e IDF", BLUE),
                                          ("VALIDACIÓN · 15%", entero(A["n_val"]), "Elige C, épocas, umbral y temperatura", TEAL),
                                          ("PRUEBA · 15%", entero(A["n_test"]), "Se usa una sola vez para reportar", NAVY)]):
    y = 1.3 + i * 1.05
    caja(s, 5.2, y, 4.3, 0.9, LIGHT)
    texto(s, 5.4, y + 0.1, 2.2, 0.3, nom, size=11, bold=True, color=col)
    texto(s, 7.6, y + 0.07, 1.75, 0.4, num, size=18, bold=True, color=col, align=PP_ALIGN.RIGHT)
    texto(s, 5.4, y + 0.45, 3.9, 0.35, uso, size=11)
texto(s, 5.2, 4.5, 4.3, 0.6, [[("Comparación justa: ", {"bold": True, "color": NAVY}),
                              (f"validación y prueba comunes de {entero(A['n_comun'])} comentarios para los 8 modelos.", {})]], size=11)

# ------------------------------------------------------------------ 6. Tres familias
s = nueva("Modelado · " + f"{POR_ROL[R_ML]} · Especialista en Machine Learning", "Tres familias de modelos sobre la misma partición",
          f"Habla: Machine Learning\n\n\"La línea base convierte el texto en TF-IDF de palabras y bigramas y lo clasifica con regresión "
          f"logística o SVM. Las redes de la sesión 3 aprenden un embedding de 64 valores por palabra y lo leen en orden con RNN, LSTM o GRU. "
          f"BETO, de la sesión 4, ya llega preentrenado con millones de textos en español: divide en subpalabras y cada token mira a todos "
          f"los demás con atención. Las tres familias usan los mismos pesos de clase y la misma partición.\"")
filas_fam = [("TF-IDF + LR / SVM", "Sin orden", ["Texto", "Palabras + bigramas", "TF-IDF 300k", "LR / SVM (C en val.)", "neg / neu / pos"], "94A3B8"),
             ("RNN · LSTM · GRU", "Lee en orden · sesión 3", ["Texto", "Índices (vocab 30k)", "Embedding 64", "Recurrente 64 u.", "Softmax"], BLUE),
             ("BETO (Transformer)", "Atención · preentrenado", ["Texto", "WordPiece (31k)", "12 capas × 12 cabezas", "[CLS] → lineal", "Softmax"], TEAL)]
for i, (nom, sub, pasos, col) in enumerate(filas_fam):
    y = 1.3 + i * 1.3
    texto(s, 0.5, y + 0.12, 1.9, 0.3, nom, size=13, bold=True, color=col if col != "94A3B8" else MUTED)
    texto(s, 0.5, y + 0.45, 1.9, 0.5, sub, size=10, color=MUTED)
    for j, p in enumerate(pasos):
        x = 2.45 + j * 1.43
        b = caja(s, x, y, 1.28, 0.95, col if j in (2, 3) else LIGHT)
        tf = b.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.text = p
        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        para.runs[0].font.size = Pt(11)
        para.runs[0].font.bold = j in (2, 3)
        para.runs[0].font.color.rgb = rgb(WHITE if j in (2, 3) else INK)
        if j < 4:
            a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 1.29), Inches(y + 0.4), Inches(0.13), Inches(0.15))
            a.fill.solid()
            a.fill.fore_color.rgb = rgb(LINE)
            a.line.fill.background()

# ------------------------------------------------------------------ 7. Atención y tokenizador
s = nueva("Sesión 4 · " + f"{POR_ROL[R_ML]} · Especialista en Machine Learning", "Atención: cada palabra mira a todas las demás a la vez",
          "Habla: Machine Learning\n\n\"La RNN tiene que cargar el inicio del comentario paso a paso. El Transformer calcula, para cada "
          "token, qué tan relevante es cada otro token: softmax de Q por K transpuesta sobre raíz de d, multiplicado por V. A la izquierda, "
          "la atención real de la última capa de nuestro BETO; a la derecha, la oclusión: cuánto cae la probabilidad si borramos cada "
          "palabra. La atención no es una explicación causal; por eso la acompañamos con oclusión. Abajo, el tokenizador: BETO parte las "
          "palabras raras en subpalabras, así que el costo se mide en subpalabras.\"")
imagen(s, "interpretabilidad.png", 0.4, 1.2, w=6.0)
texto(s, 0.5, 3.45, 5.9, 1.6, [[("Ejemplo: ", {"bold": True, "color": NAVY}),
                              ("“No llegó completo, pero el vendedor respondió rápido y lo solucionó”.", {"italic": True})],
                             [("Izquierda: ", {"bold": True, "color": NAVY}),
                              ("qué tokens mezcla cada posición en la última capa (promedio de 12 cabezas). ", {}),
                              ("Derecha: ", {"bold": True, "color": NAVY}),
                              ("cuánto cae la probabilidad de la clase predicha al borrar cada palabra. La atención no es una explicación "
                               "causal: por eso se complementa con oclusión.", {})]], size=11)
caja(s, 6.65, 1.2, 2.85, 1.25, LIGHT)
texto(s, 6.8, 1.3, 2.6, 1.1, [[("Attention(Q,K,V) =", {"bold": True, "color": NAVY})],
                             "softmax(QKᵀ/√dₖ + M) · V",
                             [("Q: qué busca · K: qué ofrece · V: qué aporta · M: máscara de relleno", {"size": 10, "color": MUTED})]], size=12)
aud = {x["término"]: x for x in s4["aud_tokenizador"] if x["tokenizador"] == "BETO"}
filas_tok = [["Término", "Subpalabras BETO"]] + [[t, aud[t]["subpalabras"]] for t in ["malísimo", "exelente", "MercadoPago", "devolución"] if t in aud]
tabla(s, 6.65, 2.6, 2.85, filas_tok, [1.15, 1.7], size=10)
truncados = A["tokens_bert"]["truncados_" + str(f["max_len"])]
texto(s, 6.65, 4.15, 2.85, 0.9, f"{n(A['tokens_bert']['subpalabras_por_palabra'], 2)} subpalabras por palabra · max_length "
      f"{f['max_len']} trunca el {pct(truncados)} de los comentarios.", size=11, color=MUTED)

# ------------------------------------------------------------------ 8. Fine-tuning
hist = [h for h in s4["historia_bert"] if "eval_f1_macro" in h]
s = nueva("Sesión 4 · " + f"{POR_ROL[R_ML]} · Especialista en Machine Learning", "Fine-tuning de BETO con las reglas de la sesión 4",
          f"Habla: Machine Learning\n\n\"Ajustamos todo BETO, 110 millones de parámetros, con AdamW, learning rate 2e-5, warmup del 10%, "
          f"scheduler lineal y padding dinámico. Sin GPU, entrenamos con {entero(f['n_train_bert'])} comentarios estratificados, no con "
          f"340 mil: tardó {n(f['seg_bert'] / 60, 0)} minutos. Para que la comparación sea justa, entrenamos también una TF-IDF con exactamente "
          f"los mismos {entero(f['n_train_bert'])} comentarios. La mejor época se eligió con F1 en validación.\"")
tabla(s, 0.5, 1.3, 4.4, [["Hiperparámetro", "Valor"], ["Checkpoint", "dccuchile/bert-base-spanish-wwm-cased"],
                         ["Estrategia", "Fine-tuning completo"], ["Optimizador", "AdamW · lr 2e-5 · decay 0,01"],
                         ["Warmup / scheduler", "10% · lineal"], ["Batch · épocas", f"16 · {len(hist)} (mejor por F1 val.)"],
                         ["max_length", f"{f['max_len']} · padding dinámico"], ["Pérdida", "Entropía cruzada con pesos de clase"],
                         ["Datos de entrenamiento", f"{entero(f['n_train_bert'])} (estratificado)"],
                         ["Hardware", f"CPU {f['entorno']['nucleos_logicos']} hilos · {f['entorno']['ram_gb'] / 1.0737:.0f} GB · sin GPU"]],
      [1.6, 2.8], size=10)
for i, (v, et, col) in enumerate([(n(f["seg_bert"] / 60, 0) + " min", "de entrenamiento en CPU", ORANGE),
                                  ("110 M", "parámetros ajustados", NAVY),
                                  (n(f["f1_val"][BERT]), f"F1 macro en validación común ({corto(BERT)})", TEAL)]):
    dato(s, 5.4, 1.3 + i * 1.25, 4.0, v, et, col, size=30)

# ------------------------------------------------------------------ 9. Resultados
orden = sorted(s4["comun"], key=lambda k: f["f1"][k])
s = nueva("Resultados · " + f"{POR_ROL[R_ML]} · Especialista en Machine Learning", "Comparación justa: los ocho modelos sobre la misma prueba",
          f"Habla: Machine Learning\n\n\"Este es el resultado central: F1 macro sobre los mismos 5.000 comentarios de prueba. "
          f"{corto(BERT)} logra {n(f['f1'][BERT])}; el mejor recurrente, {S3}, {n(f['f1'][S3])}; la TF-IDF con todos los datos, "
          f"{n(f['f1'][BASE])}. Fíjense en la barra naranja: la TF-IDF con los mismos {entero(f['n_train_bert'])} comentarios que BETO "
          f"saca {n(f['f1'][SUB])}. Esa diferencia es lo que aporta el preentrenamiento con el mismo presupuesto de datos.\"")
cd = CategoryChartData()
cd.categories = [corto(k) for k in orden]
cd.add_series("F1 macro (prueba común)", [f["f1"][k] for k in orden])
ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.3), Inches(1.2), Inches(5.6), Inches(3.95), cd).chart
estilo_grafico(ch)
vmin = max(0, round(min(f["f1"].values()) - 0.08, 1))
ejes(ch, vmin, None, grid=False)
ch.value_axis.visible = False
pl = ch.plots[0]
pl.gap_width = 40
pl.has_data_labels = True
pl.data_labels.number_format = '0.000'
pl.data_labels.number_format_is_linked = False
pl.data_labels.font.size = Pt(10)
for i, k in enumerate(orden):
    pt = pl.series[0].points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = rgb(COL_MODELO.get(k, "94A3B8"))
filas_r = [["Modelo", "Recall neg.", "F1 neutro", "p95 ms"]]
for k in [BERT, S3, BASE, SUB]:
    filas_r.append([corto(k), n(f["rec_neg"][k]), n(f["f1_neu"][k]), n(f["p95"][k], 1)])
tabla(s, 6.1, 1.35, 3.4, filas_r, [1.55, 0.65, 0.6, 0.6], size=9.5, resaltar=1, alto_fila=0.34)
texto(s, 6.1, 3.25, 3.4, 1.8, [[("Cómo leerlo. ", {"bold": True, "color": NAVY}),
                               ("El campeón se elige por F1 en validación común; la prueba se usa una sola vez. El F1 macro pesa igual las "
                                "tres clases, así que el neutro cuenta tanto como las demás.", {})]], size=11)

# ------------------------------------------------------------------ 10. Por clase + IC
s = nueva("Resultados · " + f"{POR_ROL[R_ML]} · Especialista en Machine Learning", "¿La ventaja es real? Clase por clase y con intervalo de confianza",
          "Habla: Machine Learning\n\n\"Todos los modelos son buenos en negativo y positivo; la diferencia se juega en el neutro. Y no basta "
          "con ver una diferencia: remuestreamos la prueba mil veces. Si el intervalo del 95% no cruza el cero, la diferencia no es suerte.\"")
cd = CategoryChartData()
cd.categories = ["Negativo", "Neutro", "Positivo"]
for k in [BASE, S3, BERT]:
    c = s4["comun"][k]["test"]
    cd.add_series(corto(k), [c["f1_negativo"], c["f1_neutro"], c["f1_positivo"]])
ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.3), Inches(1.2), Inches(5.4), Inches(3.95), cd).chart
estilo_grafico(ch, leyenda=True)
ejes(ch, 0.3, 1.0, grid=True)
pl = ch.plots[0]
pl.gap_width = 70
pl.overlap = -10
pl.has_data_labels = True
pl.data_labels.number_format = '0.00'
pl.data_labels.number_format_is_linked = False
pl.data_labels.font.size = Pt(9)
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
for serie, k in zip(pl.series, [BASE, S3, BERT]):
    serie.format.fill.solid()
    serie.format.fill.fore_color.rgb = rgb(COL_MODELO.get(k, "94A3B8"))
filas_ic = [["Diferencia de F1 macro", "Δ", "IC 95%"]]
for k, (m, a, b) in f["ic"].items():
    filas_ic.append([k.replace("TF-IDF + Regresión Logística", "TF-IDF"), ("+" if m >= 0 else "") + n(m), f"[{n(a)}; {n(b)}]"])
tabla(s, 5.9, 1.35, 3.6, filas_ic, [1.75, 0.65, 1.2], size=9.5, alto_fila=0.4)
texto(s, 5.9, 3.15, 3.6, 1.9, "Bootstrap de 1.000 remuestreos de la prueba común (semilla 42). Si el intervalo no contiene el cero, "
      "la diferencia es estadísticamente significativa.", size=11, color=MUTED)

# ------------------------------------------------------------------ 11. Errores y segmentos
tr = s4["segmentos"]["tramo"]
causas = s4["errores_causas"]
s = nueva("Errores · " + f"{POR_ROL[R_DS]} · Magíster en Ciencia de Datos", "Dónde se equivocan los modelos",
          f"Habla: Ciencia de Datos\n\n\"Revisamos los errores después de congelar el modelo. Por longitud, la red y BETO ganan sobre todo en "
          f"los comentarios largos, donde el orden importa. La mayoría de los errores de BETO están en la frontera con el neutro: comentarios "
          f"de 3 estrellas que suenan positivos o negativos. Parte del error es de la etiqueta, no del modelo, porque la etiqueta sale de las "
          f"estrellas. Los negativos leídos como positivos son {s4['neg_como_pos'].get('BETO')} con BETO frente a "
          f"{s4['neg_como_pos'].get(BASE)} con TF-IDF.\"")
cd = CategoryChartData()
tramos = [t for t in ["1-10", "11-25", "26-50", "51-100", ">100"] if t in tr]
cd.categories = [t + " pal." for t in tramos]
for k in [BASE, S3, "BETO"]:
    cd.add_series(corto(k), [tr[t][k] for t in tramos])
ch = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.3), Inches(1.2), Inches(5.3), Inches(3.95), cd).chart
estilo_grafico(ch, leyenda=True)
ejes(ch, None, None)
ch.has_title = True
ch.chart_title.text_frame.text = "F1 macro por longitud del comentario (prueba común)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
for serie, k in zip(ch.plots[0].series, [BASE, S3, "BETO"]):
    serie.format.line.color.rgb = rgb(COL_MODELO.get(k, "94A3B8"))
    serie.format.line.width = Pt(2.25)
    serie.smooth = False
    serie.marker.style = XL_MARKER_STYLE.CIRCLE
    serie.marker.format.fill.solid()
    serie.marker.format.fill.fore_color.rgb = rgb(COL_MODELO.get(k, "94A3B8"))
tot = sum(causas.values()) or 1
filas_c = [["Causa probable (errores de BETO)", "%"]] + [[c, pct(v / tot, 0)] for c, v in list(causas.items())[:5]]
tabla(s, 5.8, 1.35, 3.7, filas_c, [3.0, 0.7], size=10)
neg = {k: s4["neg_como_pos"].get(k) for k in [BASE, S3, "BETO"]}
texto(s, 5.8, 3.0, 3.7, 0.6, [[("Errores que la abstención manda a revisión: ", {"bold": True, "color": NAVY}),
                              (pct(s4.get("errores_abstencion"), 0), {"bold": True})]], size=11)
texto(s, 5.8, 3.55, 3.7, 1.5, [[("Error grave (negativo → positivo): ", {"bold": True, "color": RED})],
                              "  ·  ".join(f"{corto(k)} {v}" for k, v in neg.items())], size=11)

# ------------------------------------------------------------------ 12. Robustez
casos = s4["casos_robustez"]
mods = [BASE, S3, "BETO"]
s = nueva("Robustez · " + f"{POR_ROL[R_DS]} · Magíster en Ciencia de Datos", "Contrafactuales: negación, contraste y lo que no sabe",
          "Habla: Ciencia de Datos\n\n\"Probamos pares mínimos: la misma frase con y sin negación, con el contraste invertido, con errores de "
          "ortografía, mayúsculas, emojis e ironía, y textos que no tienen nada que ver con compras. Lo que buscamos en los textos fuera de "
          "dominio es que el modelo diga 'no sé' y mande el caso a una persona. Y encontramos un sesgo importante: en este corpus casi todos los comentarios cortos son negativos, así que los modelos aprendieron que corto es igual a negativo. Por eso los comentarios cortos no dispararán alertas automáticas.\"")
orden_casos = ["negación", "negación (larga)", "contraste", "contraste (largo)", "fuera de dominio", "fuera de dominio (largo)"]
sel = [x for o in orden_casos for x in casos if x["caso"] == o][:8]
filas_rb = [["Caso", "Texto", "Esperado"] + [corto(k) for k in mods]]
colores = {}
COLOR_EST = {"acierta": "D1FAE5", "se abstiene": "E0E7FF", "falla": "FEE2E2"}
for i, x in enumerate(sel, 1):
    fila = [x["caso"], x["texto"][:40] + ("…" if len(x["texto"]) > 40 else ""), x["esperado"]]
    for j, k in enumerate(mods):
        v = x[k]
        fila.append(("↷ " if "revisión" in v else "") + v.split(" ")[0])
        colores[(i, 3 + j)] = COLOR_EST.get(x.get(f"estado_{k}"), "FFFFFF")
    filas_rb.append(fila)
tabla(s, 0.5, 1.15, 9.0, filas_rb, [1.45, 3.05, 1.0, 1.15, 1.15, 1.2], size=9, alto_fila=0.29, colores_celda=colores)
rr = {(x["modelo"], x["grupo"]): x for x in s4.get("resumen_robustez", [])}
txt_rr = " · ".join(f"{corto(k)}: cortos {rr[(k, 'cortos (≤ 15 palabras)')]['acierta']}/{rr[(k, 'cortos (≤ 15 palabras)')]['casos']}, "
                    f"largos {rr[(k, 'largos (> 15 palabras)')]['acierta']}/{rr[(k, 'largos (> 15 palabras)')]['casos']}"
                    for k in mods if (k, "cortos (≤ 15 palabras)") in rr)
texto(s, 0.5, 3.83, 9, 0.3, "Verde = acierta · azul = se abstiene · rojo = falla · ↷ = revisión humana. Aciertos " + txt_rr, size=9, color=MUTED)
if SES:
    caja(s, 0.5, 4.15, 9.0, 0.95, "FEF3C7")
    rb, rg, rl = (ses_mod.get((k, "≤ 15 palabras"), {}).get("recall positivo", 0) for k in ["BETO", S3, BASE])
    texto(s, 0.7, 4.22, 8.6, 0.85, [[("Hallazgo: sesgo de longitud. ", {"bold": True, "color": "92400E"}),
                                    (f"En MeLiSA el {pct(SES['pct_negativo_le15'], 0)} de los comentarios de ≤ 15 palabras son negativos: "
                                     f"los modelos aprendieron “corto ⇒ negativo”. Recall de positivos cortos: BETO {pct(rb, 0)}, {S3} "
                                     f"{pct(rg, 0)}, TF-IDF {pct(rl, 0)}. Con más contexto BETO deja de responder “negativo” a todo; los textos ajenos van a revisión.", {})]],
          size=11, anchor=MSO_ANCHOR.MIDDLE)

# ------------------------------------------------------------------ 13. Calibración y abstención
s = nueva("Abstención · " + f"{POR_ROL[R_GIA]} · Magíster en Gerencia en IA", "Si el modelo no está seguro, decide una persona",
          f"Habla: Gerencia de IA\n\n\"Primero calibramos: con temperature scaling la confianza de 0,8 significa acertar cerca del 80%. "
          f"Luego, si la confianza es menor que el umbral, el comentario pasa a una persona. El umbral se eligió en validación para lograr "
          f"recall de negativo de 0,92 en lo que el modelo decide. Con {corto(REC)}, en la prueba común decide el "
          f"{pct(ab_rec.get('cobertura'))} y manda a revisión {entero(1000 * (ab_rec.get('abstencion') or 0))} de cada 1.000 comentarios.\"")
curvas = s4["curvas"]
cd = XyChartData()
for k in [BASE, S3, BERT]:
    if k in curvas:
        se = cd.add_series(corto(k))
        for g in curvas[k]:
            se.add_data_point(g["cobertura"], g["accuracy_auto"])
ch = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES, Inches(0.3), Inches(1.2), Inches(5.4), Inches(3.95), cd).chart
estilo_grafico(ch, leyenda=True)
ch.value_axis.tick_labels.number_format = '0%'
ch.value_axis.tick_labels.number_format_is_linked = False
ch.value_axis.major_gridlines.format.line.color.rgb = rgb("E2E8F0")
ch.category_axis.tick_labels.number_format = '0%'
ch.category_axis.tick_labels.number_format_is_linked = False
ch.category_axis.has_title = True
ch.category_axis.axis_title.text_frame.text = "Cobertura automática (calibrada, prueba común)"
ch.category_axis.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
ch.value_axis.has_title = True
ch.value_axis.axis_title.text_frame.text = "Accuracy en casos decididos"
ch.value_axis.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
ch.category_axis.minimum_scale = 0.3
ch.category_axis.maximum_scale = 1.0
ch.value_axis.maximum_scale = 1.0
ch.value_axis.minimum_scale = max(0.0, round(min(g["accuracy_auto"] for k in [BASE, S3, BERT] if k in curvas for g in curvas[k]) - 0.05, 1))
for serie, k in zip(ch.plots[0].series, [BASE, S3, BERT]):
    serie.format.line.color.rgb = rgb(COL_MODELO.get(k, "94A3B8"))
    serie.format.line.width = Pt(2)
    serie.smooth = False
    serie.marker.style = XL_MARKER_STYLE.NONE
filas_a = [["Modelo", "Umbral", "Cobertura", "Recall neg."]]
for k in [REC, BERT, S3, BASE]:
    if k in f["abst"] and filas_a[-1][0] != corto(k) and corto(k) not in [x[0] for x in filas_a]:
        a = f["abst"][k] or {}
        filas_a.append([corto(k), n(f["umbral"].get(k), 2), pct(a.get("cobertura")), n(a.get("recall_neg_auto"))])
tabla(s, 5.9, 1.35, 3.6, filas_a, [1.4, 0.65, 0.8, 0.75], size=9.5, resaltar=1, alto_fila=0.34)
cal = s4["calibracion"][BERT]
texto(s, 5.9, 3.25, 3.6, 1.8, [[("Calibración de " + corto(BERT) + ": ", {"bold": True, "color": NAVY}),
                               (f"T = {n(cal['T'], 2)}; ECE {n(cal['ece_antes'])} → {n(cal['ece_despues'])}. ", {}),
                               ("El umbral se elige en validación y se reporta en prueba.", {"color": MUTED})]], size=11)

# ------------------------------------------------------------------ 14. Costo y latencia
ef = {e["modelo"]: e for e in s4["eficiencia"]}
s = nueva("Costo · " + f"{POR_ROL[R_GIA]} · Magíster en Gerencia en IA", "Calidad frente a costo: cuánto cuesta cada punto de F1",
          f"Habla: Gerencia de IA\n\n\"Medimos todo en el mismo portátil, sin GPU. TF-IDF responde en menos de un milisegundo; BETO, en "
          f"{n(f['p95'].get('BETO'), 0)} ms en el percentil 95, y la versión cuantizada a int8 en {n(f['p95'].get('BETO int8'), 0)} ms. "
          f"Con el supuesto de una máquina de 4 vCPU a 0,20 dólares la hora, procesar un millón de comentarios cuesta centavos con TF-IDF y "
          f"unos pocos dólares con BETO: el costo de cómputo es bajo frente al valor de no perder un cliente molesto.\"")
_a, _m = caso_valor(ab_rec.get("cobertura") or 0, ab_rec.get("recall_neg_efectivo") or 0, ef[REC]["USD por millón"])
_neto_inc = _m["neto"] - _a["neto"]
mods_ef = [k for k in [BASE, S3, "BETO", "BETO int8"] if k in ef]
filas_e = [["Modelo", "F1 prueba", "p50 ms", "p95 ms", "Coment./s", "Parámetros", "US$ / millón"]]
for k in mods_ef:
    e = ef[k]
    filas_e.append([corto(k), n(e["F1 test común"]), n(e["p50 ms"], 1), n(e["p95 ms"], 1), entero(e["coment./s"]),
                    (n(e["parámetros (M)"], 1) + " M") if e["parámetros (M)"] >= 1 else entero(e["parámetros (M)"] * 1e6),
                    n(e["USD por millón"], 2)])
tabla(s, 0.5, 1.3, 9.0, filas_e, [2.3, 1.0, 1.0, 1.0, 1.2, 1.25, 1.25], size=11, alto_fila=0.38)
for i, (v, et, col) in enumerate([
        (n(f["p95"].get("BETO", 0) / max(f["p95"][BASE], 1e-6), 0) + "×", "más lento BETO que TF-IDF (p95, un comentario)", ORANGE),
        (n(f["p95"].get("BETO int8", 0), 0) + " ms", "p95 de BETO int8 (límite de la sesión 4: 150 ms)", NAVY),
        ("US$" + entero(_neto_inc / 1000) + " mil", "beneficio neto incremental al mes (supuestos ilustrativos, ver informe §7)", GREEN)]):
    dato(s, 0.5 + i * 3.05, 3.55, 2.8, v, et, col, size=28)

# ------------------------------------------------------------------ 15. Criterio de adopción
crit = s4["criterio"]
mods_c = list(dict.fromkeys(c["modelo"] for c in crit))
s = nueva("Decisión · " + f"{POR_ROL[R_GIA]} · Magíster en Gerencia en IA", "Criterio de adopción de la sesión 4",
          f"Habla: Gerencia de IA\n\n\"El criterio de la sesión 4 exige cuatro cosas a la vez: mejorar el F1 al menos 0,02 frente al campeón "
          f"anterior, recall de negativo de 0,92 con abstención, cobertura automática de al menos 75% y latencia p95 de 150 ms o menos. "
          f"{(D['criterio_estricto'] + ' ' + D['por_que']) if D else ''}\"")
criterios = list(dict.fromkeys(c["criterio"] for c in crit))
filas_k = [["Criterio", "Umbral"] + [corto(m) for m in mods_c]]
colores = {}
for i, cr in enumerate(criterios, 1):
    fila = [cr, ""]
    for j, m in enumerate(mods_c):
        c = next((x for x in crit if x["modelo"] == m and x["criterio"] == cr), None)
        if c is None:
            fila.append("—")
            continue
        fila[1] = c["umbral"].replace("Regresión Logística", "LR") if j == 0 else fila[1]
        v = c["valor"]
        if v is None or v != v:
            txt = "—"
        elif "Latencia" in cr:
            txt = n(v, 1) + " ms"
        elif "Cobertura" in cr:
            txt = pct(v)
        elif "Ganancia" in cr:
            txt = ("+" if v >= 0 else "") + n(v)
        elif "lote" in cr:
            txt = n(v, 2 if v < 1 else 1) + " min"
        else:
            txt = n(v)
        fila.append(f"{ok if c['cumple'] else no}  {txt}")
        colores[(i, 2 + j)] = "D1FAE5" if c["cumple"] else "FEE2E2"
    filas_k.append(fila)
filas_k[1][1] = f"F1 ≥ campeón anterior + 0,02"
tabla(s, 0.5, 1.3, 9.0, filas_k, [2.4, 2.1] + [4.5 / len(mods_c)] * len(mods_c), size=10.5, alto_fila=0.38, colores_celda=colores)
texto(s, 0.5, 3.7, 9.0, 0.3, "Capacidad en lote: criterio operativo fijado en la formulación (proceso por lotes horarios). Referencia: BETO y BETO int8 se comparan con " + S3 + "; " + S3 + " se compara con " + corto(BASE) + ".",
      size=10, color=MUTED)
caja(s, 0.5, 4.1, 9.0, 0.95, "E0F2F4")
texto(s, 0.7, 4.15, 8.6, 0.85, ((D["criterio_estricto"] + " ") if D else "") + decision_corta, size=11.5, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)

# ------------------------------------------------------------------ 16. Riesgos
s = nueva("Gobierno · " + f"{POR_ROL[R_GIA]} · Magíster en Gerencia en IA", "Riesgos, controles y condiciones de suspensión",
          "Habla: Gerencia de IA\n\n\"Cada riesgo tiene un control y un responsable; la matriz completa está en el Excel. Lo más importante: "
          "el modelo solo prioriza y alerta, nunca sanciona; la PII se enmascara antes de inferir; reportamos F1 por país; y tenemos "
          "condiciones explícitas para suspender o volver al modelo simple.\"")
riesgos = [("Texto fuera de dominio", "Control por similitud con reseñas (BETO)", POR_ROL[R_ML], RED),
           ("Negativo leído como positivo", "Abstención con recall ≥ 0,92", POR_ROL[R_ML], RED),
           ("Sesgo de longitud y país", "Sin alertas en cortos; F1 por país", POR_ROL[R_DS], ORANGE),
           ("Drift y etiqueta ruidosa", "Monitoreo semanal; reetiquetar 300/mes", POR_ROL[R_DS], ORANGE),
           ("PII en el texto", "Enmascarar antes de inferir y registrar", POR_ROL[R_GIA], ORANGE),
           ("Uso para sancionar", "Solo prioriza; un agente confirma", POR_ROL[R_GIA], GREEN)]
for i, (rsg, ctl, resp, col) in enumerate(riesgos):
    x, y = 0.5 + (i % 2) * 2.85, 1.3 + (i // 2) * 1.25
    caja(s, x, y, 2.7, 1.12, LIGHT)
    c = caja(s, x + 0.15, y + 0.17, 0.14, 0.14, col, forma=MSO_SHAPE.OVAL)
    texto(s, x + 0.38, y + 0.1, 2.25, 0.3, rsg, size=12, bold=True, color=NAVY)
    texto(s, x + 0.15, y + 0.42, 2.45, 0.4, ctl, size=10.5)
    texto(s, x + 0.15, y + 0.82, 2.45, 0.25, resp, size=9.5, color=MUTED)
caja(s, 6.35, 1.3, 3.15, 3.7, NAVY)
texto(s, 6.55, 1.42, 2.8, 0.3, "Suspender o revertir si…", size=13, bold=True, color=WHITE)
texto(s, 6.55, 1.85, 2.8, 3.1, ["Recall de negativo < 0,88 dos semanas seguidas", "Abstención > 35% una semana",
                               "F1 de un país < 0,65 en la muestra reetiquetada", "Cualquier PII expuesta en logs",
                               "Precisión de alertas < 80%"], size=11, color="E2E8F0", espacio=8)

# ------------------------------------------------------------------ 17. Demo y roadmap
cfg_ej = {"etiqueta": "negativo", "confianza": "0,94", "decision": "automática", "alerta_vendedor": "true",
          "evidencia": "dejó de funcionar"}
s = nueva(f"Demostración · {POR_ROL[R_ML]} (Machine Learning) y {POR_ROL[R_DS]} (Ciencia de Datos)", "El flujo completo, de la entrada a la decisión",
          "Habla: Machine Learning y Gerencia de IA\n\n\"En la demo escribimos un comentario. El sistema valida la entrada, enmascara "
          "correos y teléfonos, calcula probabilidades calibradas y aplica el umbral: si está seguro decide y, si es negativo, alerta al "
          "vendedor; si no, lo manda a una persona. Además muestra la evidencia, las palabras que más pesaron. El código está en src/demo.py "
          "y en la sección 10 del notebook. La adopción va por fases: modo sombra, humano en el ciclo y solo después alertas automáticas.\"")
pasos = ["Comentario", "Validar y enmascarar PII", "¿Fuera de dominio?", f"{corto(REC)} + temperatura", "¿Confianza ≥ umbral?", "Etiqueta + evidencia"]
for j, p in enumerate(pasos):
    x = 0.5 + j * 1.53
    b = caja(s, x, 1.3, 1.37, 0.8, TEAL if j in (2, 4) else LIGHT)
    tf = b.text_frame
    tf.word_wrap = True
    tf.text = p
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.paragraphs[0].runs[0].font.size = Pt(10.5)
    tf.paragraphs[0].runs[0].font.bold = True
    tf.paragraphs[0].runs[0].font.color.rgb = rgb(WHITE if j in (2, 4) else NAVY)
    if j < 5:
        a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 1.38), Inches(1.62), Inches(0.13), Inches(0.16))
        a.fill.solid()
        a.fill.fore_color.rgb = rgb(LINE)
        a.line.fill.background()
texto(s, 0.5, 2.2, 9, 0.3, "Fuera de dominio o confianza baja → revisión humana · ≤ 15 palabras → sin alerta automática · BETO en línea: huggingface.co/spaces/pedrogarciar13/sentimiento-melisa-beto · Respaldo TF-IDF: sentimiento-melisa-demo.onrender.com", size=10.5, color=MUTED)
caja(s, 0.5, 2.6, 4.3, 2.45, "0F172A")
ruta_demo = os.path.join(RAIZ, "resultados", "demo_ejemplo.json")
if os.path.exists(ruta_demo):
    import json
    d = json.load(open(ruta_demo, encoding="utf-8"))
    lineas = [f'$ python src/demo.py "{d["entrada"][:38]}…"' if len(d["entrada"]) > 38 else f'$ python src/demo.py "{d["entrada"]}"',
              f'{{ "etiqueta": "{d["etiqueta"]}",', f'  "confianza": {d["confianza"]},', f'  "decision": "{d["decision"]}",',
              f'  "alerta_vendedor": {str(d["alerta_vendedor"]).lower()},',
              '  "evidencia": [' + ", ".join(f'"{e}"' for e in d["evidencia"][:3]) + '] }']
    pie = f"Salida real de src/demo.py con {corto(d['modelo'])}."
else:
    lineas = ['$ python src/demo.py "<comentario>"', '{ "etiqueta": ..., "confianza": ...,', '  "decision": ..., "evidencia": [...] }']
    pie = "Ejecute src/demo.py para ver la salida."
texto(s, 0.7, 2.72, 4.0, 2.0, lineas, size=10, color="A7F3D0", font="Courier New", espacio=2)
texto(s, 0.7, 4.75, 4.0, 0.25, pie, size=8.5, color="94A3B8")
for i, (fase, dur, que) in enumerate([("1 · Modo sombra", "4 semanas", "Predice sin actuar; se compara con estrellas y agentes"),
                                      ("2 · Humano en el ciclo", "6 semanas", "Negativos a cola prioritaria; un agente confirma"),
                                      ("3 · Alertas automáticas", "Continuo", "Aviso al vendedor si suben sus negativos"),
                                      ("4 · Mejora continua", "Trimestral", "Reentrenar, reetiquetar neutro, aspectos")]):
    y = 2.6 + i * 0.62
    texto(s, 5.1, y, 2.6, 0.3, fase, size=12, bold=True, color=NAVY)
    texto(s, 7.75, y, 1.75, 0.3, dur, size=11, color=TEAL, bold=True, align=PP_ALIGN.RIGHT)
    texto(s, 5.1, y + 0.27, 4.4, 0.3, que, size=10.5, color=MUTED)

# ------------------------------------------------------------------ 18. Cierre
s = nueva(None, None, f"Hablan: los tres roles\n\nMachine Learning: \"{corto(BERT)} es el mejor modelo en F1 sobre la misma prueba.\"\n"
          f"Ciencia de Datos: \"La evaluación es válida: sin título ni estrellas, sin duplicados entre conjuntos, umbral y temperatura "
          f"elegidos en validación y prueba usada una sola vez.\"\nGerencia de IA: \"{decision}\" Gracias; quedamos atentos a sus preguntas.",
          oscura=True)
texto(s, 0.6, 0.5, 8.8, 0.3, "RECOMENDACIÓN Y SIGUIENTES PASOS · LOS TRES ROLES", size=11, bold=True, color=TEAL)
texto(s, 0.6, 0.9, 8.8, 1.3, decision_corta, size=19, bold=True, color=WHITE)
for i, (t, d) in enumerate([("Siguiente experimento", f"Fine-tuning de BETO con GPU y los 340 mil comentarios; comparar con XLM-R y RoBERTuito."),
                            ("Datos", "Agregar comentarios cortos balanceados (sesgo de longitud) y reetiquetar una muestra del neutro."),
                            ("Producto", "Sentimiento por aspecto (envío, calidad, precio, atención) para que la alerta diga qué falló.")]):
    x = 0.6 + i * 2.95
    caja(s, x, 2.55, 2.75, 1.75, "17407A")
    texto(s, x + 0.18, 2.68, 2.4, 0.3, t, size=13, bold=True, color=TEAL)
    texto(s, x + 0.18, 3.05, 2.4, 1.2, d, size=11.5, color="E2E8F0")
texto(s, 0.6, 4.65, 8.8, 0.5, "La representación define lo que el modelo puede aprender; el negocio define si vale la pena.  "
      "Gracias · Preguntas", size=13, italic=True, color="CADCFC")

assert _num[0] == TOTAL, _num[0]
os.makedirs(os.path.join(RAIZ, "docs"), exist_ok=True)
salida = os.path.join(RAIZ, "docs", "Presentacion_Proyecto_PLN.pptx")
prs.save(salida)
print("Presentación ->", salida, f"({TOTAL} diapositivas)")
