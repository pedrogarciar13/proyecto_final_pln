"""Guion de la sustentación por rol (docs/Guion_Sustentacion_por_Rol.docx) y notas del orador de la presentación.

Uso: python src/generar_guion.py [raiz]
Debe ejecutarse después de src/generar_presentacion.py: reemplaza las notas de cada diapositiva por el guion.
"""
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation

sys.path.insert(0, os.path.dirname(__file__))
from resumen import cargar, cifras, entero, n, pct  # noqa: E402
from informe_texto import POR_ROL, caso_valor  # noqa: E402

RAIZ = sys.argv[1] if len(sys.argv) > 1 else "."
r = cargar(RAIZ)
f = cifras(r)
s4, A = r["s4"], f["audit"]
S3, BASE = f["s3"], f["base"]
ab = f["abst"]["BETO"] or {}
SES = f["sesgo"] or {}
rec_corto = {m["modelo"]: m["recall positivo"] for m in SES.get("modelos", []) if m["segmento"].startswith("≤")}
ic = f["ic"][f"BETO - {S3}"]

ML, DS, GIA = "Especialista en Machine Learning", "Magíster en Ciencia de Datos", "Magíster en Gerencia en IA"
OODT = (s4.get("ood") or {}).get("tabla", [{}])[0]
RRS = {(x["modelo"], x["grupo"]): x for x in s4.get("resumen_robustez", [])}
_a, _m = caso_valor(ab.get("cobertura") or 0, ab.get("recall_neg_efectivo") or 0, 1e6 / s4["info"]["BETO"]["thr"] / 3600 * 0.20)
NETO = _m["neto"] - _a["neto"]
def quien(rol):
    if rol == f"{ML} y {DS}": return f"{POR_ROL[ML]} (Machine Learning) y {POR_ROL[DS]} (Ciencia de Datos)"
    if rol == "Los tres roles": return "Los tres integrantes"
    return f"{POR_ROL[rol]} ({rol})"
COL = {ML: "1F6FB2", DS: "0FA3B1", GIA: "0B2E59", "Los tres roles": "475569"}

# (n.º, título, rol, minutos, guion, cifras clave, transición)
G = [
    (1, "Portada", GIA, 0.5,
     f"Buenos días. Soy {POR_ROL[GIA]}, de Gerencia en IA; me acompañan {POR_ROL[ML]}, de Machine Learning, y {POR_ROL[DS]}, de "
     "Ciencia de Datos. Vamos a responder una "
     "pregunta de negocio: ¿puede Mercado Libre saber qué siente un comprador leyendo solo su comentario, y cuánta complejidad vale "
     "la pena pagar para lograrlo? Comparamos tres familias de modelos: una línea base TF-IDF, las redes recurrentes de la sesión 3 "
     "y un Transformer preentrenado en español, BETO, de la sesión 4.",
     "3 familias · 8 modelos · 487 mil reseñas", "Empiezo por la conclusión."),
    (2, "La decisión en una diapositiva", GIA, 1.0,
     f"Sobre la misma prueba de {entero(A['n_comun'])} comentarios, BETO logra un F1 macro de {n(f['f1']['BETO'])}: "
     f"{n(f['delta_bert_s3'])} más que {S3}, la red recurrente elegida en validación, y {n(f['delta_bert_base'])} más que la línea base. "
     f"Con un umbral de confianza decide solo el {pct(ab.get('cobertura'))} de los comentarios con recall de negativo de "
     f"{n(ab.get('recall_neg_auto'))}; el resto lo revisa una persona. Nuestra recomendación es un piloto de BETO en modo sombra, "
     "por lotes, con dos controles: los comentarios cortos no disparan alertas automáticas y TF-IDF queda como respaldo. En las "
     "próximas diapositivas mostramos de dónde sale cada número.",
     f"F1 {n(f['f1']['BETO'])} · +{n(f['delta_bert_s3'])} vs {S3} · cobertura {pct(ab.get('cobertura'))} · recall neg. {n(ab.get('recall_neg_auto'))}",
     "¿Para quién es esto?"),
    (3, "Quién usa el sistema y qué decisión mejora", GIA, 1.25,
     "El usuario es el analista de experiencia del comprador. La decisión que apoyamos es concreta: qué comentario se atiende primero "
     "y a qué vendedor se le avisa a tiempo. Hoy se lee por muestreo, y las estrellas solo existen en las reseñas: los chats, las "
     "preguntas y los reclamos llegan solo con texto. La etiqueta la construimos con las estrellas: una y dos son negativo, tres es "
     "neutro, cuatro y cinco positivo. El error más caro es leer como positivo a un cliente molesto, porque se pierde la alerta. Por "
     "eso la clase crítica es negativo, y los KPI son el tiempo de primera respuesta a negativos, menos de 24 horas, y el porcentaje "
     "de alertas que un agente confirma.",
     "Clase crítica: negativo · KPI: respuesta < 24 h, alertas confirmadas ≥ 80%",
     f"Le paso la palabra a {POR_ROL[DS]}, que revisó si los datos permiten una evaluación honesta."),
    (4, "MeLiSA: 487 mil reseñas reales, auditadas", DS, 1.5,
     f"Usamos MeLiSA, un corpus público de la Universidad de Buenos Aires: {entero(A['filas'])} reseñas reales de 7 países y 5 "
     f"categorías. Antes de modelar auditamos, y cada hallazgo terminó en una decisión: {A['duplicados_texto']} textos repetidos se "
     f"dejan una sola vez, para que el mismo texto no quede en entrenamiento y en prueba; {A['textos_contradictorios']} textos con "
     f"sentimientos distintos se eliminan; y los {A['pii']['URL']} enlaces y {A['pii']['TEL']} posibles teléfonos se enmascaran antes "
     f"de entrenar. La mediana es de 30 palabras y el 99% tiene 118 o menos; en subpalabras de BETO el 95% cabe en 96, así que una "
     f"longitud máxima de 128 trunca menos del 2%. Quedaron {entero(A['filas_limpias'])} reseñas.",
     f"{entero(A['filas'])} reseñas · {A['duplicados_texto']} duplicados · {A['textos_contradictorios']} contradictorios · max_length 128",
     "El riesgo más grande de un proyecto así es la fuga de información."),
    (5, "Sin fuga: el título regala la respuesta", DS, 1.5,
     f"Hicimos una prueba sencilla: el mismo modelo, leyendo solo el título, saca F1 de {n(A['f1_solo_titulo'])}; leyendo el "
     f"comentario completo saca {n(A['f1_solo_comentario'])}. Que tres palabras le ganen al comentario demuestra que el título se "
     "escribe junto con las estrellas, y en un chat no existe: lo excluimos, igual que las estrellas, que son la fuente de la "
     f"etiqueta. Partimos 70, 15 y 15 por ciento, estratificado y con semilla 42. Validación sirve para elegir; prueba se usa una "
     f"sola vez. Como BETO en CPU no alcanza a evaluar 73 mil comentarios, todos los modelos se comparan además sobre la misma "
     f"validación y prueba comunes de {entero(A['n_comun'])}.",
     f"Solo título {n(A['f1_solo_titulo'])} > solo comentario {n(A['f1_solo_comentario'])} · 340.634 / 72.993 / 72.994",
     f"Con datos limpios y sin fuga, {POR_ROL[ML]} presenta los modelos."),
    (6, "Tres familias de modelos", ML, 1.0,
     "Comparamos tres formas de representar el texto. La línea base convierte cada comentario en TF-IDF de palabras y pares de "
     "palabras, como 'no llegó', y lo clasifica con regresión logística o SVM: no ve el orden. Las redes de la sesión 3 aprenden un "
     "vector de 64 números por palabra y leen el comentario en orden con RNN, LSTM o GRU; es many-to-one: solo el último estado "
     "decide. BETO, de la sesión 4, llega ya preentrenado con millones de textos en español, divide las palabras en subpalabras y usa "
     "atención. Las tres usan la misma partición y los mismos pesos de clase, porque el neutro es la mitad de frecuente.",
     "Misma partición · mismos pesos de clase · C y épocas elegidos en validación", "¿Qué tiene de distinto el Transformer?"),
    (7, "Atención: cada palabra mira a todas las demás", ML, 1.0,
     "La red recurrente tiene que cargar el inicio del comentario paso a paso, y lo va olvidando. El Transformer calcula, para cada "
     "palabra, qué tan relevante es cada una de las otras: softmax de Q por K transpuesta sobre raíz de d, multiplicado por V. Query "
     "es lo que busca la palabra, key lo que ofrece cada otra y value lo que aporta. A la izquierda está la atención real de la última "
     "capa de nuestro BETO; a la derecha, la oclusión: cuánto cambia la probabilidad si borramos cada palabra. Ojo: la atención no es "
     "una explicación causal, por eso la acompañamos con oclusión. Abajo, el tokenizador: palabras raras como 'exelente' se parten en "
     f"cuatro piezas, así que el costo se mide en subpalabras: {n(A['tokens_bert']['subpalabras_por_palabra'], 2)} por palabra.",
     "softmax(QKᵀ/√d)·V · 12 capas × 12 cabezas · atención ≠ causalidad", "¿Cómo lo adaptamos a nuestra tarea?"),
    (8, "Fine-tuning de BETO", ML, 1.0,
     f"Ajustamos todo BETO, 110 millones de parámetros, con las reglas de la sesión 4: AdamW con learning rate de 2e-5, warmup del 10%, "
     f"decaimiento lineal, padding dinámico y la mejor época elegida por F1 en validación. Sin GPU no era viable entrenar con 340 mil "
     f"comentarios, así que usamos {entero(f['n_train_bert'])} estratificados; tardó {n(f['seg_bert'] / 60, 0)} minutos. Para que la "
     f"comparación fuera justa, entrenamos también una TF-IDF con exactamente esos mismos {entero(f['n_train_bert'])} comentarios.",
     f"lr 2e-5 · 2 épocas · {entero(f['n_train_bert'])} comentarios · {n(f['seg_bert'] / 60, 0)} min en CPU",
     "Veamos los resultados."),
    (9, "Comparación justa: ocho modelos, misma prueba", ML, 1.0,
     f"Este es el resultado central: F1 macro sobre los mismos {entero(A['n_comun'])} comentarios. BETO logra {n(f['f1']['BETO'])}; "
     f"la mejor red, {n(max(f['f1'][k] for k in ['RNN', 'LSTM', 'GRU']))}; la TF-IDF con todos los datos, {n(f['f1'][BASE])}. Y fíjense "
     f"en la barra naranja: la TF-IDF entrenada con los mismos {entero(f['n_train_bert'])} comentarios que BETO llega a "
     f"{n(f['f1'][f['sub']])}. Esa distancia es lo que aporta el preentrenamiento con el mismo presupuesto de datos.",
     f"BETO {n(f['f1']['BETO'])} · LSTM {n(f['f1']['LSTM'])} · GRU {n(f['f1']['GRU'])} · TF-IDF {n(f['f1'][BASE])} · TF-IDF 16k {n(f['f1'][f['sub']])}",
     "¿Esa ventaja es real o es suerte?"),
    (10, "¿La ventaja es real?", ML, 1.0,
     f"Todos los modelos son buenos en negativo y positivo; la diferencia se juega en el neutro, donde BETO sube a "
     f"{n(f['f1_neu']['BETO'])}. Y no basta con ver una diferencia: remuestreamos la prueba mil veces. La ventaja de BETO sobre {S3} es "
     f"{n(ic[0])} con intervalo del 95% entre {n(ic[1])} y {n(ic[2])}: no cruza el cero, así que es estadísticamente significativa.",
     f"Δ BETO − {S3} = +{n(ic[0])}, IC 95% [{n(ic[1])}; {n(ic[2])}]",
     "Carlos revisa dónde se siguen equivocando."),
    (11, "Dónde se equivocan los modelos", DS, 1.0,
     f"Revisamos los errores después de congelar el modelo; no cambiamos nada con lo que vimos aquí. El 92% de los errores de BETO "
     f"está en la frontera con el neutro: comentarios de tres estrellas que suenan positivos o negativos, y cuya etiqueta sale de las "
     f"estrellas, no de leer el texto. Lo más importante para el negocio: los negativos leídos como positivos son "
     f"{s4['neg_como_pos']['BETO']} con BETO frente a {s4['neg_como_pos'][BASE]} con TF-IDF.",
     f"92% de errores en el neutro · error grave: BETO {s4['neg_como_pos']['BETO']} vs TF-IDF {s4['neg_como_pos'][BASE]}",
     "Las métricas globales no lo dicen todo: probamos casos difíciles."),
    (12, "Contrafactuales y sesgo de longitud", DS, 1.0,
     f"Probamos pares mínimos: con y sin negación, contraste invertido, errores de ortografía e ironía, en versión corta y en versión "
     f"larga. Con los textos cortos, BETO dice negativo casi siempre, incluso ante 'El producto es excelente, lo recomiendo'. La razón "
     f"está en los datos: el {pct(SES.get('pct_negativo_le15'), 0)} de los comentarios de 15 palabras o menos son negativos, porque los "
     f"autores del corpus prefirieron reseñas largas. BETO solo reconoce el {pct(rec_corto.get('BETO'), 0)} de los positivos cortos. "
     f"Con las versiones largas mejora: el elogio ya sale positivo, aunque el contraste largo todavía sale neutro. Y los textos que no son "
     f"reseñas ya no reciben etiqueta: el control de dominio los manda a revisión.",
     f"{pct(SES.get('pct_negativo_le15'), 0)} de los cortos son negativos · recall de positivos cortos: BETO {pct(rec_corto.get('BETO'), 0)} · "
     f"ajenos detectados: {pct(OODT.get('textos fuera de dominio detectados'), 0)} con {pct(OODT.get('reseñas de prueba marcadas'))} de falsas alarmas",
     "Fabiola explica cómo convertimos esto en una operación segura."),
    (13, "Si el modelo no está seguro, decide una persona", GIA, 1.0,
     f"Primero calibramos las probabilidades con temperature scaling: así, una confianza de 0,8 significa acertar cerca del 80%. "
     f"Luego definimos la abstención: si la confianza es menor que el umbral, el comentario pasa a una persona. El umbral, "
     f"{n(f['umbral']['BETO'], 2)}, se eligió en validación para lograr recall de negativo de 0,92 en lo que el modelo decide. En la "
     f"prueba, BETO decide el {pct(ab.get('cobertura'))} y manda a revisión unos {entero(1000 * (ab.get('abstencion') or 0))} de cada "
     "mil comentarios. Es un control y también un costo, porque alguien tiene que revisarlos.",
     f"Umbral {n(f['umbral']['BETO'], 2)} · cobertura {pct(ab.get('cobertura'))} · {entero(1000 * (ab.get('abstencion') or 0))} de cada 1.000 a revisión",
     "¿Cuánto cuesta operar esto?"),
    (14, "Calidad frente a costo", GIA, 0.75,
     f"Medimos todo en el mismo portátil, sin GPU. TF-IDF responde en un milisegundo; BETO, en {n(f['p95']['BETO'], 0)} en el percentil "
     f"95, y cuantizado a int8, en {n(f['p95']['BETO int8'], 0)}. Un millón de comentarios cuesta unos 6 dólares de cómputo con BETO. Con "
     f"supuestos que marcamos como ilustrativos, el beneficio neto frente al muestreo de hoy ronda los {entero(NETO / 1000)} mil dólares al "
     f"mes, y la implementación se paga en menos de un mes. Aunque la tasa de devoluciones evitadas fuera cinco veces menor, seguiría "
     f"siendo positivo.",
     f"p95: TF-IDF 1 ms · BETO {n(f['p95']['BETO'], 0)} ms · int8 {n(f['p95']['BETO int8'], 0)} ms · ≈ US$6 por millón · neto ≈ US${entero(NETO / 1000)} mil/mes",
     "Con calidad, riesgo y costo, aplicamos el criterio de adopción."),
    (15, "Criterio de adopción de la sesión 4", GIA, 1.0,
     "El criterio exige cuatro cosas a la vez: mejorar el F1 al menos 0,02 frente al campeón anterior, recall de negativo de 0,92 con "
     "abstención, cobertura de al menos 75% y latencia de 150 milisegundos o menos. Con el criterio estricto, ningún modelo cumple todo: "
     "BETO falla la latencia por una décima de milisegundo en un portátil sin GPU, y la versión int8 mejora solo 0,014. Como el proceso "
     "es por lotes horarios, y así lo fijamos desde la formulación con un quinto criterio, capacidad en lote, que BETO cumple con "
     "margen: procesa una hora de comentarios en menos de tres minutos. Por eso no proponemos desplegar, sino un piloto "
     "condicionado en modo sombra.",
     "BETO: ✔ calidad ✔ recall ✔ cobertura ✘ latencia (150,1 ms) · int8: ✘ calidad (+0,014)",
     "¿Qué puede salir mal y cómo lo controlamos?"),
    (16, "Riesgos, controles y suspensión", GIA, 1.0,
     "Cada riesgo tiene un control y un responsable; la matriz completa, con catorce riesgos, está en el Excel. Lo más importante: el "
     "modelo solo prioriza y alerta, nunca sanciona; lo que no parece una reseña va a revisión; la información personal se enmascara "
     "antes de procesar; reportamos el F1 por "
     "país; y tenemos condiciones explícitas para suspender y volver al modelo simple: recall de negativo por debajo de 0,88 dos "
     "semanas, abstención mayor a 35% o cualquier exposición de datos personales.",
     "14 riesgos · suspender si recall < 0,88, abstención > 35%, PII expuesta o precisión de alertas < 80%",
     "Ahora lo vemos funcionando."),
    (17, "Demostración del flujo completo", f"{ML} y {DS}", 2.0,
     "Machine Learning: Abrimos la demo web en el navegador. Escribimos una queja larga sobre un producto que dejó de funcionar a la "
     "semana. El sistema valida la entrada, enmascara correos y teléfonos, revisa que el texto sí sea una reseña, calcula las "
     "probabilidades calibradas de BETO y aplica el umbral: como está seguro y es negativo, genera la alerta y muestra las palabras que "
     "más pesaron. "
     "Ciencia de Datos: Ahora tres casos difíciles. Primero, 'Excelente producto, lo recomiendo': el modelo dice negativo por el sesgo de "
     "longitud, pero la salida trae la advertencia y no dispara la alerta. Segundo, el mismo elogio con más contexto: ahora sale positivo. "
     "Tercero, una frase sobre un partido de fútbol: el control de dominio la reconoce como ajena y la manda a revisión, en lugar de "
     "inventarle un sentimiento. "
     "[abrir https://sentimiento-melisa-demo.onrender.com (o python app/app.py y http://localhost:5000 para mostrar BETO); si falla, usar la captura de respaldo o python src/demo.py \"texto\"]",
     "Entrada → PII → ¿es reseña? → BETO + temperatura → umbral → etiqueta, alerta y evidencia",
     "Cerramos con la recomendación de los tres."),
    (18, "Recomendación y siguientes pasos", "Los tres roles", 1.0,
     "Machine Learning: BETO es el mejor modelo, con una ventaja significativa, y lo logró con veinte veces menos datos. "
     "Ciencia de Datos: la evaluación es válida, sin fuga y con la prueba usada una sola vez, pero el corpus tiene un sesgo de longitud "
     "que hay que corregir con datos cortos balanceados. "
     "Gerencia en IA: recomendamos un piloto de BETO en modo sombra, por lotes y con abstención; los comentarios cortos no disparan "
     "alertas y TF-IDF queda como respaldo. El siguiente experimento es BETO con GPU y la base completa. Gracias, quedamos atentos a "
     "sus preguntas.",
     "Piloto en modo sombra · 3 controles · siguiente experimento con GPU", ""),
]
assert abs(sum(g[3] for g in G) - 19.5) < 0.01, sum(g[3] for g in G)

PREGUNTAS = [
    ("¿Qué decisión mejora el sistema y quién la toma?", GIA,
     "El analista de experiencia del comprador decide qué comentario atender primero y a qué vendedor alertar. El modelo prioriza; la "
     "persona decide en los casos dudosos y confirma las alertas."),
    ("¿Qué evidencia demuestra que el modelo supera al baseline?", ML,
     f"Misma prueba de {entero(A['n_comun'])} comentarios: BETO {n(f['f1']['BETO'])} frente a {n(f['f1'][BASE])} de TF-IDF; frente a "
     f"{S3}, +{n(ic[0])} con IC 95% [{n(ic[1])}; {n(ic[2])}]. La selección se hizo en validación y la prueba se usó una vez."),
    ("¿Cuál es la clase o el error más crítico?", GIA,
     f"Negativo leído como positivo: se pierde la alerta. BETO comete {s4['neg_como_pos']['BETO']} en la prueba común frente a "
     f"{s4['neg_como_pos'][BASE]} de TF-IDF, y con abstención el recall de negativo es {n(ab.get('recall_neg_auto'))}."),
    ("¿Qué ocurre cuando el sistema no sabe o recibe texto fuera de dominio?", ML,
     "Si la confianza calibrada es menor que el umbral, se abstiene y el caso va a una persona; el texto vacío o demasiado corto se "
     f"rechaza antes del modelo. Además, un control compara el texto con reseñas reales usando la representación de BETO: detecta el "
     f"{pct(OODT.get('textos fuera de dominio detectados'), 0)} de 20 textos ajenos y solo marca el {pct(OODT.get('reseñas de prueba marcadas'))} "
     f"de las reseñas reales; lo detectado va a revisión."),
    ("¿Qué datos no debería procesar?", DS,
     "Datos personales (se enmascaran), texto en otros idiomas o de otros dominios sin una nueva evaluación, y nunca para sancionar "
     "personas de forma automática. La data card lista los usos no previstos."),
    ("¿Cómo se monitoreará el drift?", DS,
     "Semanalmente: tasa de abstención, distribución de confianza y de clases, subpalabras por palabra, latencia y F1 en una muestra "
     "reetiquetada mensual, por país y por longitud. Se reentrena cada trimestre o si la abstención sube 10 puntos."),
    ("¿Cuál es el costo aproximado de operar la solución?", GIA,
     "Cómputo: unos 6 dólares por millón de comentarios con BETO en CPU (centavos con TF-IDF). El costo mayor es humano: con 78% de "
     "cobertura, unos 217 de cada 1.000 comentarios van a revisión."),
    ("¿Qué condición obligaría a suspender o revertir el modelo?", GIA,
     "Recall de negativo menor que 0,88 dos semanas, abstención mayor que 35%, F1 de un país menor que 0,65, cualquier PII expuesta o "
     "precisión de alertas menor que 80%. La reversión es inmediata a TF-IDF + LR."),
    ("¿Por qué BETO no usó toda la base?", ML,
     "Restricción de cómputo: sin GPU, 16 mil comentarios tomaron 108 minutos. Para aislar el efecto, comparamos con una TF-IDF "
     f"entrenada con los mismos 16 mil ({n(f['f1'][f['sub']])}): la ventaja viene del preentrenamiento, no de más datos."),
    ("¿Por qué recomiendan BETO si no cumple el criterio?", GIA,
     "Cumple calidad, clase crítica y cobertura; falla la latencia por 0,1 ms en un portátil sin GPU, y el proceso es por lotes. Por eso "
     "la recomendación es condicionada y en modo sombra, no un despliegue; en tiempo real se usaría la versión int8 o una GPU."),
]

# ------------------------------------------------------------------ documento
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
sec.left_margin = sec.right_margin = Cm(2.2)
sec.top_margin = sec.bottom_margin = Cm(2)
st = doc.styles["Normal"]
st.font.name, st.font.size = "Calibri", Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.paragraph_format.space_after = Pt(5)
for nivel, tam in [(1, 16), (2, 13)]:
    h = doc.styles[f"Heading {nivel}"]
    h.font.name, h.font.size, h.font.bold, h.font.color.rgb = "Calibri", Pt(tam), True, RGBColor(0x0B, 0x2E, 0x59)
    h.paragraph_format.keep_with_next = True


def sombra(c, color):
    pr = c._tc.get_or_add_tcPr()
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear"), el.set(qn("w:color"), "auto"), el.set(qn("w:fill"), color)
    pr.append(el)


def tabla(filas, anchos, size=10, colores=None):
    t = doc.add_table(rows=len(filas), cols=len(filas[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, fila in enumerate(filas):
        for j, v in enumerate(fila):
            c = t.cell(i, j)
            c.width = Cm(anchos[j])
            c.text = ""
            run = c.paragraphs[0].add_run(str(v))
            run.font.size = Pt(size)
            run.bold = i == 0
            if i == 0:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                sombra(c, "0B2E59")
            elif colores and (i, j) in colores:
                sombra(c, colores[(i, j)])
    doc.add_paragraph()


t = doc.add_paragraph()
run = t.add_run("Guion de la sustentación por rol")
run.bold, run.font.size, run.font.color.rgb = True, Pt(22), RGBColor(0x0B, 0x2E, 0x59)
doc.add_paragraph("Sentimiento en reseñas de Mercado Libre (MeLiSA) · Proyecto final de Procesamiento de Lenguaje Natural · "
                  "Universidad Ean · Docente: Gerson José Moreno Morales")
doc.add_paragraph("Duración: 20 minutos de presentación (el guion suma 19,5 para dejar margen) más preguntas. Todos los integrantes hablan desde su rol y la narrativa es una "
                  "sola: Gerencia abre con el problema, Ciencia de Datos valida los datos, Machine Learning presenta los modelos, Ciencia "
                  "de Datos muestra los límites, Gerencia convierte la evidencia en una decisión gobernada y los tres cierran juntos.")

doc.add_heading("1. Reparto de roles y tiempos", 1)
mins = {}
for g in G:
    for rol in ([ML, DS] if g[2] == f"{ML} y {DS}" else [ML, DS, GIA] if g[2] == "Los tres roles" else [g[2]]):
        parte = g[3] / (2 if g[2] == f"{ML} y {DS}" else 3 if g[2] == "Los tres roles" else 1)
        mins[rol] = mins.get(rol, 0) + parte
diaps = {rol: ", ".join(str(g[0]) for g in G if rol in g[2] or g[2] == "Los tres roles") for rol in [GIA, DS, ML]}
tabla([["Integrante y rol", "Responsabilidad en la sustentación", "Diapositivas", "Minutos"],
       [f"{POR_ROL[GIA]} · {GIA}", "Problema, valor y KPI; abstención, costo, criterio de adopción, riesgos y recomendación", diaps[GIA], n(mins[GIA], 1)],
       [f"{POR_ROL[DS]} · {DS}", "Datos, auditoría, fuga y partición; errores, contrafactuales y sesgo de longitud; parte de la demo", diaps[DS], n(mins[DS], 1)],
       [f"{POR_ROL[ML]} · {ML}", "Arquitecturas, atención y BETO, fine-tuning, resultados e intervalos de confianza; demo", diaps[ML], n(mins[ML], 1)]],
      [4.2, 8.3, 2.6, 1.8], colores={(1, 0): "E2E8F0", (2, 0): "E0F7FA", (3, 0): "DBEAFE"})
tabla([["Bloque (instrucciones del curso)", "Tiempo", "Diapositivas", "Habla"],
       ["Problema, usuario y decisión", "2,75 min", "1–3", "Fabiola · Gerencia en IA"],
       ["Datos, calidad, privacidad y partición", "3 min", "4–5", "Carlos · Ciencia de Datos"],
       ["Baseline, modelo avanzado y diseño experimental", "5 min", "6–10", "Pedro · Machine Learning"],
       ["Resultados, errores, robustez y costo", "3,75 min", "11–14", "Carlos y Fabiola"],
       ["Riesgos, gobierno y revisión humana", "2 min", "15–16", "Fabiola · Gerencia en IA"],
       ["Demostración", "2 min", "17", "Pedro y Carlos"],
       ["Recomendación y siguiente paso", "1 min", "18", "Los tres"]], [6.6, 2, 2.4, 5.9])

doc.add_heading("2. Guion diapositiva por diapositiva", 1)
acum = 0.0
for num, titulo, rol, minutos, texto, cif, trans in G:
    ini = acum
    acum += minutos
    h = doc.add_heading(f"Diapositiva {num}. {titulo}", 2)
    p = doc.add_paragraph()
    rr = p.add_run(f"{quien(rol)}  ·  {n(minutos, 2).rstrip('0').rstrip(',')} min  ·  minuto {n(ini, 2).rstrip('0').rstrip(',') or '0'} a "
                   f"{n(acum, 2).rstrip('0').rstrip(',')}")
    rr.bold, rr.font.size = True, Pt(10)
    rr.font.color.rgb = RGBColor.from_string(COL.get(rol, "475569"))
    partes = texto.split("Ciencia de Datos: ") if rol == f"{ML} y {DS}" else [texto]
    for i, parte in enumerate(partes):
        doc.add_paragraph(("Ciencia de Datos: " if i else "") + parte.strip())
    p = doc.add_paragraph()
    p.add_run("Cifras que debe recordar: ").bold = True
    p.add_run(cif).italic = True
    if trans:
        p = doc.add_paragraph()
        p.add_run("Transición: ").bold = True
        p.add_run(f"“{trans}”")

doc.add_heading("3. Preparación de preguntas", 1)
doc.add_paragraph("Las ocho preguntas que el documento de instrucciones pide poder responder, más dos que probablemente hará el "
                  "jurado. Responde quien tiene el rol indicado; los demás complementan solo si hace falta.")
tabla([["Pregunta", "Responde", "Respuesta corta"]] + [[q, rol.replace("Especialista en ", "").replace("Magíster en ", ""), a]
                                                      for q, rol, a in PREGUNTAS], [4.6, 2.6, 9.7], size=9.5)

doc.add_heading("4. Recomendaciones para la sustentación", 1)
for t_ in ["Cada persona dice su rol al tomar la palabra (“Desde Ciencia de Datos…”). Usen las transiciones para no fragmentar la narrativa.",
           "Ensayen con cronómetro: el bloque de Machine Learning (5 min) es el que más tiende a alargarse; el guion deja 30 segundos de margen.",
           "Tengan la demo abierta antes de empezar: https://sentimiento-melisa-demo.onrender.com (el plan gratuito se duerme y la primera visita tarda unos 50 segundos) o, para mostrar BETO, python app/app.py en local. Tengan también la página docs/Demo_Respaldo.html por si algo falla.",
           "No lean las diapositivas: las cifras clave de cada una están en este guion y en las notas del orador del PPT.",
           "Si preguntan por algo que no se ejecutó, díganlo así: está en “siguiente experimento”."]:
    doc.add_paragraph(t_, style="List Bullet")

os.makedirs(os.path.join(RAIZ, "docs"), exist_ok=True)
salida = os.path.join(RAIZ, "docs", "Guion_Sustentacion_por_Rol.docx")
doc.save(salida)
print("Guion ->", salida, {k: round(v, 2) for k, v in mins.items()})

# ------------------------------------------------------------------ notas del orador en la presentación
ruta_ppt = os.path.join(RAIZ, "docs", "Presentacion_Proyecto_PLN.pptx")
prs = Presentation(ruta_ppt)
assert len(prs.slides) == len(G)
for sl, (num, titulo, rol, minutos, texto, cif, trans) in zip(prs.slides, G):
    nota = f"Habla: {quien(rol)} · {n(minutos, 2).rstrip('0').rstrip(',')} min\n\n{texto}\n\nCifras clave: {cif}"
    if trans:
        nota += f"\n\nTransición: “{trans}”"
    sl.notes_slide.notes_text_frame.text = nota
prs.save(ruta_ppt)
print("Notas del orador actualizadas en", ruta_ppt)
