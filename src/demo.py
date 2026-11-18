"""Demostración del flujo completo desde la terminal.

Requiere haber ejecutado el notebook principal (genera la carpeta modelos/).

Uso:
    python src/demo.py "El producto llegó roto y nadie responde"
    python src/demo.py            # modo interactivo (línea vacía para salir)
    python src/demo.py --modelo "TF-IDF + Regresión Logística" "Excelente, lo recomiendo"

Salida: etiqueta, probabilidades calibradas, confianza, decisión (automática / revisión humana),
alerta al vendedor y evidencia (términos que más empujaron la decisión según la regresión logística).
"""
import argparse
import json
import os
import re
import sys

os.environ.setdefault("TF_CPP_MIN_LOG_LEVEL", "2")
os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")
import joblib
import numpy as np

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODELOS = os.path.join(RAIZ, "modelos")
CLASES = ["negativo", "neutro", "positivo"]
PII = {"EMAIL": r"[\w.+-]+@[\w-]+\.[\w.]+",
       "URL": r"(?:https?://|www\.)\S+",
       "TEL": r"(?<!\d)\+?\d[\d\s\-]{7,}\d(?!\d)"}


def enmascarar(s):
    for k, p in PII.items():
        s = re.sub(p, f"<{k}>", s)
    return s


def softmax(z, T=1.0):
    z = np.asarray(z, dtype=float) / T
    e = np.exp(z - z.max(-1, keepdims=True))
    return e / e.sum(-1, keepdims=True)


def cargar_predictor(nombre, max_len):
    if nombre.startswith("TF-IDF"):
        archivo = "tfidf_lr.joblib" if "Log" in nombre else "tfidf_svm.joblib"
        p = joblib.load(os.path.join(MODELOS, archivo))
        orden = [list(p.classes_).index(c) for c in CLASES]
        return lambda t: p.predict_proba(t)[:, orden]
    if nombre in ("RNN", "LSTM", "GRU"):
        import tensorflow as tf
        vocab = json.load(open(os.path.join(MODELOS, "vocabulario_rnn.json"), encoding="utf-8"))
        vec = tf.keras.layers.TextVectorization(output_sequence_length=100, standardize="lower_and_strip_punctuation",
                                                vocabulary=vocab[2:])
        m = tf.keras.models.load_model(os.path.join(MODELOS, f"{nombre.lower()}.keras"))
        return lambda t: m.predict(vec(tf.constant(t)), verbose=0)
    import torch
    from transformers import AutoModelForSequenceClassification, AutoTokenizer
    ruta = os.path.join(MODELOS, "beto_sentimiento")
    tok = AutoTokenizer.from_pretrained(ruta)
    m = AutoModelForSequenceClassification.from_pretrained(ruta).eval()
    if nombre == "BETO int8":
        m = torch.ao.quantization.quantize_dynamic(m, {torch.nn.Linear}, dtype=torch.qint8)

    @torch.no_grad()
    def fn(t):
        b = tok(list(t), truncation=True, max_length=max_len, padding=True, return_tensors="pt")
        return softmax(m(**b).logits.numpy())
    return fn


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("texto", nargs="*")
    ap.add_argument("--modelo", default=None, help="Por defecto, el recomendado por el notebook")
    ap.add_argument("--guardar", action="store_true", help="Guarda la salida en resultados/demo_ejemplo.json")
    a = ap.parse_args()

    cfg = json.load(open(os.path.join(MODELOS, "config_demo.json"), encoding="utf-8"))
    res = json.load(open(os.path.join(RAIZ, "resultados", "resultados.json"), encoding="utf-8"))["s4"]
    ruta_dec = os.path.join(RAIZ, "resultados", "decision.json")   # decisión del comité (modelo recomendado)
    recomendado = json.load(open(ruta_dec, encoding="utf-8"))["modelo"] if os.path.exists(ruta_dec) else cfg["modelo"]
    nombre = a.modelo or recomendado
    umbral = (res["abstencion"].get(nombre) or {}).get("umbral") or cfg["umbral"] or 1.0
    T = (res["calibracion"].get(nombre) or {}).get("T", 1.0)
    predictor = cargar_predictor(nombre, cfg["max_len"])
    lr = joblib.load(os.path.join(MODELOS, "tfidf_lr.joblib"))
    vocab = lr.named_steps["tfidf"].get_feature_names_out()
    clases_lr = list(lr.named_steps["clf"].classes_)

    def predecir(texto):
        if not texto or len(texto.strip()) < 3:
            return {"estado": "error", "detalle": "texto vacío o demasiado corto", "decision": "revisión humana"}
        limpio = enmascarar(texto.strip())
        P = softmax(np.log(np.clip(predictor([limpio])[0], 1e-9, 1)), T)
        clase, conf = CLASES[int(P.argmax())], float(P.max())
        x = lr.named_steps["tfidf"].transform([limpio])
        contrib = x.multiply(lr.named_steps["clf"].coef_[clases_lr.index(clase)]).toarray()[0]
        evid = [vocab[i] for i in np.argsort(contrib)[::-1][:4] if contrib[i] > 0]
        corto = len(limpio.split()) <= 15
        salida = {"modelo": nombre, "etiqueta": clase, "confianza": round(conf, 3),
                  "probabilidades": {c: round(float(p), 3) for c, p in zip(CLASES, P)},
                  "decision": "automática" if conf >= umbral else "revisión humana", "umbral": round(umbral, 2),
                  # control del sesgo de longitud: los comentarios cortos no disparan alertas automáticas
                  "alerta_vendedor": clase == "negativo" and conf >= umbral and not corto, "evidencia": evid}
        if corto:
            salida["advertencia"] = ("comentario de 15 palabras o menos: en el corpus de entrenamiento el 94% de estos es negativo; "
                                     "la etiqueta puede estar sesgada y no genera alerta automática")
        return salida

    if a.texto:
        entrada = " ".join(a.texto)
        salida = predecir(entrada)
        print(json.dumps(salida, ensure_ascii=False, indent=1))
        if a.guardar:   # ejemplo que usa la presentación (diapositiva de demostración)
            json.dump({"entrada": entrada, **salida}, open(os.path.join(RAIZ, "resultados", "demo_ejemplo.json"), "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
        return
    print(f"Modelo: {nombre} · umbral de abstención {umbral:.2f} · escriba un comentario (línea vacía para salir)")
    for linea in sys.stdin:
        if not linea.strip():
            break
        print(json.dumps(predecir(linea), ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
