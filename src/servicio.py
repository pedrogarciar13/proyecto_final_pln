"""Lógica de predicción compartida por la demo de terminal (src/demo.py) y la demo web (app/app.py).

Flujo: validar → enmascarar PII → detectar fuera de dominio → probabilidades calibradas → umbral de abstención →
regla de comentarios cortos → etiqueta, decisión, alerta y evidencia.
"""
import json
import os
import re

os.environ.setdefault("TF_CPP_MIN_LOG_LEVEL", "2")
os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")
os.environ.setdefault("TRANSFORMERS_VERBOSITY", "error")
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")
import joblib
import numpy as np

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLASES = ["negativo", "neutro", "positivo"]
PII = {"EMAIL": r"[\w.+-]+@[\w-]+\.[\w.]+",
       "URL": r"(?:https?://|www\.)\S+",
       "TEL": r"(?<!\d)\+?\d[\d\s\-]{7,}\d(?!\d)"}
PALABRAS_CORTO = 15
TOKEN = re.compile(r"(?u)\b\w\w+\b")


def enmascarar(s):
    for k, p in PII.items():
        s = re.sub(p, f"<{k}>", s)
    return s


def softmax(z, T=1.0):
    z = np.asarray(z, dtype=float) / T
    e = np.exp(z - z.max(-1, keepdims=True))
    return e / e.sum(-1, keepdims=True)


def _buscar(*rutas):
    for r in rutas:
        if os.path.exists(r):
            return r
    return None


class Servicio:
    """Carga el modelo elegido y aplica toda la política de decisión."""

    def __init__(self, modelo=None):
        res = json.load(open(os.path.join(RAIZ, "resultados", "resultados.json"), encoding="utf-8"))["s4"]
        ruta_dec = os.path.join(RAIZ, "resultados", "decision.json")
        recomendado = json.load(open(ruta_dec, encoding="utf-8"))["modelo"] if os.path.exists(ruta_dec) else res["modelo_demo"]
        beto_ok = os.path.isdir(os.path.join(RAIZ, "modelos", "beto_sentimiento"))
        pedido = modelo or os.environ.get("MODELO_DEMO") or recomendado
        if pedido.startswith("BETO") and not beto_ok:   # sin artefacto de BETO (por ejemplo en Render) se usa el respaldo
            pedido = "TF-IDF + Regresión Logística"
        self.nombre = pedido
        self.umbral = (res["abstencion"].get(pedido) or {}).get("umbral") or 1.0
        self.T = (res["calibracion"].get(pedido) or {}).get("T", 1.0)
        self.max_len = res.get("max_len", 128)
        ruta_lr = _buscar(os.path.join(RAIZ, "modelos", "tfidf_lr.joblib"), os.path.join(RAIZ, "app", "modelo", "tfidf_lr.joblib"))
        self.lr = joblib.load(ruta_lr)
        self.tfidf = self.lr.named_steps["tfidf"]
        self.vocab = self.tfidf.get_feature_names_out()
        self.vocab_set = set(self.tfidf.vocabulary_)
        self.clases_lr = list(self.lr.named_steps["clf"].classes_)
        ruta_ood = _buscar(os.path.join(RAIZ, "resultados", "ood_referencia.json"))
        self.ood_cfg = json.load(open(ruta_ood, encoding="utf-8")) if ruta_ood else None
        self._emb = None
        self.predictor = self._cargar(pedido)

    # ------------------------------------------------------------------ modelos
    def _cargar(self, nombre):
        if nombre.startswith("TF-IDF"):
            orden = [self.clases_lr.index(c) for c in CLASES]
            return lambda t: self.lr.predict_proba(list(t))[:, orden]
        if nombre in ("RNN", "LSTM", "GRU"):
            import tensorflow as tf
            vocab = json.load(open(os.path.join(RAIZ, "modelos", "vocabulario_rnn.json"), encoding="utf-8"))
            vec = tf.keras.layers.TextVectorization(output_sequence_length=100, standardize="lower_and_strip_punctuation",
                                                    vocabulary=vocab[2:])
            m = tf.keras.models.load_model(os.path.join(RAIZ, "modelos", f"{nombre.lower()}.keras"))
            return lambda t: m.predict(vec(tf.constant(list(t))), verbose=0)
        import torch
        from transformers import AutoModelForSequenceClassification, AutoTokenizer
        ruta = os.path.join(RAIZ, "modelos", "beto_sentimiento")
        self.tok = AutoTokenizer.from_pretrained(ruta)
        m = AutoModelForSequenceClassification.from_pretrained(ruta).eval()
        if nombre == "BETO int8":
            m = torch.ao.quantization.quantize_dynamic(m, {torch.nn.Linear}, dtype=torch.qint8)
        self.bert = m

        @torch.no_grad()
        def fn(t):
            b = self.tok(list(t), truncation=True, max_length=self.max_len, padding=True, return_tensors="pt")
            out = m(**b, output_hidden_states=True)
            h, mask = out.hidden_states[-1], b["attention_mask"].unsqueeze(-1)
            v = torch.nn.functional.normalize((h * mask).sum(1) / mask.sum(1), dim=1)
            self._emb = v.numpy()
            return softmax(out.logits.numpy())
        return fn

    # ------------------------------------------------------------------ fuera de dominio
    def fuera_de_dominio(self, texto):
        """Devuelve (es_ood, puntaje, método). Con BETO: similitud media con los 5 vecinos más cercanos de un conjunto de reseñas
        de referencia; con TF-IDF: proporción de palabras dentro del vocabulario de entrenamiento."""
        cfg = self.ood_cfg or {}
        if self._emb is not None and cfg.get("emb") is not None and hasattr(self, "bert"):
            if not hasattr(self, "_ref"):
                self._ref = np.asarray(np.load(os.path.join(RAIZ, cfg["emb"])), dtype=np.float32)
            s = float(np.sort(self._emb @ self._ref.T, axis=1)[:, -5:].mean())
            return s < cfg["umbral_emb"], s, "similitud semántica con reseñas"
        palabras = TOKEN.findall(texto.lower())
        cob = float(np.mean([w in self.vocab_set for w in palabras])) if palabras else 0.0
        return cob < cfg.get("umbral_cobertura", 0.85), cob, "cobertura de vocabulario"

    # ------------------------------------------------------------------ predicción
    def predecir(self, texto):
        if not isinstance(texto, str) or len(texto.strip()) < 3:
            return {"estado": "error", "detalle": "Texto vacío o demasiado corto.", "decision": "revisión humana"}
        limpio = enmascarar(texto.strip())
        P = softmax(np.log(np.clip(self.predictor([limpio])[0], 1e-9, 1)), self.T)
        clase, conf = CLASES[int(P.argmax())], float(P.max())
        ood, puntaje, metodo = self.fuera_de_dominio(limpio)
        n_pal = len(limpio.split())
        corto = n_pal <= PALABRAS_CORTO
        x = self.tfidf.transform([limpio])
        contrib = x.multiply(self.lr.named_steps["clf"].coef_[self.clases_lr.index(clase)]).toarray()[0]
        evidencia = [self.vocab[i] for i in np.argsort(contrib)[::-1][:4] if contrib[i] > 0]
        automatica = conf >= self.umbral and not ood
        advertencias = []
        if ood:
            advertencias.append(f"Posible texto fuera de dominio ({metodo}: {puntaje:.2f}); no parece una opinión sobre una compra.")
        if corto:
            advertencias.append(f"Comentario de {n_pal} palabras: en el corpus de entrenamiento el 94% de los comentarios de 15 palabras "
                                "o menos son negativos, así que la etiqueta puede estar sesgada. No genera alerta automática.")
        if conf < self.umbral:
            advertencias.append(f"Confianza {conf:.2f} menor que el umbral {self.umbral:.2f}: lo revisa una persona.")
        return {"estado": "ok", "modelo": self.nombre, "texto": limpio, "pii_enmascarada": limpio != texto.strip(),
                "etiqueta": clase, "confianza": round(conf, 3), "probabilidades": {c: round(float(p), 3) for c, p in zip(CLASES, P)},
                "umbral": round(self.umbral, 2), "decision": "automática" if automatica else "revisión humana",
                "alerta_vendedor": bool(clase == "negativo" and automatica and not corto),
                "fuera_de_dominio": bool(ood), "palabras": n_pal, "evidencia": evidencia, "advertencias": advertencias}
