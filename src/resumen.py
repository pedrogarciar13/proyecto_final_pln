"""Lectura de resultados/resultados.json y cifras clave compartidas por el informe, la presentación y la matriz de riesgos."""
import json
import os


def cargar(raiz="."):
    r = json.load(open(os.path.join(raiz, "resultados", "resultados.json"), encoding="utf-8"))
    r["_raiz"] = raiz
    return r


def n(x, dec=3):
    """Número con coma decimal (formato es-CO)."""
    if x is None:
        return "—"
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x, dec=1):
    return "—" if x is None else n(100 * x, dec) + "%"


def entero(x):
    return f"{int(round(x)):,}".replace(",", ".")


def corto(nombre):
    return (nombre.replace("TF-IDF + Regresión Logística", "TF-IDF + LR").replace("TF-IDF + SVM lineal", "TF-IDF + SVM")
            .replace(", mismo subconjunto", ""))


def cifras(r):
    """Diccionario con las cifras que se citan en todos los entregables."""
    s4 = r["s4"]
    base, s3, sub = s4["base"], s4["campeon_s3"], s4["nombre_sub"]
    c = s4["comun"]
    ab = s4["abstencion"]
    variantes = [k for k in ["BETO", "BETO int8"] if k in c]
    mejor_bert = max(variantes, key=lambda k: c[k]["val"]["f1_macro"])
    rec = s4["modelo_demo"]
    f = {
        "base": base, "s3": s3, "sub": sub, "bert": mejor_bert, "rec": rec, "variantes": variantes,
        "f1": {k: v["test"]["f1_macro"] for k, v in c.items()},
        "f1_val": {k: v["val"]["f1_macro"] for k, v in c.items()},
        "rec_neg": {k: v["test"]["recall_negativo"] for k, v in c.items()},
        "f1_neu": {k: v["test"]["f1_neutro"] for k, v in c.items()},
        "f1_full": {k: v["test"]["f1_macro"] for k, v in r["resultados_s3"].items()},
        "p95": {k: v["p95"] for k, v in s4["info"].items()},
        "thr": {k: v["thr"] for k, v in s4["info"].items()},
        "seg": {k: v["seg_entreno"] for k, v in s4["info"].items()},
        "params": {k: v["params"] for k, v in s4["info"].items()},
        "abst": {k: (v or {}).get("test") for k, v in ab.items()},
        "umbral": {k: (v or {}).get("umbral") for k, v in ab.items()},
        "ic": s4["ic_bootstrap"], "pasa": s4["pasa"], "cumple_bert": s4["cumple_bert"],
        "n_train_bert": s4["n_train_bert"], "seg_bert": s4["seg_bert"], "max_len": s4["max_len"],
        "neg_pos": s4["neg_como_pos"], "audit": r["audit"], "entorno": r["entorno"],
    }
    f["decision"] = None
    if os.path.exists(os.path.join(r.get("_raiz", "."), "resultados", "decision.json")):
        d = json.load(open(os.path.join(r.get("_raiz", "."), "resultados", "decision.json"), encoding="utf-8"))
        f["decision"], f["rec"] = d, d["modelo"]
    f["sesgo"] = r["audit"].get("sesgo_longitud")
    f["delta_bert_s3"] = f["f1"][mejor_bert] - f["f1"][s3]
    f["delta_bert_base"] = f["f1"][mejor_bert] - f["f1"][base]
    f["delta_s3_base"] = f["f1"][s3] - f["f1"][base]
    return f
