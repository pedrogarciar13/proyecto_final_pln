"""Crea la muestra controlada del repositorio a partir de la base completa de MeLiSA.

- Estratificada por sentimiento y país (semilla 42).
- Sin duplicados ni textos contradictorios.
- PII enmascarada (correos, URL y teléfonos) en el comentario y el título.

Uso:  python src/crear_muestra.py [--n 5000] [--entrada data/data_melisa_comentarios.csv]
"""
import argparse
import re

import pandas as pd

PII = {"EMAIL": r"[\w.+-]+@[\w-]+\.[\w.]+",
       "URL": r"(?:https?://|www\.)\S+",
       "TEL": r"(?<!\d)\+?\d[\d\s\-]{7,}\d(?!\d)"}


def enmascarar(s: str) -> str:
    for k, p in PII.items():
        s = re.sub(p, f"<{k}>", s)
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entrada", default="data/data_melisa_comentarios.csv")
    ap.add_argument("--salida", default="data/muestra_melisa_estratificada.csv")
    ap.add_argument("--n", type=int, default=5000)
    ap.add_argument("--semilla", type=int, default=42)
    a = ap.parse_args()

    df = pd.read_csv(a.entrada)
    contradic = df.groupby("review_content")["sentimiento"].nunique()
    df = df[~df.review_content.isin(contradic[contradic > 1].index)].drop_duplicates("review_content")
    frac = a.n / len(df)
    muestra = (df.groupby(["sentimiento", "country"], group_keys=False)
                 .sample(frac=frac, random_state=a.semilla)
                 .sort_values("id"))
    for col in ["review_content", "review_title"]:
        muestra[col] = muestra[col].astype(str).map(enmascarar)
    muestra.to_csv(a.salida, index=False, encoding="utf-8")
    print(f"{len(muestra)} filas -> {a.salida}")
    print(muestra.sentimiento.value_counts(normalize=True).round(3).to_dict())


if __name__ == "__main__":
    main()
