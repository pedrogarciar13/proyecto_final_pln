---
title: Sentimiento MeLiSA con BETO
emoji: 🛒
colorFrom: blue
colorTo: green
sdk: static
app_file: index.html
pinned: false
license: mit
short_description: Sentimiento en reseñas de Mercado Libre con BETO
---

# Demo con BETO en el navegador

Página estática que corre nuestro BETO ajustado directamente en el navegador del visitante, sin servidor. Está publicada en
https://huggingface.co/spaces/pedrogarciar13/sentimiento-melisa-beto.

- El modelo es BETO comprimido a int8 en formato ONNX (110 MB), descargado de
  [pedrogarciar13/beto-sentimiento-melisa](https://huggingface.co/pedrogarciar13/beto-sentimiento-melisa). En una muestra de 1.000
  reseñas coincide con el modelo original en el 98% de las etiquetas (F1 0,799 frente a 0,806) y detecta los 20 textos fuera de
  dominio de la prueba del notebook.
- Aplica la misma política que `src/servicio.py`: enmascarar datos personales, control de fuera de dominio por similitud con 3.000
  reseñas (`ood_ref_f16.bin`), calibración por temperatura y umbral de abstención (`config.json`), regla de comentarios cortos y
  evidencia con los términos de TF-IDF + regresión logística (`evidencia.json`).
- Para probarla en local: `python -m http.server 8765 -d demo_navegador` y abrir http://localhost:8765.

Proyecto final de PLN, Universidad Ean: Pedro García Romero, Carlos Trilleras y Fabiola Aldana Mora.
