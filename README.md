# Sentimiento en reseñas de Mercado Libre (MeLiSA) · Proyecto final PLN

**Universidad Ean · Procesamiento de Lenguaje Natural** · Docente: Gerson José Moreno Morales · Estudiante: Jose Romero

¿Un **Transformer preentrenado en español (BETO)** o un **modelo secuencial (RNN/LSTM/GRU)** justifica su mayor costo frente a una
línea base **TF-IDF + Regresión Logística / SVM** para clasificar el sentimiento (negativo, neutro, positivo) de los comentarios de
compradores de Mercado Libre, y bajo qué controles se podría adoptar?

| | |
|---|---|
| **Usuario** | Analista de experiencia del comprador (Atención al Cliente / Reputación de vendedores) |
| **Decisión** | Qué comentario atender primero y a qué vendedor enviar una alerta temprana |
| **Entrada → salida** | Un comentario en español → etiqueta + probabilidad calibrada + decisión automática o revisión humana + evidencia |
| **Clase crítica** | Negativo (error más costoso: negativo leído como positivo) |
| **KPI** | Primera respuesta a negativos < 24 h · % de alertas confirmadas · reclamos/devoluciones evitados |

## Resultados

Todos los modelos usan la misma partición (70/15/15, semilla 42). La comparación final se hace sobre una **prueba común de 5.000
comentarios**; los modelos de la sesión 3 también se reportan sobre la prueba completa (72.994) en el notebook.

| Modelo | Datos de entrenamiento | F1 macro (prueba común) | F1 neutro | Recall negativo | p95 por comentario (CPU) |
|---|---:|---:|---:|---:|---:|
| **BETO** (fine-tuning, sesión 4) | 16.000 | **0,793** | **0,624** | 0,854 | 150,1 ms |
| BETO int8 (cuantizado) | 16.000 | 0,773 | 0,592 | 0,876 | 121,3 ms |
| LSTM (sesión 3) | 340.634 | 0,766 | 0,578 | 0,833 | 4,9 ms |
| GRU (campeón recurrente por validación) | 340.634 | 0,758 | 0,584 | 0,824 | 6,5 ms |
| RNN | 340.634 | 0,753 | 0,569 | 0,798 | 4,4 ms |
| TF-IDF + Regresión Logística (línea base) | 340.634 | 0,742 | 0,556 | 0,798 | 1,1 ms |
| TF-IDF + SVM lineal | 340.634 | 0,721 | 0,481 | 0,851 | 1,1 ms |
| TF-IDF + LR con el mismo subconjunto que BETO | 16.000 | 0,704 | 0,493 | 0,785 | 1,4 ms |

- **BETO − GRU = +0,035 de F1 macro, IC 95% [0,024; 0,045]**: la diferencia es significativa (bootstrap de 1.000 remuestreos).
- Con calibración por temperatura y un umbral de abstención de 0,69, **BETO decide solo el 78,3%** de los comentarios con **recall de
  negativo 0,922**; el resto pasa a revisión humana.
- **Hallazgo de datos: sesgo de longitud.** El 93,8% de los comentarios de 15 palabras o menos son negativos. Los modelos aprendieron
  "corto ⇒ negativo": BETO recupera solo el 9,5% de los positivos cortos.
- **Criterio de la sesión 4:** ningún modelo cumple los cuatro requisitos a la vez. BETO falla la latencia por 0,1 ms en un portátil
  sin GPU y BETO int8 mejora solo +0,014.
- **Decisión del comité:** piloto de BETO en **modo sombra**, por lotes horarios y con abstención. Los comentarios cortos no disparan
  alertas automáticas al vendedor y TF-IDF + LR queda como respaldo. Detalle en `resultados/decision.json` y en la sección 9 del notebook.

## Estructura del repositorio
```
├── Proyecto_PLN_Sentimiento_MeLiSA.ipynb   Notebook principal, ejecutado de inicio a fin (EDA → baseline → RNN → BETO → evaluación → demo)
├── data/
│   ├── README.md                             Data card, diccionario de datos, origen, licencia y limitaciones
│   └── muestra_melisa_estratificada.csv      Muestra controlada (5.001 reseñas, PII enmascarada)
├── src/
│   ├── demo.py                               Demostración del flujo completo desde la terminal
│   ├── crear_muestra.py                      Genera la muestra controlada desde la base completa
│   ├── generar_presentacion.py               docs/Presentacion_Proyecto_PLN.pptx desde los resultados
│   ├── generar_informe.py                    docs/Informe_Tecnico_Proyecto_PLN.docx desde los resultados
│   ├── generar_matriz_riesgos.py             docs/Matriz_Riesgos.xlsx desde los resultados
│   └── resumen.py                            Lectura de resultados compartida
├── resultados/                               Métricas (resultados.json, tablas CSV) y figuras del notebook
├── docs/                                     Entregables: presentación, informe técnico y matriz de riesgos
├── modelos/                                  Artefactos entrenados (no se versionan: se regeneran al ejecutar el notebook)
├── requirements.txt · .env.example · LICENSE
```

## Entregables del proyecto final
| N.º | Entregable | Dónde |
|---|---|---|
| 1 | Repositorio con código, README, dependencias y licencia | Este repositorio · `requirements.txt` · `LICENSE` |
| 2 | Notebook principal ejecutado de inicio a fin | [`Proyecto_PLN_Sentimiento_MeLiSA.ipynb`](Proyecto_PLN_Sentimiento_MeLiSA.ipynb) |
| 3 | Dataset o muestra controlada + diccionario y origen | [`data/`](data/) |
| 4 | Informe técnico (DOCX) | [`docs/Informe_Tecnico_Proyecto_PLN.docx`](docs/Informe_Tecnico_Proyecto_PLN.docx) |
| 5 | Presentación final (PPTX, 18 diapositivas con notas por rol) | [`docs/Presentacion_Proyecto_PLN.pptx`](docs/Presentacion_Proyecto_PLN.pptx) |
| 6 | Matriz de riesgos (XLSX) | [`docs/Matriz_Riesgos.xlsx`](docs/Matriz_Riesgos.xlsx) |
| 7 | Demostración | `python src/demo.py "texto"` y sección 10 del notebook |

## Cómo reproducirlo
Requisitos: Python 3.11. Probado en Windows 11 con CPU (12 hilos, 16 GB de RAM, sin GPU).

```bash
python -m venv .venv
.venv\Scripts\activate          # Linux/Mac: source .venv/bin/activate
pip install -r requirements.txt
```

1. **Datos.** Descargue MeLiSA (Hugging Face: `lpsc-fiuba/melisa`), conserve los sitios en español (todos menos `MLB`, Brasil),
   agregue la columna `sentimiento` (1–2★ negativo, 3★ neutro, 4–5★ positivo) y guárdela como `data/data_melisa_comentarios.csv`
   con las columnas `id, country, category, review_content, review_title, review_rate, sentimiento` (136 MB; no se sube por el
   límite de GitHub).
   Sin ese archivo el notebook usa automáticamente `data/muestra_melisa_estratificada.csv`.
2. **Notebook.** Ejecútelo de inicio a fin (Jupyter, VS Code o `jupyter nbconvert --to notebook --execute --inplace
   Proyecto_PLN_Sentimiento_MeLiSA.ipynb`). Los parámetros se cambian con variables de entorno (ver `.env.example`); por ejemplo,
   `FRACCION=0.05` hace una corrida rápida. La primera ejecución descarga BETO (~440 MB) de Hugging Face.
3. **Demo.** `python src/demo.py "El cargador dejó de funcionar a la semana"` o `python src/demo.py` para el modo interactivo.
4. **Documentos.** `python src/generar_presentacion.py`, `python src/generar_informe.py` y `python src/generar_matriz_riesgos.py`
   regeneran los entregables de `docs/` a partir de `resultados/resultados.json`.

Semilla 42 en todo el flujo. La prueba se usa una sola vez; C, épocas, temperatura, umbral y modelo campeón se eligen en validación.

## Datos
MeLiSA (*Mercado Libre for Sentiment Analysis*), LPSC-FIUBA, Universidad de Buenos Aires. **Datos reales y públicos**, licencia
"other", usados con fines académicos. Detalle completo, auditoría, PII, fuga y usos no previstos en [`data/README.md`](data/README.md).

## Licencia
Código bajo licencia MIT. Los datos (MeLiSA) y el modelo preentrenado (BETO, CC BY 4.0) conservan sus licencias originales; ver
[`LICENSE`](LICENSE).

## Declaración de uso de IA generativa
Se usó un asistente de IA (Claude, de Anthropic) para estructurar el código del notebook, redactar borradores de documentación y
escribir los scripts que generan la presentación, el informe y la matriz de riesgos. Todas las cifras provienen de la ejecución del
notebook (`resultados/resultados.json`) y fueron verificadas por el equipo; no se presenta como evidencia nada que no se haya ejecutado.
No se enviaron datos confidenciales a servicios externos (el corpus es público). El equipo mantiene la responsabilidad y la autoría de las
decisiones de diseño, la interpretación y la recomendación.
