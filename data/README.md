# Data card · MeLiSA (reseñas de Mercado Libre en español)

## Origen y licencia
| Campo | Detalle |
|---|---|
| Nombre | MeLiSA: *Mercado Libre for Sentiment Analysis* |
| Autor | Laboratorio de Procesamiento de Señales y Comunicaciones (LPSC), FIUBA, Universidad de Buenos Aires |
| Publicación | Hugging Face: [`lpsc-fiuba/melisa`](https://huggingface.co/datasets/lpsc-fiuba/melisa) |
| Licencia | "other" según la tarjeta del dataset. Uso **académico y no comercial** en este proyecto. |
| Tipo de dato | **Real y público**: reseñas que los compradores publicaron en Mercado Libre. No es sintético. |
| Periodo | Reseñas recolectadas entre **agosto de 2020 y enero de 2021** (tarjeta del dataset). El archivo no trae fecha por reseña. |
| Selección | Los autores priorizaron reseñas según su **longitud** y su valoración (likes menos dislikes) y equilibraron las estrellas (~20% cada una). Ese filtro explica el sesgo de longitud que encontramos. |
| Idioma | Español (este archivo solo contiene los 7 sitios hispanohablantes; la versión original también incluye portugués de Brasil). |
| Permisos | Datos públicos publicados por los autores para investigación; no se recolectaron datos nuevos. |

## Archivos de esta carpeta
| Archivo | En el repositorio | Descripción |
|---|---|---|
| `data_melisa_comentarios.csv` | **No** (136 MB, supera el límite de 100 MB de GitHub) | Base completa: 487.368 reseñas. Descárguela de Hugging Face (versión en español, 3 clases) y guárdela con ese nombre. |
| `muestra_melisa_estratificada.csv` | Sí (1,4 MB) | **Muestra controlada** de 5.001 reseñas, estratificada por sentimiento y país, sin duplicados, con PII enmascarada. Generada con `src/crear_muestra.py` (semilla 42). El notebook la usa automáticamente si no encuentra la base completa. |

## Diccionario de datos
| Columna | Tipo | Ejemplo | Uso en el proyecto |
|---|---|---|---|
| `id` | entero | 1 | Identificador de fila. No se usa para modelar. |
| `country` | texto (7 valores) | `MLA` | **Solo para evaluar por segmento** (sesgo por país). MLA = Argentina, MLM = México, MCO = Colombia, MLC = Chile, MLU = Uruguay, MLV = Venezuela, MPE = Perú. |
| `category` | texto (5 valores) | `Hogar / Casa` | Solo para evaluar por segmento. |
| `review_content` | texto libre | "Esta de muy baja calidad, no fue lo esperado." | **Única entrada de los modelos.** |
| `review_title` | texto corto | "No fue lo esperado" | **Excluida por fuga**: con solo el título, un TF-IDF obtiene F1 0,78, más que con el comentario (0,72). En chats y reclamos no existe título. |
| `review_rate` | entero 1–5 | 2 | **Excluida por fuga**: es la fuente de la etiqueta. Se usa solo en el análisis de errores. |
| `sentimiento` | categórica | `negativo` | **Variable objetivo.** |

## Esquema de etiquetas
La etiqueta **no fue anotada por personas**: se deriva de las estrellas (supervisión débil).

| Estrellas | Sentimiento | Acción de negocio |
|---|---|---|
| 1 – 2 | negativo (40%) | Alerta temprana, cola prioritaria de atención |
| 3 | neutro (20%) | Oportunidad de mejora |
| 4 – 5 | positivo (40%) | Testimonio, reputación del vendedor |

**Clase crítica:** negativo. **Error más costoso:** negativo leído como positivo.

## Auditoría (base completa)
| Revisión | Hallazgo | Decisión |
|---|---|---|
| Nulos | 0 en todas las columnas | — |
| Duplicados exactos de texto | 713 | Se conserva una copia (evita fuga entre conjuntos) |
| Textos contradictorios | 34 (mismo texto, distinto sentimiento) | Se eliminan |
| Filas finales | 486.621 | — |
| Balance | 1★…5★ casi uniformes (~97 mil c/u) → 40/20/40 | Pesos de clase balanceados en todos los modelos |
| Longitud | Mediana 30 palabras · P95 75 · P99 118 | RNN: 100 tokens; BETO: 128 subpalabras (trunca < 2%) |
| Países | MLA 43%, MLM 31%, MCO 10%, MLC 9%, MLU 5%, MLV 1%, MPE 0,7% | Reportar F1 por país; Perú y Venezuela con pocos datos |
| PII | 0 correos, 25 URL y 194 posibles teléfonos (muchos son números de modelo o medidas) | Se enmascaran con `<EMAIL>`, `<URL>`, `<TEL>` |
| **Sesgo de longitud** | 17,5% del corpus tiene ≤ 15 palabras y el **93,8%** de esos comentarios es negativo; los positivos casi siempre son largos (efecto del filtro de selección por longitud) | Los modelos aprenden "corto ⇒ negativo" (recall de positivos cortos: BETO 9,5%). Control: sin alertas automáticas para textos cortos; reentrenar con cortos balanceados |
| **Vocabulario de posventa ausente** | "llegó", "envío", "entrega" y "servicio" no aparecen en ninguna reseña; "vendedor" aparece 2 veces | Las reseñas hablan del producto, no del envío ni del vendedor. Antes de usar el modelo en chats o reclamos hay que reentrenar con ese tipo de texto |
| Fuga | Título y estrellas; además el texto a veces dice "doy tres estrellas" | Título y estrellas excluidos; la mención en el texto queda documentada |

## Partición
Aleatoria y estratificada por sentimiento, semilla 42: **70% entrenamiento (340.634) · 15% validación (72.993) · 15% prueba (72.994)**.
No hay fecha, cliente ni conversación para una partición temporal o por entidad; las reseñas son independientes.
Para comparar con BETO se usan además subconjuntos estratificados comunes de 5.000 comentarios de validación y de prueba.

## Población, cobertura y limitaciones
- Compradores de Mercado Libre en 7 países hispanohablantes que dejaron una reseña: no representa a quien no escribe reseñas
  ni el lenguaje de chats, preguntas o reclamos (cambio de dominio esperado en producción).
- 5 categorías de producto; Alimentos y Bebidas tiene solo 1% de los datos.
- Etiqueta ruidosa: la persona puede poner 3 estrellas y escribir una queja (o al revés).
- Sin fecha: no se puede medir *drift* temporal con estos datos.

## Usos no previstos
- Tomar decisiones automáticas sobre personas (sanciones a vendedores o compradores) sin revisión humana.
- Inferir atributos personales (género, edad, ubicación exacta) a partir del texto.
- Aplicarlo a portugués u otros idiomas, o a dominios distintos al comercio electrónico, sin reentrenar y volver a evaluar.
- Uso comercial sin revisar la licencia del dataset.
