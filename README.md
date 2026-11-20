# ¿Qué siente el comprador? Sentimiento en reseñas de Mercado Libre

Proyecto final de **Procesamiento de Lenguaje Natural**, Universidad Ean. Docente: Gerson José Moreno Morales.

| Integrante | Perfil | En la sustentación |
|---|---|---|
| Pedro García Romero | Especialista en Machine Learning | Modelos, BETO, resultados y demo |
| Carlos Trilleras | Magíster en Ciencia de Datos | Datos, fuga, errores, robustez y sesgo de longitud |
| Fabiola Aldana Mora | Magíster en Gerencia en IA | Problema, valor, abstención, costo, riesgos y decisión |

Queríamos saber si vale la pena usar un Transformer preentrenado en español (BETO) para leer el sentimiento de los compradores, o si
basta con algo más simple. Comparamos una línea base TF-IDF con regresión logística y SVM, las redes recurrentes de la sesión 3 (RNN,
LSTM y GRU) y BETO, de la sesión 4, todos sobre la misma partición. Después pensamos cómo llevar el ganador a operación sin que tome
decisiones que le corresponden a una persona.

- **Usuario:** el analista de experiencia del comprador.
- **Decisión:** qué comentario atender primero y a qué vendedor avisarle a tiempo.
- **Entrada y salida:** un comentario en español entra; sale la etiqueta, una probabilidad calibrada, la decisión (automática o revisión humana) y las palabras que más pesaron.
- **Clase crítica:** negativo. Lo peor que puede pasar es leer como positivo a un cliente molesto.

## Lo que encontramos

Todos los modelos se compararon sobre la misma prueba de 5.000 comentarios:

| Modelo | Datos de entrenamiento | F1 macro | F1 neutro | Recall negativo | p95 por comentario (CPU) |
|---|---:|---:|---:|---:|---:|
| **BETO** | 16.000 | **0,793** | **0,624** | 0,854 | 150,1 ms |
| BETO int8 (cuantizado) | 16.000 | 0,773 | 0,592 | 0,876 | 121,3 ms |
| LSTM | 340.634 | 0,766 | 0,578 | 0,833 | 4,9 ms |
| GRU (mejor recurrente en validación) | 340.634 | 0,758 | 0,584 | 0,824 | 6,5 ms |
| RNN | 340.634 | 0,753 | 0,569 | 0,798 | 4,4 ms |
| TF-IDF + regresión logística | 340.634 | 0,742 | 0,556 | 0,798 | 1,1 ms |
| TF-IDF + SVM lineal | 340.634 | 0,721 | 0,481 | 0,851 | 1,1 ms |
| TF-IDF + LR con los mismos datos que BETO | 16.000 | 0,704 | 0,493 | 0,785 | 1,4 ms |

- La ventaja de BETO sobre GRU es de **+0,035 de F1 (IC 95% [0,024; 0,045])**, y la logra con veinte veces menos datos.
- **Abstención:** con probabilidades calibradas y un umbral de 0,69, BETO decide solo el **78%** de los comentarios con recall de negativo **0,92**. El resto lo revisa una persona.
- **Fuera de dominio:** un control previo compara cada texto con reseñas reales. Detectó los 20 textos ajenos que probamos (deportes, recetas, código, otros idiomas) y marca el 2,7% de las reseñas reales.
- **El hallazgo que más nos sorprendió:** el 94% de los comentarios de 15 palabras o menos son negativos, porque los autores del corpus seleccionaron reseñas largas. Los modelos aprendieron que "corto" es casi lo mismo que "queja". Por eso los comentarios cortos nunca disparan una alerta automática.
- **Decisión:** proponemos un piloto de BETO en modo sombra, por lotes y con abstención, con TF-IDF como respaldo. El detalle está en la sección 9 del notebook y en `resultados/decision.json`.

## Demostración

**Demo en línea:** https://sentimiento-melisa-demo.onrender.com (corre TF-IDF; la primera visita puede tardar
unos 50 segundos porque el plan gratuito se duerme).


**En el navegador (recomendado para la sustentación).** La interfaz muestra el flujo completo: enmascarar datos personales, revisar si
el texto es una reseña, clasificar, aplicar el umbral y explicar la decisión.
```bash
pip install -r requirements.txt
python app/app.py            # abre http://localhost:5000
```
Si existe la carpeta `modelos/` (la genera el notebook), la demo usa BETO. Si no, usa TF-IDF con la misma política de decisión.

**En la terminal:**
```bash
python src/demo.py "Me llegó el cargador y a la semana dejó de funcionar, y el vendedor nunca respondió"
```

**En Render.** La versión en línea (https://sentimiento-melisa-demo.onrender.com) se despliega con `render.yaml`: en Render, *New → Blueprint* y se elige
este repositorio. El plan gratuito tiene 512 MB de memoria y BETO no cabe, así que en Render la demo usa TF-IDF + regresión logística (el
modelo va en `app/modelo/`) con la misma calibración, abstención y regla de comentarios cortos. Para mostrar BETO en línea se
necesita un plan de al menos 2 GB o un Space de Hugging Face.

## Qué hay en el repositorio
```
├── Proyecto_PLN_Sentimiento_MeLiSA.ipynb   Notebook principal, ejecutado de principio a fin
├── app/                                    Demo web (Flask): app.py, plantilla HTML y modelo liviano para Render
├── data/
│   ├── README.md                           Data card: origen, periodo, diccionario, auditoría, limitaciones
│   └── muestra_melisa_estratificada.csv    Muestra de 5.001 reseñas con datos personales enmascarados
├── src/
│   ├── servicio.py                         Lógica de decisión compartida por la demo web y la de terminal
│   ├── demo.py                             Demo por terminal
│   ├── crear_muestra.py                    Genera la muestra de data/
│   ├── generar_presentacion.py             Presentación desde resultados/resultados.json
│   ├── generar_informe.py                  Informe técnico desde resultados/resultados.json
│   ├── generar_matriz_riesgos.py           Matriz de riesgos
│   └── generar_guion.py                    Guion de la sustentación y notas del orador
├── resultados/                             Métricas, figuras, decisión del comité y referencia del control de dominio
├── docs/                                   Presentación, informe, guion y matriz de riesgos
├── render.yaml                             Despliegue de la demo en Render
└── requirements.txt · .env.example · LICENSE
```

## Entregables
| N.º | Entregable | Dónde |
|---|---|---|
| 1 | Repositorio con código, README, dependencias y licencia | Este repositorio |
| 2 | Notebook principal | [`Proyecto_PLN_Sentimiento_MeLiSA.ipynb`](Proyecto_PLN_Sentimiento_MeLiSA.ipynb) |
| 3 | Muestra de datos, diccionario y origen | [`data/`](data/) |
| 4 | Informe técnico | [`docs/Informe_Tecnico_Proyecto_PLN.docx`](docs/Informe_Tecnico_Proyecto_PLN.docx) |
| 5 | Presentación (18 diapositivas) y guion por rol | [`docs/Presentacion_Proyecto_PLN.pptx`](docs/Presentacion_Proyecto_PLN.pptx) · [`docs/Guion_Sustentacion_por_Rol.docx`](docs/Guion_Sustentacion_por_Rol.docx) |
| 6 | Matriz de riesgos | [`docs/Matriz_Riesgos.xlsx`](docs/Matriz_Riesgos.xlsx) |
| 7 | Demostración | https://sentimiento-melisa-demo.onrender.com (en línea), `python app/app.py` (web con BETO), `python src/demo.py` (terminal) y sección 10 del notebook |

## Cómo reproducirlo
Usamos Python 3.11 en un portátil con Windows 11, 12 hilos de CPU, 16 GB de RAM y sin GPU. La corrida completa del notebook tomó
unas tres horas; casi todo ese tiempo se fue en ajustar BETO.

```bash
python -m venv .venv
.venv\Scripts\activate          # en Linux o Mac: source .venv/bin/activate
pip install -r requirements.txt
```

1. **Datos.** Descarguen MeLiSA de Hugging Face (`lpsc-fiuba/melisa`), dejen solo los sitios en español (todos menos `MLB`),
   agreguen la columna `sentimiento` (1–2★ negativo, 3★ neutro, 4–5★ positivo) y guárdenla como
   `data/data_melisa_comentarios.csv`, con las columnas `id, country, category, review_content, review_title, review_rate,
   sentimiento`. Son 136 MB y GitHub no los acepta. Si el archivo no está, el notebook usa la muestra de `data/`.
2. **Notebook.** Ejecútenlo completo. Los parámetros se cambian con variables de entorno (`.env.example`); con `FRACCION=0.05` se
   hace una corrida rápida de prueba. La primera vez descarga BETO (~440 MB).
3. **Documentos.** Los scripts de `src/generar_*.py` regeneran la presentación, el informe, la matriz y el guion a partir de
   `resultados/resultados.json`.

Usamos la semilla 42 en todo. La prueba se mira una sola vez: C, épocas, temperatura, umbral y modelo campeón se eligen en validación.

## Datos y licencias
MeLiSA (LPSC-FIUBA, Universidad de Buenos Aires) son reseñas reales y públicas recolectadas entre agosto de 2020 y enero de 2021, con
licencia "other"; las usamos con fines académicos. El código del repositorio es MIT. BETO (CC BY 4.0) y MeLiSA conservan sus
licencias. Más detalle en [`data/README.md`](data/README.md) y [`LICENSE`](LICENSE).

## Uso de IA generativa
Como pide el curso, lo declaramos: usamos Claude (Anthropic) como asistente de programación para depurar y organizar el código,
escribir los scripts que generan los documentos y revisar la redacción. Las decisiones de diseño, la interpretación de los
resultados, la revisión manual de errores y la recomendación son nuestras. Todas las cifras salen de la ejecución del notebook y las
verificamos. No enviamos datos confidenciales a ningún servicio externo: el corpus es público.
