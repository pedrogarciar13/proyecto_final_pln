# Declaración de uso de IA generativa

Las especificaciones del proyecto (sección 10) permiten usar asistentes de IA siempre que el equipo responda por el contenido y declare
cómo los usó. Esta es nuestra declaración. Preferimos contarlo con detalle en lugar de quedarnos cortos.

> Los espacios marcados con **[COMPLETAR]** solo los puede llenar el equipo. Hay que llenarlos antes de entregar.

## 1. Herramientas y para qué las usamos

| Herramienta | Para qué |
|---|---|
| **Claude Code (Anthropic)**, en la aplicación de escritorio | Asistente principal de programación y redacción (detalle abajo). |
| [COMPLETAR: ChatGPT, Copilot, Gemini u otra, si alguien la usó] | [COMPLETAR: para qué] |

Con Claude Code hicimos esto:

- **Código.** Escribió y depuró buena parte del código del notebook: limpieza, líneas base TF-IDF, redes recurrentes, ajuste de BETO,
  calibración, abstención, intervalos bootstrap, control de fuera de dominio y contrafactuales. También escribió los scripts de
  `src/` que generan la presentación, el informe, la matriz de riesgos y el guion; la demo web (`app/`); la demo que corre BETO en el
  navegador (`demo_navegador/`), con la conversión del modelo a ONNX int8; y la solución del taller de la Sesión 4.
- **Análisis.** Después de la revisión del proyecto, propuso e implementó varios análisis: el control de textos fuera de dominio, la
  variabilidad entre semillas, el resumen agregado, el caso de valor con sensibilidad y el hallazgo sobre el vocabulario de posventa.
- **Redacción.** Redactó los borradores del informe, el README, la data card, el guion, el texto de las diapositivas y las lecturas
  de las celdas del notebook. Nosotros los revisamos y ajustamos.
- **Despliegue.** Guió la publicación de la demo en Render y del modelo y la demo en Hugging Face.

Lo que hizo el equipo:

- Definir el problema, el usuario y la decisión que apoya el modelo, y elegir el corpus MeLiSA.
- Revisar el código y los resultados, aprobar o pedir cambios en cada entrega y validar las conclusiones.
- Tomar la decisión final del comité (piloto de BETO en modo sombra con controles) y preparar y presentar la sustentación.
- [COMPLETAR: lo que hizo cada integrante, por ejemplo:]
  - Pedro García Romero (Machine Learning): [COMPLETAR]
  - Carlos Trilleras (Ciencia de Datos): [COMPLETAR]
  - Fabiola Aldana Mora (Gerencia de IA): [COMPLETAR]

## 2. Qué verificamos y cómo

- **Cifras.** Ninguna cifra del informe, la presentación ni el README se escribió a mano. Todas salen de ejecutar el notebook completo
  (unas tres horas en CPU) y quedan en `resultados/resultados.json`. Los scripts de `src/` leen ese archivo para armar los documentos,
  así que un número en un documento siempre viene de una ejecución.
- **Afirmaciones.** Varias interpretaciones se escribieron antes de ver el resultado y salieron mal. Las corregimos después de revisar
  las salidas:
  - la robustez en pares largos de contrafactuales;
  - si el control de fuera de dominio marcaba las quejas de envío;
  - las conclusiones del taller de la Sesión 4, donde el Transformer no le ganó a TF-IDF.

  Una interpretación queda en los entregables solo si coincide con la salida ejecutada.
- **Código.** Lo ejecutamos de principio a fin. Las partes críticas tienen comprobaciones explícitas en el notebook: la misma partición
  para todos los modelos, la prueba se mira una sola vez y se usan tres semillas para las redes recurrentes. La demo del navegador se
  comparó con BETO original en 1.000 reseñas (98% de etiquetas iguales).
- **Errores del modelo.** La revisión de 20 errores de BETO (Tabla 7 del informe) la redactó la IA a partir de los ejemplos.
  [COMPLETAR: confirmar que el equipo leyó cada caso y está de acuerdo con la causa asignada.]
- **Referencias.** [COMPLETAR: confirmar que se revisaron los enlaces y las citas del anexo de referencias (MeLiSA, BETO, artículos).]

## 3. Datos y servicios externos

- El corpus MeLiSA es público. No usamos datos de clientes ni datos confidenciales de ninguna organización.
- El texto de las reseñas puede traer correos o teléfonos. Por eso los enmascaramos antes de entrenar, en la muestra publicada en
  `data/` y en las demos.
- Al asistente le compartimos el código, los resultados y muestras de reseñas públicas. Nada más.
- En Hugging Face publicamos el modelo ajustado (`pedrogarciar13/beto-sentimiento-melisa`) y la demo. Ninguno de los dos contiene
  datos personales.

## 4. Contenido generado frente a evidencia experimental

- Los resultados que presentamos como evidencia (métricas, intervalos, tiempos, costos calculados, ejemplos de la demo) salen de
  código ejecutado. Ninguno lo inventó la IA.
- El caso de valor usa **supuestos ilustrativos**: volumen, costo por hora, tasa de devoluciones evitadas. Así lo decimos en el
  informe y en la presentación. No son datos medidos.
- Los ejemplos de texto de la demo y de los contrafactuales los escribimos para probar el modelo. No son reseñas reales.

## 5. Trazabilidad y autoría

- **Historial.** El historial de commits de este repositorio muestra cómo se construyó cada entrega. Todos los commits salieron de una
  sola cuenta (`pedrogarciar13`); los que tienen la línea `Co-Authored-By: Claude`, que son casi todos, se hicieron con el asistente.
  [COMPLETAR, si aplica: cada integrante puede subir sus revisiones desde su propia cuenta.]
- **Decisión.** La decisión del comité y sus condiciones están en `resultados/decision.json`, y las razones, en la sección 9 del
  notebook y en el informe.
- **Responsables.** La autoría y los roles están en la tabla del README y en el guion de la sustentación. Cada integrante responde por
  la parte que presenta.
