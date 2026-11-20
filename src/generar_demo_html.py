"""Página HTML estática con la salida real de la demo (respaldo para la sustentación, se abre sin servidor).

Uso: python src/generar_demo_html.py   → docs/Demo_Respaldo.html
"""
import html
import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from servicio import RAIZ, Servicio  # noqa: E402

EJEMPLOS = [
    ("Queja larga sobre el producto", "Lo usé una semana y dejó de funcionar. La calidad es pésima, el material se siente barato y no lo recomiendo para nada."),
    ("Elogio corto (sesgo de longitud)", "Excelente producto, lo recomiendo"),
    ("El mismo elogio con contexto", "Excelente producto, se nota la buena calidad del material, funciona tal como dice la descripción y lo recomiendo a quien lo necesite"),
    ("Opinión mixta", "La tela es buena y el diseño es bonito, aunque la talla es más pequeña de lo esperado y el color es más claro que en la foto."),
    ("Datos personales", "Dejó de funcionar a la semana y la calidad es muy mala, escríbanme a juan.perez@correo.com o al 3001234567"),
    ("Queja de envío (vocabulario que no está en el corpus)", "El paquete llegó tarde y el vendedor nunca respondió mis mensajes, pésimo servicio de entrega"),
    ("Fuera de dominio", "El partido terminó dos a uno y el delantero fue expulsado en el segundo tiempo por una falta muy fuerte"),
    ("Entrada vacía", ""),
]
COLOR = {"negativo": "#E4572E", "neutro": "#94A3B8", "positivo": "#2BB673"}


def tarjeta(titulo, texto, r):
    if r.get("estado") != "ok":
        cuerpo = f'<div class="aviso">{html.escape(r["detalle"])} Decisión: {r["decision"]}.</div>'
    else:
        barras = "".join(
            f'<div class="barra"><span>{c}</span><div class="pista"><div class="relleno" style="width:{100 * p:.1f}%;background:{COLOR[c]}">'
            f'</div><div class="umbral" style="left:{100 * r["umbral"]:.1f}%"></div></div><b>{100 * p:.0f}%</b></div>'
            for c, p in r["probabilidades"].items())
        chips = "".join(f'<span class="chip">{html.escape(e)}</span>' for e in r["evidencia"]) or '<span class="meta">—</span>'
        avisos = "".join(f'<div class="aviso">{html.escape(a)}</div>' for a in r["advertencias"])
        alerta = '<span class="decision alerta">Alerta al vendedor</span>' if r["alerta_vendedor"] else ""
        clase_dec = "auto" if r["decision"] == "automática" else "humana"
        pii = " · PII enmascarada: " + html.escape(r["texto"]) if r["pii_enmascarada"] else ""
        cuerpo = (f'<div class="fila"><span class="etiqueta" style="background:{COLOR[r["etiqueta"]]}">{r["etiqueta"]}</span>'
                  f'<span class="decision {clase_dec}">Decisión {r["decision"]}</span>{alerta}</div>{barras}'
                  f'<p class="meta">Confianza {r["confianza"]} · umbral {r["umbral"]} · {r["palabras"]} palabras{pii}</p>'
                  f'<div class="meta">Evidencia:</div><div class="chips">{chips}</div>{avisos}')
    return (f'<section class="card"><h2>{html.escape(titulo)}</h2><p class="entrada">“{html.escape(texto) or "(vacío)"}”</p>{cuerpo}</section>')


def main():
    srv = Servicio()
    tarjetas = "".join(tarjeta(t, x, srv.predecir(x)) for t, x in EJEMPLOS)
    plantilla = open(os.path.join(RAIZ, "app", "templates", "index.html"), encoding="utf-8").read()
    estilos = plantilla[plantilla.index("<style>"):plantilla.index("</style>") + 8]
    pagina = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Demo de respaldo · Sentimiento MeLiSA</title>{estilos}
<style>main {{ display:grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); }} .entrada {{ font-style:italic; color:#334155; }}</style>
</head><body><header><div class="wrap"><small>Proyecto final PLN · Universidad Ean · respaldo de la demostración</small>
<h1>¿Qué siente el comprador?</h1><p>Salida real del servicio con el modelo <b>{srv.nombre}</b> (umbral {srv.umbral:.2f}) · generada el {date.today():%d/%m/%Y}.
La demo interactiva está en <code>python app/app.py</code>.</p></div></header><main>{tarjetas}</main>
<footer>Pedro García Romero · Carlos Trilleras · Fabiola Aldana Mora. Barras: probabilidad calibrada por clase; la línea vertical es el umbral de abstención.</footer></body></html>"""
    salida = os.path.join(RAIZ, "docs", "Demo_Respaldo.html")
    open(salida, "w", encoding="utf-8").write(pagina)
    print("Demo de respaldo ->", salida)


if __name__ == "__main__":
    main()
