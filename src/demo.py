"""Demostración del flujo completo desde la terminal.

Requiere haber ejecutado el notebook principal (genera la carpeta modelos/). Para la versión web, ver app/.

Uso:
    python src/demo.py "El producto llegó roto y nadie responde"
    python src/demo.py            # modo interactivo (línea vacía para salir)
    python src/demo.py --modelo "TF-IDF + Regresión Logística" "Excelente, lo recomiendo"

Salida: etiqueta, probabilidades calibradas, confianza, decisión (automática / revisión humana), alerta al vendedor, control de
fuera de dominio, advertencia por comentario corto y evidencia (términos que más pesaron según la regresión logística).
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from servicio import RAIZ, Servicio  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("texto", nargs="*")
    ap.add_argument("--modelo", default=None, help="Por defecto, el aprobado por el comité (resultados/decision.json)")
    ap.add_argument("--guardar", action="store_true", help="Guarda la salida en resultados/demo_ejemplo.json")
    a = ap.parse_args()
    srv = Servicio(a.modelo)
    if a.texto:
        entrada = " ".join(a.texto)
        salida = srv.predecir(entrada)
        print(json.dumps(salida, ensure_ascii=False, indent=1))
        if a.guardar:   # ejemplo que usa la presentación (diapositiva de demostración)
            json.dump({"entrada": entrada, **salida}, open(os.path.join(RAIZ, "resultados", "demo_ejemplo.json"), "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
        return
    print(f"Modelo: {srv.nombre} · umbral de abstención {srv.umbral:.2f} · escriba un comentario (línea vacía para salir)")
    for linea in sys.stdin:
        if not linea.strip():
            break
        print(json.dumps(srv.predecir(linea), ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
