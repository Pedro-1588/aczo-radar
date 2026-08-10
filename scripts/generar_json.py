#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ACZO · Genera la API pública del radar (JSON) desde los CSV de datos/.
Salida en docs/api/ — es lo que consume Tailor y lo que pinta el panel web.
Lo llama actualizar_mercados.py al final del ciclo diario.
"""
import csv, datetime, json, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATOS = os.path.join(ROOT, "datos")
API = os.path.join(ROOT, "docs", "api")
os.makedirs(API, exist_ok=True)

def cfg():
    c = {"umbral_20": 110, "umbral_6p": 110, "mejor_fijo": 110,
         "peajes": 28, "cobertura": 10, "umbral_gas": 55}
    p = os.path.join(ROOT, "config.csv")
    if os.path.exists(p):
        with open(p) as f:
            rd = csv.reader(f); next(rd, None)
            for row in rd:
                if len(row) >= 2 and row[0] in c:
                    try: c[row[0]] = float(row[1])
                    except ValueError: pass
    return c

def leer(nombre):
    p = os.path.join(DATOS, nombre)
    if not os.path.exists(p): return []
    with open(p) as f:
        rd = csv.reader(f); next(rd, None)
        return [r for r in rd]

def num(x):
    try: return float(x)
    except (TypeError, ValueError): return None

def escribir(nombre, obj):
    with open(os.path.join(API, nombre), "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    print("  ·", nombre)

C = cfg()
hoy = datetime.date.today()
sello = datetime.datetime.now().isoformat(timespec="seconds")

# ---------------- LUZ · OMIE mensual (2.0 y 6 periodos) ----------------
men20 = [{"mes": f"{r[0]}-{int(r[1]):02d}",
          "p1": round(num(r[2]) * 1000, 2), "p2": round(num(r[3]) * 1000, 2), "p3": round(num(r[4]) * 1000, 2),
          "media": round(sum(num(v) for v in r[2:5]) / 3 * 1000, 2)}
         for r in leer("omie_mensual_periodos.csv") if num(r[2]) is not None]

men6 = []
for r in leer("omie_6p.csv"):
    fila = {"mes": f"{r[0]}-{int(r[1]):02d}"}
    vals = []
    for i, p in enumerate(("p1", "p2", "p3", "p4", "p5", "p6")):
        v = num(r[2 + i])
        fila[p] = round(v * 1000, 2) if v is not None else None
        if v is not None: vals.append(v)
    fila["media"] = round(sum(vals) / len(vals) * 1000, 2) if vals else None
    men6.append(fila)

dia_luz = [{"fecha": r[0], "min": num(r[1]), "media": num(r[2]), "max": num(r[3])}
           for r in leer("omie_diario_minmax.csv")]

curva_luz = [{"clave": r[0], "instrumento": r[1], "precio": num(r[2]), "peso_forward": int(r[3])}
             for r in leer("omip.csv")]
fecha_omip = ""
p = os.path.join(DATOS, "omip_fecha.txt")
if os.path.exists(p): fecha_omip = open(p).read().strip()

peso_total = sum(c["peso_forward"] for c in curva_luz) or 1
forward_luz = round(sum(c["precio"] * c["peso_forward"] for c in curva_luz) / peso_total, 2)
fijo_teorico = round(forward_luz + C["peajes"] + C["cobertura"], 2)
ventana = round(fijo_teorico - C["mejor_fijo"], 2)
lectura = "ABIERTA" if ventana > 5 else ("CATALOGO_CARO" if ventana < -5 else "NEUTRAL")

media_mes_luz = round(sum(d["media"] for d in dia_luz) / len(dia_luz), 2) if dia_luz else None
ult12 = [m["media"] for m in men20[-12:]]
media_12m_luz = round(sum(ult12) / len(ult12), 2) if ult12 else None

# ---------------- GAS · MIBGAS ----------------
men_gas = [{"mes": f"{r[0]}-{int(r[1]):02d}", "media": num(r[2])}
           for r in leer("mibgas_mensual.csv") if num(r[2]) is not None]
dia_gas = [{"fecha": r[0], "precio": num(r[1])} for r in leer("mibgas_diario.csv")]
curva_gas = [{"producto": r[0], "etiqueta": r[1], "entrega_ini": r[2], "entrega_fin": r[3],
              "precio": num(r[4]), "peso_forward": int(r[5])} for r in leer("mibgas_curva.csv")]
fecha_gas = ""
p = os.path.join(DATOS, "mibgas_fecha.txt")
if os.path.exists(p): fecha_gas = open(p).read().strip()

peso_g = sum(c["peso_forward"] for c in curva_gas) or 1
forward_gas = round(sum(c["precio"] * c["peso_forward"] for c in curva_gas) / peso_g, 2)
media_mes_gas = round(sum(d["precio"] for d in dia_gas) / len(dia_gas), 2) if dia_gas else None
g12 = [m["media"] for m in men_gas[-12:]]
media_12m_gas = round(sum(g12) / len(g12), 2) if g12 else None
curva_vs_spot = round(forward_gas - media_mes_gas, 2) if media_mes_gas else None

# ---------------- salidas ----------------
print("Generando API en docs/api/")

escribir("resumen.json", {
    "actualizado": sello,
    "unidad": "EUR/MWh",
    "luz": {
        "fuente": "OMIE (spot) + OMIP (futuros)",
        "spot_mes_en_curso": {"mes": f"{hoy:%Y-%m}", "media": media_mes_luz, "dias_con_dato": len(dia_luz)},
        "spot_ultimo_mes_cerrado": men20[-1] if men20 else None,
        "spot_media_12m": media_12m_luz,
        "forward_12m": forward_luz,
        "fecha_curva_omip": fecha_omip,
        "fijo_teorico": fijo_teorico,
        "mejor_fijo_catalogo": C["mejor_fijo"],
        "ventana_fijacion": ventana,
        "lectura": lectura,
        "parametros": {"peajes_ajustes": C["peajes"], "cobertura": C["cobertura"]},
    },
    "gas": {
        "fuente": "MIBGAS (spot PVB) + MIBGAS Derivatives (futuros)",
        "spot_mes_en_curso": {"mes": f"{hoy:%Y-%m}", "media": media_mes_gas, "dias_con_dato": len(dia_gas)},
        "spot_ultimo_mes_cerrado": men_gas[-1] if men_gas else None,
        "spot_media_12m": media_12m_gas,
        "forward_12m": forward_gas,
        "fecha_curva": fecha_gas,
        "forward_menos_spot": curva_vs_spot,
        "estructura": ("BACKWARDATION" if curva_vs_spot is not None and curva_vs_spot < -2
                       else "CONTANGO" if curva_vs_spot is not None and curva_vs_spot > 2 else "PLANA"),
    },
})
escribir("luz_omie_mensual.json", {"actualizado": sello, "unidad": "EUR/MWh",
                                   "periodos_2_0TD": men20, "periodos_6P_30_61TD": men6})
escribir("luz_omie_diario.json", {"actualizado": sello, "unidad": "EUR/MWh",
                                  "mes": f"{hoy:%Y-%m}", "dias": dia_luz})
escribir("luz_omip_curva.json", {"actualizado": sello, "unidad": "EUR/MWh",
                                 "fecha_curva": fecha_omip, "forward_12m": forward_luz,
                                 "instrumentos": curva_luz})
escribir("gas_mibgas_mensual.json", {"actualizado": sello, "unidad": "EUR/MWh", "meses": men_gas})
escribir("gas_mibgas_diario.json", {"actualizado": sello, "unidad": "EUR/MWh",
                                    "mes": f"{hoy:%Y-%m}", "dias": dia_gas})
escribir("gas_mibgas_curva.json", {"actualizado": sello, "unidad": "EUR/MWh",
                                   "fecha_curva": fecha_gas, "forward_12m": forward_gas,
                                   "productos": curva_gas})
print(f"OK · luz: {len(men20)} meses, {len(dia_luz)} días, forward {forward_luz} | "
      f"gas: {len(men_gas)} meses, {len(dia_gas)} días, forward {forward_gas}")
