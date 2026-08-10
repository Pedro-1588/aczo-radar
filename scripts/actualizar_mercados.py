#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ACZO · Actualizador diario del Radar de Mercados.
1) Descarga de OMIE los días que falten (incluido mañana si ya está publicado).
2) Cierra meses: añade a los CSV mensuales (2.0TD y 6 periodos) los meses completos nuevos.
3) Regenera el CSV diario del mes en curso (mín/media/máx cuartohorarios).
4) Lee la curva de futuros de OMIP y calcula los pesos del forward 12 meses.
5) Reconstruye ACZO_Radar_Mercados.xlsx.
Pensado para ejecutarse a diario a las 21:15 vía launchd. Sin argumentos.
"""
import csv, datetime, os, re, ssl, sys, html, subprocess
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.dirname(os.path.abspath(__file__))
SP = os.path.join(ROOT, "datos")
LOG = os.path.join(ROOT, "log.txt")

CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE

FESTIVOS = {(1,1),(1,6),(5,1),(8,15),(10,12),(11,1),(12,6),(12,8),(12,25),
            (2026,4,3),(2027,3,26)}  # fijos nacionales + Viernes Santo por año

def es_festivo(d):
    return (d.month,d.day) in FESTIVOS or (d.year,d.month,d.day) in FESTIVOS

def log(msg):
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(LOG,"a") as f: f.write(f"{stamp}  {msg}\n")
    print(msg)

# ---------------- OMIE ----------------
def descargar_dia(d):
    """Devuelve lista de 24 precios horarios (€/MWh) o None si no publicado."""
    for ver in (1,2,3):
        url = f"https://www.omie.es/es/file-download?parents=marginalpdbc&filename=marginalpdbc_{d:%Y%m%d}.{ver}"
        try:
            with urllib.request.urlopen(url, context=CTX, timeout=30) as r:
                t = r.read().decode("utf-8", errors="replace")
        except Exception:
            continue
        if "MARGINAL" not in t:
            continue
        filas = []
        for ln in t.splitlines():
            if ";" not in ln or "MARGINAL" in ln: continue
            f = ln.split(";")
            if len(f) < 6: continue
            try: filas.append((int(f[3]), float(f[5])))
            except ValueError: continue
        if not filas: continue
        maxper = max(p for p,_ in filas)
        if maxper > 25:  # cuartohorario desde oct-2025: 4 cuartos = 1 hora
            horas = {}
            for p,v in filas: horas.setdefault((p-1)//4+1, []).append(v)
            return {h: sum(v)/len(v) for h,v in horas.items()}, [v for _,v in filas]
        return {p:v for p,v in filas}, [v for _,v in filas]
    return None

def periodo20(d,h):
    if d.weekday()>=5 or es_festivo(d): return "P3"
    if h<=8: return "P3"
    return "P2" if h in (9,10,15,16,17,18,23,24) else "P1"

SEASON={1:(1,2),2:(1,2),3:(2,3),4:(4,5),5:(4,5),6:(3,4),7:(1,2),8:(3,4),9:(3,4),10:(4,5),11:(2,3),12:(1,2)}
PUNTA={10,11,12,13,14,19,20,21,22}

def periodo6(d,h):
    if d.weekday()>=5 or es_festivo(d) or h<=8: return "P6"
    hi,lo = SEASON[d.month]
    return f"P{hi}" if h in PUNTA else f"P{lo}"

def leer_csv(path):
    with open(path) as f:
        rd=csv.reader(f); head=next(rd); return head,[r for r in rd]

def actualizar_omie():
    hoy = datetime.date.today()
    head20,filas20 = leer_csv(os.path.join(SP,"omie_mensual_periodos.csv"))
    head6, filas6  = leer_csv(os.path.join(SP,"omie_6p.csv"))
    y,m = int(filas20[-1][0]), int(filas20[-1][1])
    # primer mes sin cerrar en los CSV
    y0,m0 = (y,m+1) if m<12 else (y+1,1)
    ini = datetime.date(y0,m0,1)
    dias, cuartos_dia = {}, {}
    d = ini
    while d <= hoy + datetime.timedelta(days=1):
        r = descargar_dia(d)
        if r is not None:
            dias[d], cuartos_dia[d] = r
        d += datetime.timedelta(days=1)
    log(f"OMIE: {len(dias)} días descargados desde {ini}")
    # cerrar meses completos
    mm = (y0,m0)
    while True:
        y_,m_ = mm
        nxt = datetime.date(y_,m_,28)+datetime.timedelta(days=4)
        fin = datetime.date(nxt.year,nxt.month,1)-datetime.timedelta(days=1)
        dias_mes = [d for d in dias if d.year==y_ and d.month==m_]
        if len(dias_mes) < fin.day: break  # mes incompleto → paramos
        agg20, agg6 = {}, {}
        for d in dias_mes:
            for h,v in dias[d].items():
                agg20.setdefault(periodo20(d,h),[]).append(v)
                agg6.setdefault(periodo6(d,h),[]).append(v)
        filas20.append([y_,m_]+[f"{sum(agg20[p])/len(agg20[p])/1000:.5f}" for p in ("P1","P2","P3")])
        f6=[y_,m_]
        for p in ("P1","P2","P3","P4","P5","P6"):
            f6.append(f"{sum(agg6[p])/len(agg6[p])/1000:.5f}" if p in agg6 else "")
        filas6.append(f6)
        log(f"OMIE: mes cerrado {y_}-{m_:02d} añadido")
        mm = (y_,m_+1) if m_<12 else (y_+1,1)
    for path,head,filas in ((os.path.join(SP,"omie_mensual_periodos.csv"),head20,filas20),
                            (os.path.join(SP,"omie_6p.csv"),head6,filas6)):
        with open(path,"w",newline="") as f:
            w=csv.writer(f); w.writerow(head); w.writerows(filas)
    # diario del mes en curso (mín/máx a nivel cuartohorario, media del día)
    mes_curso = [d for d in sorted(dias) if (d.year,d.month)==(hoy.year,hoy.month)] or \
                [d for d in sorted(dias)][-31:]
    with open(os.path.join(SP,"omie_diario_minmax.csv"),"w",newline="") as f:
        w=csv.writer(f); w.writerow(["fecha","min","media","max"])
        for d in mes_curso:
            q = cuartos_dia[d]
            w.writerow([d.isoformat(), f"{min(q):.2f}", f"{sum(q)/len(q):.2f}", f"{max(q):.2f}"])
    log(f"OMIE: diario del mes → {len(mes_curso)} días")

# ---------------- OMIP ----------------
MES_EN = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,"Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}

def leer_omip():
    hoy = datetime.date.today()
    url = f"https://www.omip.pt/es/dados-mercado?date={hoy:%Y-%m-%d}&product=EL&zone=ES&instrument=FTB"
    req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0"})
    with urllib.request.urlopen(req, context=CTX, timeout=45) as r:
        h = r.read().decode("utf-8", errors="replace")
    out = {"M":[], "Q":[], "Y":[]}
    for chunk in re.findall(r"<tr[^>]*>(.*?)</tr>", h, re.S):
        if "FTB" not in chunk: continue
        txt = html.unescape(re.sub(r"<[^>]+>", " ", chunk))
        m = re.search(r"€/MWh\s*FTB\s+(\S+)(?:\s+(\S+))?", txt)
        resto = [g for g in (m.groups() if m else ()) if g]
        if not resto: continue
        nums = [float(x) for x in re.findall(r"\b(\d+\.\d+)\b", txt) if float(x)>0]
        if not nums: continue
        precio = nums[-1]
        t1 = resto[0]
        if t1=="M" and len(resto)>1:
            mm = re.match(r"([A-Za-z]{3})-(\d{2})", resto[1])
            if mm: out["M"].append((f"FTB M {resto[1]}", 2000+int(mm.group(2)), MES_EN[mm.group(1)], precio))
        elif t1.startswith("Q"):
            mm = re.match(r"Q(\d)-(\d{2})", t1)
            if mm: out["Q"].append((f"FTB {t1}", 2000+int(mm.group(2)), int(mm.group(1)), precio))
        elif t1.startswith("YR"):
            mm = re.match(r"YR-(\d{2})", t1)
            if mm: out["Y"].append((f"FTB {t1}", 2000+int(mm.group(1)), precio))
    if not out["M"] or not out["Q"]:
        raise RuntimeError("OMIP: la página no devolvió instrumentos M/Q — se mantiene la última lectura")
    # pesos del forward: 12 meses desde el mes que viene; trimestre completo dentro de
    # la ventana manda, si no instrumento mensual, si no trimestre parcial
    y,m = (hoy.year, hoy.month+1) if hoy.month<12 else (hoy.year+1,1)
    objetivo = []
    for i in range(12):
        t = (y*12+m-1)+i
        objetivo.append((t//12, t%12+1))
    qmeses = {}
    for name,qy,qn,p in out["Q"]:
        qmeses[name] = [(qy, 3*(qn-1)+k) for k in (1,2,3)]
    mkeys = {(my,mmo):name for name,my,mmo,p in out["M"]}
    pesos = {}
    for (ty,tm) in objetivo:
        qname = next((n for n,ms in qmeses.items() if (ty,tm) in ms), None)
        if qname and all(mm_ in objetivo for mm_ in qmeses[qname]):
            pesos[qname] = pesos.get(qname,0)+1
        elif (ty,tm) in mkeys:
            pesos[mkeys[(ty,tm)]] = pesos.get(mkeys[(ty,tm)],0)+1
        elif qname:
            pesos[qname] = pesos.get(qname,0)+1
        else:
            log(f"OMIP: sin instrumento para {ty}-{tm:02d} (el chivato de pesos avisará)")
    filas = []
    for i,(name,my,mmo,p) in enumerate(out["M"][:6],1):
        filas.append([f"M{i}", name, f"{p:.2f}", pesos.get(name,0)])
    for i,(name,qy,qn,p) in enumerate(out["Q"][:7],1):
        filas.append([f"Q{i}", name, f"{p:.2f}", pesos.get(name,0)])
    for i,(name,yy,p) in enumerate(out["Y"][:3],1):
        filas.append([f"Y{i}", name, f"{p:.2f}", 0])
    with open(os.path.join(SP,"omip.csv"),"w",newline="") as f:
        w=csv.writer(f); w.writerow(["clave","instrumento","precio","peso"]); w.writerows(filas)
    with open(os.path.join(SP,"omip_fecha.txt"),"w") as f:
        f.write(hoy.isoformat())
    suma = sum(int(r[3]) for r in filas)
    log(f"OMIP: {len(filas)} instrumentos · suma de pesos {suma} (debe ser 12)")

# ---------------- main ----------------
if __name__ == "__main__":
    try:
        actualizar_omie()
    except Exception as e:
        log(f"ERROR OMIE: {e} — el Excel se regenera con lo último disponible")
    try:
        leer_omip()
    except Exception as e:
        log(f"ERROR OMIP: {e} — el Excel mantiene la última curva guardada")
    r = subprocess.run([sys.executable, os.path.join(SCRIPTS,"generar_excel.py")],
                       capture_output=True, text=True)
    if r.returncode == 0:
        log("Excel regenerado OK")
    else:
        log(f"ERROR generando Excel: {r.stderr.strip()[-400:]}")
    # radar de gas (MIBGAS): mismo ciclo diario
    g = subprocess.run([sys.executable, os.path.join(SCRIPTS,"actualizar_gas.py")],
                       capture_output=True, text=True)
    if g.returncode != 0:
        log(f"ERROR radar gas: {g.stderr.strip()[-400:]}")
    j = subprocess.run([sys.executable, os.path.join(SCRIPTS,"generar_json.py")],
                       capture_output=True, text=True)
    if j.returncode == 0:
        log("JSON de la API regenerado OK")
    else:
        log(f"ERROR generando JSON: {j.stderr.strip()[-400:]}")
    if r.returncode != 0 or g.returncode != 0 or j.returncode != 0:
        sys.exit(1)
