# ACZO · Radar de mercados

Servicio que cada noche descarga los precios de la electricidad y el gas, y publica tres cosas:

| Salida | Para quién | Dónde |
|---|---|---|
| **JSON** con la foto del mercado | El motor / Tailor | `docs/api/*.json` |
| **Panel web** | El equipo, desde el móvil | `docs/index.html` (GitHub Pages) |
| **Dos Excel** actualizados | Quien prefiera la hoja de cálculo | `excel/` |

Todo sale del mismo proceso, así que las tres salidas dicen siempre lo mismo. No depende de ningún ordenador encendido: lo ejecuta GitHub.

---

## Fuentes

| Mercado | Pasado (spot) | Futuro (curva) |
|---|---|---|
| Luz | OMIE, precio horario desde 2023, clasificado por periodos 2.0TD y por los 6 periodos de 3.0/6.1TD | OMIP (FTB España): meses, trimestres y años |
| Gas | MIBGAS, precio día-siguiente (PVB) desde 2023 | MIBGAS Derivatives: meses, trimestres, estaciones y años |

Todos los datos son públicos y se descargan de las webs oficiales.

---

## La API que consume el motor

Punto de entrada principal — la foto del mercado en un solo fichero:

```
https://<usuario>.github.io/aczo-radar/api/resumen.json
```

```json
{
  "actualizado": "2026-08-10T21:30:11",
  "unidad": "EUR/MWh",
  "luz": {
    "spot_mes_en_curso":     { "mes": "2026-08", "media": 119.33, "dias_con_dato": 11 },
    "spot_ultimo_mes_cerrado": { "mes": "2026-07", "p1": 95.03, "p2": 98.86, "p3": 112.38, "media": 102.09 },
    "spot_media_12m": 61.97,
    "forward_12m": 79.38,
    "fijo_teorico": 117.38,
    "ventana_fijacion": 7.48,
    "lectura": "ABIERTA"
  },
  "gas": {
    "spot_mes_en_curso": { "mes": "2026-08", "media": 56.82 },
    "forward_12m": 46.28,
    "forward_menos_spot": -10.54,
    "estructura": "BACKWARDATION"
  }
}
```

Series completas, por si el motor necesita el detalle:

| Fichero | Contiene |
|---|---|
| `api/luz_omie_mensual.json` | Spot mensual desde 2023: P1/P2/P3 (2.0TD) y P1…P6 (3.0/6.1TD) |
| `api/luz_omie_diario.json` | Mes en curso día a día (mínimo, media y máximo) |
| `api/luz_omip_curva.json` | Curva de futuros de la luz y pesos del forward |
| `api/gas_mibgas_mensual.json` | Spot del gas mensual desde 2023 |
| `api/gas_mibgas_diario.json` | Mes en curso día a día, incluida la entrega de mañana |
| `api/gas_mibgas_curva.json` | Curva de futuros del gas |

Cómo se lee `ventana_fijacion` (luz): es el precio fijo teórico de hoy menos el mejor fijo del catálogo. Positivo = el catálogo todavía no ha recogido la subida del mercado, fijar sale a cuenta (`lectura: ABIERTA`). Negativo = el catálogo está caro.

---

## Puesta en marcha (una sola vez)

1. Crear un repositorio en GitHub llamado `aczo-radar`.
2. Subir esta carpeta:
   ```bash
   git remote add origin https://github.com/<usuario>/aczo-radar.git
   git branch -M main
   git push -u origin main
   ```
3. En el repositorio: **Settings → Pages → Source: Deploy from a branch → Branch: `main` / carpeta `/docs`**. En un par de minutos el panel está en `https://<usuario>.github.io/aczo-radar/`.
4. Comprobar que funciona: pestaña **Actions → Radar de mercados (diario) → Run workflow**. Debe terminar en verde y publicar un commit con los datos del día.

A partir de ahí se ejecuta solo cada noche a las **19:30 UTC** (21:30 en horario de verano español, 20:30 en invierno), cuando OMIE ya ha publicado el precio del día siguiente y OMIP y MIBGAS han cerrado sesión.

---

## Mantenimiento

- **Umbrales y parámetros**: `config.csv` (umbral de color de los Excel, mejor fijo del catálogo, peajes y cobertura de la fórmula del fijo teórico). Se edita y se hace commit; el robot lo aplica esa misma noche.
- **Festivos**: a principios de año, añadir el Viernes Santo a la lista `FESTIVOS` de `scripts/actualizar_mercados.py` (los fijos ya están). Solo afecta al reparto por periodos.
- **Si un día falla una fuente**: el ciclo no se rompe. Mantiene el último dato bueno, lo anota en `log.txt` y sigue con el resto.
- **Ejecutar a mano**: pestaña Actions → *Run workflow*. O en local: `python3 scripts/actualizar_mercados.py`.

---

## Qué NO va en este repositorio

Aquí solo viven **datos públicos de mercado**. No se suben nunca tarifarios de comercializadoras, tablas de comisiones, bases de precios de canal ni datos de clientes (CUPS, titulares, facturas). Esa información vive en los archivos internos del equipo.

---

## Estructura

```
scripts/     el motor (Python)
  actualizar_mercados.py   orquesta todo: OMIE + OMIP → Excel → llama a gas y a json
  actualizar_gas.py        MIBGAS spot + curva → Excel de gas
  generar_excel.py         construye ACZO_Radar_Mercados.xlsx
  generar_excel_gas.py     construye ACZO_Radar_Gas.xlsx
  generar_json.py          construye la API de docs/api/
datos/       histórico acumulado en CSV (crece solo, nunca se reescribe entero)
excel/       los dos Excel, regenerados cada noche
docs/        panel web + api/ (lo que se publica)
log.txt      registro de cada ejecución
```
