"""
Reporte FEN · PROPUESTA con siniestros vs. proyección
=====================================================
Mismo reporte que "generar reporte observatorio.py" (clima desde el Observatorio, Pacífico desde el txt)
más una tarjeta "SINIESTROS FEN VS. PROYECCIÓN" para dos líneas:

  • Comerciales ........ línea "Comerciales" del Excel
  • Propiedad .......... "Vivienda" + "Negocios"

Por línea compara, en soles, el monto RETENIDO (y en la opción B también el monto DIRECTO) contra la
PROYECCIÓN de la aseguradora, con una barra horizontal cuya longitud total es la proyección.

  • Excel: siniestros.xlsx o data.xlsx en esta carpeta (columnas Linea, MontoDirectoUSD, MontoRetencionUSD).
  • Tipo de cambio: el mismo que usa el Excel para pasar soles a dólares (se detecta solo; hoy 3.55).
    Se puede fijar en configuracion.txt con   TIPO_CAMBIO: 3.55
  • Proyecciones (soles), en configuracion.txt:
        PROYECCION_COMERCIALES: 11900000
        PROYECCION_PROPIEDAD: 2500000
  • Gráfico:  GRAFICO_PROPUESTA: A  (retenido vs. proyección)   |   B  (retenido + directo vs. proyección)

NO envía nada a OneDrive ni por correo: el PPT y el PNG quedan en la carpeta "reporte_generado_propuesta".

Uso:  python "generar reporte propuesta.py"          (usa mensaje_observatorio.txt)
Necesita en la misma carpeta: "generar reporte observatorio.py" y "generar reporte fen.py".
"""
import importlib.util
import os
import re
import sys
import warnings

VERSION_PROP = "2026-11-24b (propuesta siniestros por línea)"
CARPETA = os.path.dirname(os.path.abspath(__file__))
SALIDA = os.path.join(CARPETA, "reporte_generado_propuesta")
PROYECCION_DEF = {"Comerciales": 11_900_000.0, "Propiedad": 2_500_000.0}
TC_DEF = 3.55


def _cargar(nombre_archivo, alias):
    ruta = os.path.join(CARPETA, nombre_archivo)
    if not os.path.exists(ruta):
        sys.exit(f"Falta '{nombre_archivo}' junto a este script ({CARPETA}).")
    spec = importlib.util.spec_from_file_location(alias, ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


obs = _cargar("generar reporte observatorio.py", "reporte_fen_observatorio")
base = obs.base


# ---------------------------------------------------------------- configuración propia
def leer_extra():
    conf = {"tc": None, "proy": dict(PROYECCION_DEF), "grafico": "A"}
    ruta = os.path.join(CARPETA, "configuracion.txt")
    if os.path.exists(ruta):
        with open(ruta, encoding="utf-8-sig") as f:
            for raw in f:
                k, sep, v = raw.strip().partition(":")
                if not sep or raw.strip().startswith("#"):
                    continue
                k, v = k.strip().upper(), v.strip().replace(",", "").replace(" ", "")
                try:
                    if k == "TIPO_CAMBIO" and v:
                        conf["tc"] = float(v)
                    elif k == "PROYECCION_COMERCIALES" and v:
                        conf["proy"]["Comerciales"] = float(v)
                    elif k == "PROYECCION_PROPIEDAD" and v:
                        conf["proy"]["Propiedad"] = float(v)
                except ValueError:
                    print(f"  [!] configuracion.txt: valor no numérico en {k}: {v}")
                if k == "GRAFICO_PROPUESTA" and v:
                    conf["grafico"] = v.upper()[:1]
    if os.environ.get("FEN_GRAFICO"):
        conf["grafico"] = os.environ["FEN_GRAFICO"].upper()[:1]
    return conf


# ---------------------------------------------------------------- Excel de siniestros
def cargar_siniestros_proyeccion(conf):
    ruta = next((os.path.join(CARPETA, n) for n in ("siniestros.xlsx", "data.xlsx")
                 if os.path.exists(os.path.join(CARPETA, n))), None)
    if not ruta:
        print("  [!] No encuentro siniestros.xlsx ni data.xlsx en la carpeta: el reporte sale sin la tarjeta.")
        return None
    try:
        import openpyxl
    except ImportError:
        print("  [!] Para leer el Excel instala openpyxl:  pip install openpyxl")
        return None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        filas = list(openpyxl.load_workbook(ruta, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))
    cab = [str(c or "").strip().lower() for c in filas[0]]

    def col(*n):
        return next((i for i, c in enumerate(cab) if c in n), None)
    i_lin, i_dir, i_ret = col("linea", "línea"), col("montodirectousd"), col("montoretencionusd")
    i_mon = col("moneda")
    if None in (i_lin, i_dir, i_ret):
        print(f"  [!] {os.path.basename(ruta)}: faltan columnas Linea / MontoDirectoUSD / MontoRetencionUSD.")
        return None
    grupos = {"Comerciales": [0, 0.0, 0.0], "Propiedad": [0, 0.0, 0.0]}
    tcs = []
    num = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
    for f in filas[1:]:
        if not f or len(f) <= max(i_lin, i_dir, i_ret) or not f[i_lin]:
            continue                                     # fila de totales / filtros
        lin = str(f[i_lin]).strip().lower()
        g = "Comerciales" if lin.startswith("comercial") else ("Propiedad" if lin in ("vivienda", "negocios", "negocio")
                                                                 else None)
        if not g:
            continue
        x = grupos[g]
        x[0] += 1
        x[1] += num(f[i_dir])
        x[2] += num(f[i_ret])
        # el Excel convierte soles a dólares con un tipo de cambio fijo: se detecta de los montos redondos en soles
        if i_mon is not None and "sol" in str(f[i_mon] or "").lower() and num(f[i_dir]) > 0:
            for tc_ in (3.40, 3.45, 3.50, 3.55, 3.60, 3.65, 3.70, 3.75, 3.80):
                if abs(num(f[i_dir]) * tc_ - round(num(f[i_dir]) * tc_, 2)) < 1e-6:
                    tcs.append(tc_)
                    break
    tc = conf["tc"] or (max(set(tcs), key=tcs.count) if tcs else TC_DEF)
    origen_tc = "configuracion.txt" if conf["tc"] else ("detectado en el Excel" if tcs else "por defecto")
    res = {"archivo": os.path.basename(ruta), "tc": tc, "origen_tc": origen_tc, "grafico": conf["grafico"],
           "lineas": []}
    for nombre, (cant, d_usd, r_usd) in grupos.items():
        proy = conf["proy"][nombre]
        res["lineas"].append({"nombre": nombre, "cant": cant, "directo": d_usd * tc, "retenido": r_usd * tc,
                              "proy": proy, "detalle": "Vivienda + Negocios" if nombre == "Propiedad" else ""})
    return res


def soles(v):
    if v >= 1e6:
        return f"S/ {v / 1e6:.1f} MM".replace(".0 MM", " MM")
    if v >= 1e3:
        return f"S/ {v / 1e3:,.0f} mil"
    return f"S/ {v:,.0f}"


def pct(a, b):
    p_ = 100 * a / b if b else 0
    return f"{p_:.1f}%" if p_ >= 1 else f"{p_:.2f}%"


def compactar(item, maximo=150):
    """Ítem más corto para que todo entre con la tarjeta de siniestros: primera oración (o hasta 'maximo'),
    sin el n.° de aviso entre paréntesis."""
    etq, sep, txt = item.partition(":")
    if not sep or len(etq) > 40:
        etq, txt = "", item
    txt = re.sub(r"\s*\((?:aviso|avisos)[^)]*\)", "", txt).strip()
    txt = re.sub(r"^Desde el \w+ \d+ al \w+ \d+ de \w+,\s*", "", txt)
    txt = re.sub(r"aviso especial de oleaje anómalo N\.° (\S+) de la DHN, vigente del \S+ al (\S+)\.",
                 r"aviso de oleaje anómalo de la DHN (\1) vigente hasta el \2.", txt)
    if len(txt) > maximo:
        fin = [m.end() for m in re.finditer(r"(?<!\b[A-ZÁÉÍÓÚÑ])\.(?=\s|$)", txt) if 50 < m.end() <= maximo]
        if fin:
            txt = txt[:fin[-1]]
        else:
            coma = txt.rfind(", ", 60, maximo)
            txt = (txt[:coma] + ".") if coma > 0 else txt[:maximo].rsplit(" ", 1)[0] + "…"
    txt = txt[:1].upper() + txt[1:]
    return f"{etq.strip()}: {txt}" if etq else txt


def compactar_calor(item):
    """'Costa y sierra: máximas de hasta 37 °C. Nivel extremo (rojo) en 11 dptos. (incl. …) (aviso …)'
    -> 'Costa y sierra: hasta 37 °C; extremo en 11 dptos. (incl. …).'"""
    t = re.sub(r"\s*\(aviso[^)]*\)", "", item)
    t = re.sub(r"máximas de hasta (\d+ °C)\.\s*Nivel ", r"hasta \1; ", t)
    t = re.sub(r"\s*\((rojo|naranja|amarillo)\)", "", t)
    return t.rstrip(".") + "."


SEDES_DPTO = ("Lima", "Piura", "Lambayeque", "La Libertad", "Junín", "Arequipa", "Cusco")


def resumir_listas(texto, maximo=4):
    """'… en A, B, C, D, E, F y G (aviso …)' -> '… en 7 dptos. (incl. Lima y Piura) (aviso …)' cuando la lista es larga.
    Se priorizan los departamentos con sede de Pacífico."""
    def corto(m):
        deps = [x.strip() for x in re.split(r",\s*|\s+y\s+", m.group(2)) if x.strip()]
        if len(deps) <= maximo:
            return m.group(0)
        con_sede = [x for x in SEDES_DPTO if x in deps][:3]
        incl = f" (incl. {base._unir_y(con_sede)})" if con_sede else ""
        return f"{m.group(1)} {len(deps)} dptos.{incl}"
    return re.sub(r"(\ben)\s+((?:[A-ZÁÉÍÓÚÑ][\wáéíóúñ]+(?: (?:de )?[A-ZÁÉÍÓÚÑ][\wáéíóúñ]+)*)(?:(?:,\s*|\s+y\s+)"
                  r"[A-ZÁÉÍÓÚÑ][\wáéíóúñ]+(?: (?:de )?[A-ZÁÉÍÓÚÑ][\wáéíóúñ]+)*)+)", corto, texto)


# ---------------------------------------------------------------- tarjeta con el gráfico
def dibujar_proyeccion(slide, sn, x, y, w, h):
    B = base
    head = 0.38
    c = B.caja(slide, x, y, w, h, B.WHITE, B.P_LINE, B.MSO_SHAPE.ROUNDED_RECTANGLE, "Tarjeta siniestros")
    c.adjustments[0] = 0.07
    B.sombra_suave(c)
    hb = B.caja(slide, x, y, w, head, B.P_TINT, None, B.MSO_SHAPE.ROUND_2_SAME_RECTANGLE, "Tarjeta siniestros encabezado")
    hb.adjustments[0] = 0.32
    hb.adjustments[1] = 0.0
    B.texto(slide, x + 0.3, y, w * 0.75, head, "SINIESTROS POR LÍNEA", 12, B.P_BLUE_DK, "xb",
            anchor=B.MSO_ANCHOR.MIDDLE)
    B.texto(slide, x + w * 0.7, y, w * 0.3 - 0.25, head, f"En soles · TC {sn['tc']:.2f}", 8, B.P_GREY,
            align=B.PP_ALIGN.RIGHT, anchor=B.MSO_ANCHOR.MIDDLE)

    con_directo = sn["grafico"] == "B"
    C_RET = B.RGBColor(0x0A, 0x5F, 0xA8)      # retenido: azul oscuro
    C_DIR = B.RGBColor(0x8F, 0xC3, 0xEA)      # directo: azul claro (mismo tono)
    C_TRK = B.RGBColor(0xE4, 0xEA, 0xF2)      # proyección: riel gris
    x0, ancho = x + 0.3, w - 0.6
    alto_fila = (h - head - 0.12) / len(sn["lineas"])    # filas más separadas
    yy = y + head + 0.1
    for ln in sn["lineas"]:
        # fila 1: nombre + n.º de siniestros | proyección
        B.texto(slide, x0, yy, ancho * 0.6, 0.2,
                [[(ln["nombre"].upper(), "xb", B.P_INK),
                  (f"  {ln['cant']} siniestros" + "", False,
                   B.P_GREY)]], 9, nombre="Fila " + ln["nombre"])
        B.texto(slide, x0 + ancho * 0.6, yy, ancho * 0.4, 0.2,
                [[("Proyección ", False, B.P_GREY), (soles(ln["proy"]), True, B.P_INK)]], 9,
                align=B.PP_ALIGN.RIGHT)
        # barra: riel = proyección; relleno = directo (B) y retenido
        by, bh = yy + 0.21, 0.15
        trk = B.caja(slide, x0, by, ancho, bh, C_TRK, None, B.MSO_SHAPE.ROUNDED_RECTANGLE, "Proyeccion " + ln["nombre"])
        trk.adjustments[0] = 0.5
        if con_directo and ln["directo"] > 0:
            wd = max(0.05, ancho * min(1, ln["directo"] / ln["proy"]))
            bd = B.caja(slide, x0, by, wd, bh, C_DIR, None, B.MSO_SHAPE.ROUNDED_RECTANGLE, "Directo " + ln["nombre"])
            bd.adjustments[0] = 0.5
        if ln["retenido"] > 0:
            wr = max(0.05, ancho * min(1, ln["retenido"] / ln["proy"]))
            br = B.caja(slide, x0, by, wr, bh, C_RET, None, B.MSO_SHAPE.ROUNDED_RECTANGLE, "Retenido " + ln["nombre"])
            br.adjustments[0] = 0.5
        # fila 3: valores (texto en tinta; el cuadrado de color da la identidad)
        partes = [("■ ", True, C_RET), ("Retenido ", False, B.P_GREY), (soles(ln["retenido"]), True, B.P_INK),
                  (f" ({pct(ln['retenido'], ln['proy'])})", False, B.P_INK)]
        if con_directo:
            partes += [("     ■ ", True, C_DIR), ("Directo ", False, B.P_GREY), (soles(ln["directo"]), True, B.P_INK),
                       (f" ({pct(ln['directo'], ln['proy'])})", False, B.P_INK)]
        B.texto(slide, x0, by + bh + 0.03, ancho, 0.18, [partes], 8.5, nombre="Valores " + ln["nombre"])
        yy += alto_fila



# ---------------------------------------------------------------- arranque
def main():
    conf = leer_extra()
    leer_obs = base.leer_datos             # lo deja puesto obs.main(); aquí se envuelve

    def leer_con_siniestros(ruta):
        d = obs.leer_datos_observatorio(ruta)
        # Para hacer espacio a la tarjeta de siniestros, el pronóstico queda en lo esencial:
        # Costa, Sierra, niveles de lluvia y quebradas (sin "Acumulados previstos" ni "Otros avisos").
        d["p1"] = [resumir_listas(it) for it in d["p1"]
                   if it.split(":")[0].strip() in ("Costa", "Sierra", "Selva", "Quebradas")]   # los niveles ya están en el mapa
        d["calor"] = sorted((compactar_calor(resumir_listas(it)) for it in d.get("calor") or []),
                            key=lambda t: "dptos." not in t)     # el aviso que cubre más departamentos va primero
        for k in ("lluvias", "vias", "rios", "huaicos", "oleaje", "p1"):
            d[k] = [compactar(it) for it in d.get(k) or []]
        # Calor y oleaje van en la misma lista del pronóstico (un solo título = más espacio para todo):
        #   "Calor (costa y sierra): …", "Oleaje (Lima (Costa Verde)): …"
        base.repartir_adicionales(d)               # las alertas del txt (Costa Verde) entran antes de unir
        unidos = []
        for it in d.get("calor") or []:
            etq, _, txt = it.partition(":")
            unidos.append(f"Calor en {etq.strip().lower()}: {txt.strip()}")
        for it in d.get("oleaje") or []:
            etq, _, txt = it.partition(":")
            lugar = "el litoral" if etq.strip().lower() == "litoral" else etq.strip()
            unidos.append(f"Oleaje en {lugar}: {txt.strip()}")
        d["p1"] = d["p1"] + unidos
        d["calor"], d["oleaje"] = [], []
        d["avisos_derecha"] = True      # Estatus = hechos reportados; derecha = pronóstico + avisos (calor, oleaje)
        sn = cargar_siniestros_proyeccion(conf)
        if sn:
            d["siniestros"] = sn
            d["siniestros_alto"] = 1.8
            print(f"[OK] SINIESTROS ({sn['archivo']}) · TC {sn['tc']:.2f} ({sn['origen_tc']}) · gráfico {sn['grafico']}")
            for ln in sn["lineas"]:
                print(f"     {ln['nombre']}: {ln['cant']} siniestros · retenido {soles(ln['retenido'])} "
                      f"({pct(ln['retenido'], ln['proy'])}) · directo {soles(ln['directo'])} "
                      f"({pct(ln['directo'], ln['proy'])}) · proyección {soles(ln['proy'])}")
        return d
    del leer_obs

    orig_conf = base.leer_configuracion

    def conf_propuesta(carpeta):
        c = orig_conf(carpeta)
        c.update({"salida": SALIDA, "modo": "ninguno", "siniestros": "no"})   # nada a OneDrive ni correo
        return c

    os.makedirs(SALIDA, exist_ok=True)
    base.leer_configuracion = conf_propuesta
    base.dibujar_siniestros = dibujar_proyeccion
    base.VERSION = f"{VERSION_PROP} · {obs.VERSION_OBS} · diseño {base.VERSION}"
    if len(sys.argv) == 1 or sys.argv[1].startswith("-"):
        sys.argv.insert(1, os.path.join(CARPETA, "mensaje_observatorio.txt"))
    base.leer_datos = leer_con_siniestros
    base.main()


if __name__ == "__main__":
    main()
