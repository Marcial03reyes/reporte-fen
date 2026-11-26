"""
Reporte FEN · PROPUESTA con siniestros por línea (real vs. estimado)
====================================================================
Mismo reporte que "generar reporte observatorio.py" (clima desde el Observatorio, Pacífico desde el txt)
más una tarjeta "SINIESTROS POR LÍNEA". Para cada línea de seguro hay dos barras, en soles:

  • IMPACTO NETO DE REASEGURO: lo real (monto RETENIDO del Excel) vs. lo estimado (se digita abajo).
  • IMPACTO BRUTO:             lo real (monto DIRECTO del Excel)  vs. lo estimado (se digita abajo).

La barra gris completa es el estimado y la parte pintada es cuánto va avanzado lo real.

  • Excel: siniestros.xlsx o data.xlsx en esta carpeta (columnas Linea, MontoDirectoUSD, MontoRetencionUSD).
  • Los ESTIMADOS, el TIPO DE CAMBIO y las OPCIONES se llenan a mano más abajo, en el bloque
    "LLENAR A MANO" (no salen de ninguna fuente).

NO envía nada a OneDrive ni por correo: el PPT y el PNG quedan en la carpeta "reporte_generado_propuesta".

Uso:  python "generar reporte propuesta.py"          (usa mensaje_observatorio.txt; genera la opción 1 y la 2)
Necesita en la misma carpeta: "generar reporte observatorio.py" y "generar reporte fen.py".
"""
import importlib.util
import os
import re
import sys
import warnings

VERSION_PROP = "2026-11-26 (opción 1 = total; bruto arriba, neto abajo)"
CARPETA = os.path.dirname(os.path.abspath(__file__))
SALIDA = os.path.join(CARPETA, "reporte_generado_propuesta")
TC_DEF = 3.55

# ╔════════════════════════════════════════════════════════════════════════════════════════════════╗
# ║                                 LLENAR A MANO  (en soles)                                       ║
# ╚════════════════════════════════════════════════════════════════════════════════════════════════╝
#
# 1) ESTIMADOS POR LÍNEA DE SEGURO (en soles)
#    Cada fila es:
#      ("NOMBRE EN EL REPORTE",  estimado NETO de reaseguro,  estimado BRUTO,  [cómo se llama en la columna "Linea" del Excel])
#      • Los montos van sin "S/" ni comas. El guion bajo solo separa los miles: 23_600_000 = 23,600,000.
#      • NETO de reaseguro se compara con el monto RETENIDO del Excel; BRUTO, con el monto DIRECTO.
#      • Una línea puede sumar varias del Excel: PROPIEDAD = Vivienda + Negocios.
#      • El orden de la lista es el orden en el reporte.
#      • Una fila con "#" adelante NO sale en el reporte (ni en el gráfico).
#        Para activar VIDA o AUTOS: borra el "#", pon sus estimados y revisa que el nombre entre corchetes
#        sea el de la columna "Linea" del Excel (mayúsculas y tildes no importan).
LINEAS = [
    #  NOMBRE          NETO           BRUTO          LÍNEAS DEL EXCEL
    # ("VIDA",                  0,             0,  ["vida"]),
    # ("AUTOS",                 0,             0,  ["autos", "vehiculos"]),
    ("PROPIEDAD",      11_900_000,   202_500_000,  ["vivienda", "negocios"]),
    ("EMPRESAS",       23_600_000,   440_100_000,  ["comerciales"]),
]

# 2) TIPO DE CAMBIO para pasar a soles los montos en dólares del Excel.
#    Para usar el mismo que trae el Excel, escribe  TIPO_CAMBIO = None
TIPO_CAMBIO = 3.50

# 3) OPCIONES: cada vez que corres el script se genera un PPT y un PNG por cada opción de esta lista.
#    "1" = OPCIÓN 1: TOTAL de todas las líneas (suma), con impacto bruto y neto  → ..._Opcion1.pptx / .png
#    "2" = OPCIÓN 2: por LÍNEA de seguro, con impacto bruto y neto              → ..._Opcion2.pptx / .png
#    En ambas va primero la barra de impacto bruto y debajo la de impacto neto de reaseguro.
#    Cuando decidan, deja solo la elegida, por ejemplo  OPCIONES = ["2"]  (y entonces sale sin el sufijo).
#    Si se quedan con la 2, también puedes borrar en el código el bloque marcado
#    >>> OPCIÓN 1 ... <<< FIN OPCIÓN 1.
OPCIONES = ["1", "2"]
# ══════════════════════════════════════════════════════════════════════════════════════════════════


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
def _norm(t):
    import unicodedata
    t = unicodedata.normalize("NFD", str(t or "").strip().lower())
    return "".join(ch for ch in t if unicodedata.category(ch) != "Mn")


_OPCION_ACTUAL = ["2"]          # la cambia main() en cada vuelta


def leer_extra():
    # solo las filas de LINEAS sin "#" y con algún estimado > 0; las demás no salen ni en el gráfico
    activas = [(nom, float(neto or 0), float(bruto or 0), {_norm(x) for x in excel})
               for nom, neto, bruto, excel in LINEAS if (neto or 0) > 0 or (bruto or 0) > 0]
    return {"tc": TIPO_CAMBIO, "lineas": activas, "grafico": "B" if _OPCION_ACTUAL[0] == "2" else "A"}


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
    if not conf["lineas"]:
        print("  [!] No hay líneas con proyección en LINEAS: el reporte sale sin la tarjeta.")
        return None
    grupos = {nom: [0, 0.0, 0.0] for nom, _, _, _ in conf["lineas"]}
    a_grupo = {x: nom for nom, _, _, excel in conf["lineas"] for x in excel}
    tcs = []
    num = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
    for f in filas[1:]:
        if not f or len(f) <= max(i_lin, i_dir, i_ret) or not f[i_lin]:
            continue                                     # fila de totales / filtros
        g = a_grupo.get(_norm(f[i_lin]))
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
    origen_tc = "fijado en el script" if conf["tc"] else ("detectado en el Excel" if tcs else "por defecto")
    res = {"archivo": os.path.basename(ruta), "tc": tc, "origen_tc": origen_tc, "grafico": conf["grafico"],
           "lineas": []}
    for nombre, neto_est, bruto_est, _ in conf["lineas"]:
        cant, d_usd, r_usd = grupos[nombre]
        res["lineas"].append({"nombre": nombre, "cant": cant,
                              "neto": r_usd * tc, "neto_est": neto_est,        # retenido vs. estimado neto
                              "bruto": d_usd * tc, "bruto_est": bruto_est})    # directo  vs. estimado bruto
    return res


def soles(v):
    if v >= 1e8:
        return f"S/ {v / 1e6:,.1f} MM".replace(".0 MM", " MM")
    if v >= 1e6:
        return f"S/ {v / 1e6:.1f} MM".replace(".0 MM", " MM")
    if v >= 1e3:
        return f"S/ {v / 1e3:,.0f} mil"
    return f"S/ {v:,.0f}"


def pct(a, b):
    p_ = 100 * a / b if b else 0
    if p_ >= 1:
        return f"{p_:.1f}%"
    return f"{p_:.2f}%" if p_ >= 0.01 or p_ == 0 else "<0.01%"


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
    if txt.count("(") > txt.count(")"):                  # no cortar dentro de un paréntesis
        corte = max(txt.rfind(";", 0, txt.rfind("(")), txt.rfind(",", 0, txt.rfind("(")))
        txt = (txt[:corte] if corte > 40 else txt[:txt.rfind("(")]).rstrip(" ,;") + "."
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
C_NETO = (0x0A, 0x5F, 0xA8)     # impacto neto de reaseguro: azul oscuro
C_BRUTO = (0x4F, 0xA3, 0xDB)    # impacto bruto: azul medio (mismo tono, más claro)
C_RIEL = (0xE4, 0xEA, 0xF2)     # estimado: riel gris


def medidas(sn):
    """Las barras de cada fila (en las dos opciones): primero bruto, debajo neto de reaseguro."""
    return [("Impacto bruto", "bruto", "bruto_est", C_BRUTO),
            ("Impacto neto de reaseguro", "neto", "neto_est", C_NETO)]


def total_de(sn):
    """OPCIÓN 1: una sola fila con la suma de todas las líneas activas."""
    ls = sn["lineas"]
    return [{"nombre": "TOTAL", "detalle": " + ".join(l_["nombre"].capitalize() for l_ in ls),
             "cant": sum(l_["cant"] for l_ in ls),
             **{k: sum(l_[k] for l_ in ls) for k in ("bruto", "bruto_est", "neto", "neto_est")}}]


def alto_tarjeta(sn):
    n, m = len(sn["lineas"]), len(medidas(sn))
    por_linea = (0.21 + 0.31 * m + 0.07) if n <= 2 else (0.2 + 0.19 * m + 0.05)
    return 0.38 + 0.12 + 0.06 + n * por_linea


def _barra(slide, x, y, w, h, real, est, color, nombre):
    B = base
    riel = B.caja(slide, x, y, w, h, B.RGBColor(*C_RIEL), None, B.MSO_SHAPE.ROUNDED_RECTANGLE, "Estimado " + nombre)
    riel.adjustments[0] = 0.5
    if real > 0 and est > 0:
        b = B.caja(slide, x, y, max(0.06, w * min(1, real / est)), h, B.RGBColor(*color), None,
                   B.MSO_SHAPE.ROUNDED_RECTANGLE, "Real " + nombre)
        b.adjustments[0] = 0.5


def _valor(ln, clave, clave_est):
    """'S/ 76 mil de S/ 11.9 MM · 0.64%'  (texto en tinta; el color lo lleva la barra)"""
    B = base
    real, est = ln[clave], ln[clave_est]
    if not est:
        return [(soles(real), True, B.P_INK), ("  sin estimado", False, B.P_GREY)]
    return [(soles(real), True, B.P_INK), (" de " + soles(est), False, B.P_GREY), (" · " + pct(real, est), True, B.P_INK)]


def dibujar_proyeccion(slide, sn, x, y, w, h):
    B = base
    head = 0.38
    c = B.caja(slide, x, y, w, h, B.WHITE, B.P_LINE, B.MSO_SHAPE.ROUNDED_RECTANGLE, "Tarjeta siniestros")
    c.adjustments[0] = 0.07
    B.sombra_suave(c)
    hb = B.caja(slide, x, y, w, head, B.P_TINT, None, B.MSO_SHAPE.ROUND_2_SAME_RECTANGLE, "Tarjeta siniestros encabezado")
    hb.adjustments[0] = 0.32
    hb.adjustments[1] = 0.0
    B.texto(slide, x + 0.3, y, w * 0.75, head, "SINIESTROS TOTALES" if sn.get("por_linea") else "SINIESTROS POR LÍNEA", 12, B.P_BLUE_DK, "xb",
            anchor=B.MSO_ANCHOR.MIDDLE)
    B.texto(slide, x + w * 0.55, y, w * 0.45 - 0.25, head, f"En soles · TC {sn['tc']:.2f}", 8,
            B.P_GREY, align=B.PP_ALIGN.RIGHT, anchor=B.MSO_ANCHOR.MIDDLE)

    x0, ancho = x + 0.3, w - 0.6
    ms = medidas(sn)
    compacta = len(sn["lineas"]) > 2           # 3 o 4 líneas: etiqueta, barra y monto en un solo renglón
    alto_linea = (h - head - 0.18) / len(sn["lineas"])
    yy = y + head + 0.1
    for ln in sn["lineas"]:
        B.texto(slide, x0, yy, ancho, 0.2,
                [[(ln["nombre"].upper(), "xb", B.P_INK),
                  (f"  {ln['cant']} siniestros" + (f" · {ln['detalle']}" if ln.get("detalle") else ""), False, B.P_GREY)]],
                8.5 if compacta else 9, nombre="Fila " + ln["nombre"])
        sy = yy + (0.2 if compacta else 0.22)
        for etiqueta, clave, clave_est, color in ms:
            nom = f"{etiqueta} {ln['nombre']}"
            if compacta:
                corta = "Neto de reaseguro" if clave == "neto" else "Bruto"
                B.texto(slide, x0, sy, 1.15, 0.18, [[("■ ", True, B.RGBColor(*color)), (corta, False, B.P_GREY)]], 7.5,
                        anchor=B.MSO_ANCHOR.MIDDLE)
                ab = ancho - 1.2 - 1.75
                _barra(slide, x0 + 1.2, sy + 0.045, ab, 0.09, ln[clave], ln[clave_est], color, nom)
                B.texto(slide, x0 + ancho - 1.7, sy, 1.7, 0.18, [_valor(ln, clave, clave_est)], 7.5,
                        align=B.PP_ALIGN.RIGHT, anchor=B.MSO_ANCHOR.MIDDLE)
                sy += 0.19
            else:
                B.texto(slide, x0, sy, ancho * 0.5, 0.17, [[(etiqueta, False, B.P_GREY)]], 8)
                B.texto(slide, x0 + ancho * 0.4, sy, ancho * 0.6, 0.17, [_valor(ln, clave, clave_est)], 8,
                        align=B.PP_ALIGN.RIGHT)
                _barra(slide, x0, sy + 0.18, ancho, 0.1, ln[clave], ln[clave_est], color, nom)
                sy += 0.31
        yy += alto_linea


# ---------------------------------------------------------------- arranque
def main():
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
        # con dos barras por línea la tarjeta de siniestros crece: el pronóstico se acorta a lo esencial
        d["p1"] = [compactar(it, 115) for it in d["p1"] + unidos]
        d["calor"], d["oleaje"] = [], []
        d["avisos_derecha"] = True      # Estatus = hechos reportados; derecha = pronóstico + avisos (calor, oleaje)
        sn = cargar_siniestros_proyeccion(leer_extra())     # se lee en cada opción
        if sn:
            # >>> OPCIÓN 1: todo sumado en una fila TOTAL. Borrar este bloque si se quedan con la opción 2.
            if sn["grafico"] == "A":
                sn["por_linea"] = sn["lineas"]
                sn["lineas"] = total_de(sn)
            # <<< FIN OPCIÓN 1
            d["siniestros"] = sn
            d["siniestros_alto"] = alto_tarjeta(sn)
            print(f"[OK] SINIESTROS ({sn['archivo']}) · TC {sn['tc']:.2f} ({sn['origen_tc']}) · opción {'2' if sn['grafico'] == 'B' else '1'}")
            for ln in sn.get("por_linea", []) + sn["lineas"]:
                print(f"     {ln['nombre']}: {ln['cant']} siniestros · neto (retenido) {soles(ln['neto'])} de "
                      f"{soles(ln['neto_est'])} ({pct(ln['neto'], ln['neto_est'])}) · bruto (directo) "
                      f"{soles(ln['bruto'])} de {soles(ln['bruto_est'])} ({pct(ln['bruto'], ln['bruto_est'])})")
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

    opciones = [o for o in (str(x).strip() for x in OPCIONES) if o in ("1", "2")] or ["2"]
    if os.environ.get("FEN_GRAFICO"):                       # solo para pruebas
        opciones = ["1" if os.environ["FEN_GRAFICO"].upper().startswith(("A", "1")) else "2"]
    # el mismo número de reporte para todas las opciones (si NUMERO está vacío en el txt)
    sig_orig, num = base.siguiente_numero, []
    base.siguiente_numero = lambda carpeta: num[0] if num else (num.append(sig_orig(carpeta)) or num[0])
    abrir = getattr(os, "startfile", None)
    for i, op in enumerate(opciones):
        _OPCION_ACTUAL[0] = op
        if len(opciones) > 1:
            print(f"\n================ OPCIÓN {op} ================")
        antes = {f: os.path.getmtime(os.path.join(SALIDA, f)) for f in os.listdir(SALIDA)}
        if abrir and i < len(opciones) - 1:
            os.startfile = lambda *a, **k: None             # la carpeta se abre una sola vez, al final
        try:
            base.main()
        finally:
            if abrir:
                os.startfile = abrir
        if len(opciones) > 1:                                # agrega el sufijo _Opcion1 / _Opcion2
            for f in os.listdir(SALIDA):
                nuevo = f not in antes or os.path.getmtime(os.path.join(SALIDA, f)) != antes[f]
                if nuevo and f.startswith("Reporte_FEN_") and "_Opcion" not in f and f.endswith((".pptx", ".png")):
                    raiz, ext = os.path.splitext(f)
                    os.replace(os.path.join(SALIDA, f), os.path.join(SALIDA, f"{raiz}_Opcion{op}{ext}"))
                    print(f"  → {raiz}_Opcion{op}{ext}")


if __name__ == "__main__":
    main()
