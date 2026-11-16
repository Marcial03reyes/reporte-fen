"""
Reporte FEN · versión Observatorio
==================================
Arma el mismo Reporte FEN (PPT + PNG, mismo diseño y mismo envío por Power Automate) pero con TODA
la información climática tomada del Observatorio Pacífico (snapshot.json):

  • Lluvias reportadas, Ríos y quebradas, Huaicos, Vías ........ publicaciones de cuentas OFICIALES en X
                                                                 (SENAMHI, COEN/INDECI, MTC, Provías, ANA…)
  • Pronóstico de hoy ........................................... SENAMHI (pronóstico en X + aviso de lluvias)
                                                                 y otros avisos vigentes (calor, viento, oleaje)
  • Mapa ........................................................ avisos oficiales de lluvias y quebradas

De WhatsApp solo se usa lo de Pacífico: confirmaciones, mensaje de agencias (sedes) y colaboradores.
Se descarta todo lo que el Observatorio marca como prensa "pendiente de corroboración", las cuentas que no son
institucionales y los textos que dicen "pendiente", "por confirmar", etc.

Uso:  python "generar reporte observatorio.py"            (usa mensaje_observatorio.txt)
      python "generar reporte observatorio.py" otro.txt
Necesita "generar reporte fen.py" en la misma carpeta (de ahí toma el diseño, el mapa y el envío).
"""
import importlib.util
import os
import re
import sys
from datetime import datetime, timedelta, timezone

VERSION_OBS = "2026-11-16c (Observatorio)"
CARPETA = os.path.dirname(os.path.abspath(__file__))
LIMA = timezone(timedelta(hours=-5))

# Cuentas institucionales cuyas publicaciones en X se aceptan como hechos (el resto se ignora)
CUENTAS_OFICIALES = {
    "@senamhiperu": "SENAMHI", "@senamhidz05": "SENAMHI", "@coenperu": "COEN", "@indeciperu": "INDECI",
    "@mtc_gobperu": "MTC", "@proviasnac_peru": "Provías", "@anaperu": "ANA", "@bomberos_pe": "Bomberos",
    "@coercallao": "COER Callao", "@naval_peru": "Marina", "@enfenperu": "ENFEN",
}
# Frases que indican que el propio texto no está confirmado
NO_CONFIRMADO = re.compile(r"pendiente|por confirmar|sin confirmar|por corroborar|no confirmad|se viene verificando|"
                           r"en evaluaci[oó]n preliminar|presunt|habr[ií]a\b", re.I)
# Publicaciones que no son hechos del día (campañas, notas de prensa, actividades)
NO_HECHO = re.compile(r"nota de prensa|prep[aá]rate|recomienda|recomendaciones|campa[ñn]a|simulacro|papa\b|le[oó]n xiv|"
                      r"visita|#cuidatupiel|conoce c[oó]mo|infograf", re.I)
PRONOSTICO = re.compile(r"^\s*se prev[eé]n?|contin[uú]an? las lluvias|\(\d{1,2}:\d{2} a \d{1,2}:\d{2} ?h\)|"
                        r"toma tus precauciones|pr[oó]ximas \d* ?horas|#pron[oó]stico|#aviso\b|se estima|podr[ií]an? alcanzar", re.I)
TEMPERATURA = re.compile(r"°\s*c|temperatura|calor|radiaci[oó]n", re.I)

SECCION_DE = [  # prioridad: el primer patrón que calce decide la sección
    ("huaicos", re.compile(r"huaico|deslizamiento|derrumbe|movimiento en masa|ca[ií]da de rocas", re.I)),
    ("rios", re.compile(r"quebrada|\br[ií]os?\b|caudal|desborde|inundaci|crecida", re.I)),
    ("vias", re.compile(r"(\bv[ií]as?\b|carretera|tr[aá]nsito|puente|tramo).{0,80}(interrump|restring|bloque|cerrad|"
                        r"intransitable|afectad)|(interrump|restring|bloque|cerrad).{0,60}(\bv[ií]as?\b|carretera|tr[aá]nsito)",
                        re.I)),
    ("lluvias", re.compile(r"lluvi|precipitaci|acumul|\bmm\b|tormenta|chubasco|granizo|nevada", re.I)),
]
MAX_POR_SECCION = 5
MINUSCULA = {"huaico", "huaicos", "lluvias", "lluvia", "clima", "tiempo", "deslizamiento", "inundacion"}
META = {"ahora", "minam", "senamhi", "eldato", "clima", "pronostico", "pronóstico", "aviso", "perúresponde",
        "peruresponde", "preparadosantelluviasintensas", "preparadosanteelfen", "atención", "atencion", "coen"}


def cargar_base():
    ruta = os.path.join(CARPETA, "generar reporte fen.py")
    if not os.path.exists(ruta):
        sys.exit(f"Falta 'generar reporte fen.py' junto a este script ({CARPETA}). Copia ambos a la misma carpeta.")
    spec = importlib.util.spec_from_file_location("reporte_fen_base", ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


base = cargar_base()
leer_datos_whatsapp = base.leer_datos          # lector original (para la parte de Pacífico del txt)


# ---------------------------------------------------------------- utilidades de texto
def a_lima(iso):
    try:
        t = datetime.fromisoformat((iso or "").replace("Z", "+00:00"))
    except ValueError:
        return None
    return t.astimezone(LIMA) if t.tzinfo else t.replace(tzinfo=LIMA)


def _hashtag(m):
    palabra = m.group(1)
    if palabra.lower() in META:
        return ""
    partes = re.findall(r"[A-ZÁÉÍÓÚÑ]?[a-záéíóúñ]+|[A-ZÁÉÍÓÚÑ]+(?![a-z])|\d+", palabra)
    if len(partes) > 1 and palabra.lower() not in MINUSCULA:      # #LluviasIntensas -> lluvias intensas
        return " ".join(p.lower() for p in partes)
    return palabra.lower() if palabra.lower() in MINUSCULA else palabra


def limpiar_post(texto):
    t = re.sub(r"https?://\S+", " ", texto)
    t = re.sub(r"[\U0001F000-\U0001FFFF☀-➿⬀-⯿️‍]", " ", t)
    t = re.sub(r"#(\w+)", _hashtag, t)
    t = re.sub(r"@\w+", "", t)
    t = re.sub(r"\s+", " ", t).strip(" |:-–")
    t = re.sub(r"^(\|\s*)+", "", t).strip()
    return t


def recortar(texto, maximo=260):
    """Hasta la última oración completa (las publicaciones llegan cortadas) y sin pasar de 'maximo'."""
    t = texto.strip()
    # fin de oración: un punto que no cierra una abreviatura de una letra ("C. P.") y va seguido de espacio
    fines = [m.end() for m in re.finditer(r"(?<!\b[A-ZÁÉÍÓÚÑ])\.(?=\s|$)", t)]
    if not t.endswith((".", "!", ")")):
        ok = [f for f in fines if f > 40]
        t = t[:ok[-1]] if ok else t.rstrip(",;: ") + "."
    if len(t) > maximo:
        ok = [f for f in fines if 60 < f <= maximo]
        t = t[:ok[-1]] if ok else t[:maximo].rsplit(" ", 1)[0].rstrip(",;:") + "…"
    return t[:1].upper() + t[1:]


def quitar_previsiones(texto):
    """Dentro de un hecho, quita las oraciones de pronóstico ('Se prevé que las lluvias continúen…')."""
    frases = re.split(r"(?<=[.!])\s+", texto)
    quedan = [f for f in frases if not re.match(r"^(se prev[eé]n?|se espera|se esperan|podr[ií]a)", f.strip(), re.I)]
    return " ".join(quedan) if quedan else texto


def etiqueta(texto, deps):
    """Lugar del hecho: el que encabeza el texto ('Piura | …', 'Moquegua: …') o los departamentos nombrados."""
    m = re.match(r"^\s*([A-ZÁÉÍÓÚÑ][\wáéíóúñ]+(?: [A-ZÁÉÍÓÚÑ][\wáéíóúñ]+)?)\s*[:|]\s*(.+)$", texto)
    if m and len(m.group(1)) <= 25:
        return m.group(1), m.group(2)
    nombrados = [d for d in deps if re.search(r"\b" + re.escape(d) + r"\b(?! Libre)", texto)] or list(deps)
    return (" y ".join([", ".join(nombrados[:-1]), nombrados[-1]]) if len(nombrados) > 1 else
            (nombrados[0] if nombrados else "")), texto


# ---------------------------------------------------------------- extracción del snapshot
def hechos_oficiales(snap, desde, hasta):
    """Publicaciones institucionales en X dentro de la ventana, clasificadas por sección."""
    secciones = {k: [] for k, _ in SECCION_DE}
    descartes = {"no oficial": 0, "fuera de ventana": 0, "pronóstico": 0, "no es un hecho": 0, "no confirmado": 0}
    for x in (snap.get("datasets") or {}).get("senales_x") or []:
        cuenta = (x.get("cuenta") or "").lower()
        if cuenta not in CUENTAS_OFICIALES:
            descartes["no oficial"] += 1
            continue
        t = a_lima(x.get("publicado"))
        if not t or not (desde <= t <= hasta):
            descartes["fuera de ventana"] += 1
            continue
        crudo = x.get("resumen") or ""
        if NO_CONFIRMADO.search(crudo):
            descartes["no confirmado"] += 1
            continue
        if NO_HECHO.search(crudo):
            descartes["no es un hecho"] += 1
            continue
        texto = limpiar_post(crudo)
        if PRONOSTICO.search(crudo) or PRONOSTICO.search(texto):
            descartes["pronóstico"] += 1
            continue
        seccion = next((k for k, rx in SECCION_DE if rx.search(texto)), None)
        if not seccion or (seccion == "lluvias" and TEMPERATURA.search(texto) and not re.search(r"\bmm\b|lluvios", texto)):
            descartes["no es un hecho"] += 1
            continue
        lugar, cuerpo = etiqueta(texto, x.get("departamentos") or [])
        cuerpo = quitar_previsiones(cuerpo)
        deps = {base._sin_tildes(x_) for x_ in (x.get("departamentos") or [])} or {base._sin_tildes(lugar)}
        secciones[seccion].append({"t": t, "cuenta": CUENTAS_OFICIALES[cuenta], "lugar": lugar, "deps": deps,
                                   "texto": recortar(cuerpo)})
    # Una quebrada o huaico que corta una vía también va (corto) a Vías restringidas
    for k in ("rios", "huaicos"):
        for h in list(secciones[k]):
            m = re.search(r"(interrump|restring|bloque|cerrad)\w*\s+(?:el tr[aá]nsito en\s+)?(?:la |el )?"
                          r"((?:v[ií]a|carretera|tramo|puente)\s+[^,.;]+)", h["texto"], re.I)
            if m:
                estado = {"interrump": "interrumpido", "restring": "restringido", "bloque": "bloqueado",
                          "cerrad": "cerrado"}[m.group(1).lower()]
                secciones["vias"].append({**h, "texto": f"Tránsito {estado} en la {m.group(2).strip()} "
                                                        f"({'por quebradas' if k == 'rios' else 'por huaico'})."})
    # Lluvias: SENAMHI publica varios balances al día; queda el más reciente de cada departamento
    # (una publicación más antigua se descarta si todos sus departamentos ya salen en una más nueva).
    # Ríos, huaicos y vías: cada publicación es un hecho distinto; solo se quitan textos repetidos.
    for k, lista in secciones.items():
        lista.sort(key=lambda h: h["t"], reverse=True)
        cubiertos, textos, unicos = set(), set(), []
        for h in lista:
            if k == "lluvias" and h["deps"] <= cubiertos:
                continue
            clave_txt = re.sub(r"\W+", "", h["texto"].lower())[:120]
            if clave_txt in textos:
                continue
            cubiertos |= h["deps"]
            textos.add(clave_txt)
            unicos.append(h)
        secciones[k] = unicos[:MAX_POR_SECCION]
    return secciones, descartes


def pronostico(snap, dia, desde):
    """Costa / Sierra / Selva del último pronóstico de SENAMHI en X; si no hay, el aviso de lluvias del día.
    Más los otros avisos vigentes (calor, viento, oleaje)."""
    items = []
    posts = [x for x in (snap.get("datasets") or {}).get("senales_x") or []
             if (x.get("cuenta") or "").lower() == "@senamhiperu" and re.search(r"#pron[oó]stico", x.get("resumen") or "", re.I)
             and (a_lima(x.get("publicado")) or desde) >= desde - timedelta(hours=12)]
    por_region = {}

    def region_de(frase):
        f = frase.lower()
        return "Sierra" if "sierra" in f else "Selva" if "selva" in f else "Costa"

    def limpia(frase):
        f = re.sub(r"^(asimismo|adem[aá]s|por otro lado),?\s*", "", frase.strip(), flags=re.I)
        f = re.sub(r"^el senamhi informa que,?\s*", "", f, flags=re.I)
        f = re.sub(r"^desde el \w+ \d+ hasta el \w+ \d+ de \w+,?\s*", "", f, flags=re.I)
        return f[:1].upper() + f[1:]

    if posts:
        ult = max(posts, key=lambda x: x.get("publicado") or "")
        txt = limpiar_post(ult["resumen"])
        etiquetas = list(re.finditer(r"\b(Costa|Sierra|Selva)\s*:\s*(.+?)(?=\s*\b(?:Costa|Sierra|Selva)\s*:|$)", txt, re.I))
        if etiquetas:
            for m in etiquetas:
                por_region.setdefault(m.group(1).capitalize(), recortar(m.group(2).strip(), 160))
        else:                                   # texto libre: cada oración va a su región
            txt = re.sub(r"^.*?\|\s*", "", txt)
            for f in re.split(r"(?<=\.)\s+", txt):
                if len(f) > 20:
                    r_ = region_de(f)
                    por_region[r_] = (por_region.get(r_, "") + " " + recortar(limpia(f), 160)).strip()
    filas = [r for r in (snap.get("sources") or {}).get("senamhi", {}).get("rows", [])
             if "lluvias" in (r.get("eventIds") or []) and r.get("jornada") == dia]
    if filas:
        ev = filas[0].get("evento") or ""
        n = filas[0].get("numero_aviso")
        cuerpo = ev.split(". ", 1)[1] if ". " in ev else ev
        cuerpo = re.sub(r"^el \w+ \d+ de \w+,?\s*", "", cuerpo, flags=re.I)
        desc = filas[0].get("descripcion_oficial") or ""
        for f in re.split(r"(?<=\.)\s+", desc):           # lo que el post no cubre, del aviso oficial
            reg = region_de(f)
            if reg not in por_region and re.search(reg, f, re.I):
                por_region[reg] = recortar(limpia(f), 170)
    for reg in ("Costa", "Sierra", "Selva"):
        if reg in por_region:
            items.append(f"{reg}: {por_region[reg]}")
    if filas:
        items.append(f"Acumulados previstos (aviso SENAMHI {n}): " + recortar(cuerpo, 170))
    # Departamentos por nivel oficial, por peligro (como "Pronóstico: nivel … hoy" de la página)
    filas_dia = [r for r in (snap.get("sources") or {}).get("senamhi", {}).get("rows", [])
                 if r.get("jornada") == dia and not r.get("solo_contexto") and r.get("nivel") in base.NIVEL_ORDEN]
    for ev, nombre in (("lluvias", "Lluvias"), ("calor", "Calor")):
        dep, avisos = {}, set()
        for r in filas_dia:
            if ev in (r.get("eventIds") or []):
                avisos.add(r.get("numero_aviso"))
                for d_ in r.get("departamentos") or []:
                    if base.NIVEL_ORDEN[r["nivel"]] > base.NIVEL_ORDEN.get(dep.get(d_), 0):
                        dep[d_] = r["nivel"]
        if not dep:
            continue
        partes = []
        for nv in ("Rojo", "Naranja", "Amarillo"):
            lista = sorted((d_ for d_, v in dep.items() if v == nv), key=base._sin_tildes)
            if lista:
                partes.append(f"{base.INTENSIDAD[nv].lower()} ({nv.lower()}) en " + base._unir_y(lista))
        num = " y ".join(str(a_) for a_ in sorted(a_ for a_ in avisos if a_))
        items.append(f"{nombre}: nivel " + "; ".join(partes) + f" (aviso{'s' if len(avisos) > 1 else ''} SENAMHI {num}).")
    qn = (snap.get("sources") or {}).get("indeci", {}).get("quebradaNotice") or {}
    fin_q = a_lima(qn.get("fin"))
    if qn.get("number") and fin_q and fin_q.strftime("%Y-%m-%d") >= dia:
        hora = f"{fin_q:%H} h" if fin_q.minute == 0 else f"{fin_q:%H}.{fin_q:%M} h"
        items.append(f"Quebradas: activación de quebradas en nivel {(qn.get('level') or '').lower()} "
                     f"(aviso {qn['number']}), vigente hasta las {hora} del {fin_q:%d/%m}.")
    otros = []
    for no in (snap.get("sources") or {}).get("senamhi", {}).get("notices", []):
        ini, fin = a_lima(no.get("start")), a_lima(no.get("end"))
        ev_ = no.get("eventIds") or []
        if "lluvias" in ev_ or "calor" in ev_ or not ini or not fin:
            continue
        if ini.strftime("%Y-%m-%d") <= dia <= fin.strftime("%Y-%m-%d"):
            tit = re.sub(r"\s*\(.*?\)", "", no.get("title") or "").strip().capitalize()
            tit = re.sub(r"^Incremento de ", "", tit).capitalize()
            otros.append(f"{tit} ({no.get('number')}, hasta {fin:%d/%m})")
    for no in (snap.get("sources") or {}).get("dhn", {}).get("notices", []):
        ini, fin = a_lima(no.get("inicio")), a_lima(no.get("fin"))
        if ini and fin and ini.strftime("%Y-%m-%d") <= dia <= fin.strftime("%Y-%m-%d"):
            otros.append(f"Oleaje anómalo en el litoral (DHN {no.get('number')}, hasta {fin:%d/%m})")
    if otros:
        items.append("Otros avisos: " + "; ".join(otros) + ".")
    return items


# ---------------------------------------------------------------- lectura del txt + snapshot
def leer_datos_observatorio(ruta):
    d = _leer_txt_pacifico(ruta)
    snap, origen = base.cargar_radar([CARPETA])
    if not snap:
        sys.exit("[X] Sin datos del Observatorio: " + origen + ".\n"
                 f"    Guarda el snapshot.json de hoy en {CARPETA} y vuelve a ejecutar.")
    capt = a_lima(snap.get("at")) or datetime.now(LIMA)
    if not d["fecha"]:
        d["fecha"] = capt.strftime("%d/%m/%Y")
    dd, mm, aa = re.findall(r"\d+", d["fecha"])[:3]
    dia = f"{int(aa) if len(aa) == 4 else 2000 + int(aa)}-{int(mm):02d}-{int(dd):02d}"
    hasta = capt
    desde = datetime.strptime(dia, "%Y-%m-%d").replace(tzinfo=LIMA) - timedelta(hours=18)   # desde las 6 a. m. del día anterior
    d["horario"] = f"Corte: {capt:%H:%M} horas"

    secciones, descartes = hechos_oficiales(snap, desde, hasta)
    for k, lista in secciones.items():
        d[k] = [(f"{h['lugar']}: " if h["lugar"] else "") + h["texto"] for h in lista]
    d["p1_titulo"] = "HOY"
    d["p1"] = pronostico(snap, dia, desde)
    d["nota"] = d.get("nota") or "Se mantiene el monitoreo."
    base.sintetizar(d)

    print("-" * 70)
    print(f"[OK] OBSERVATORIO: datos del {capt:%d/%m %H:%M} · hechos de {desde:%d/%m %H:%M} a {hasta:%d/%m %H:%M}")
    for k, nombre in (("lluvias", "Lluvias"), ("rios", "Ríos y quebradas"), ("huaicos", "Huaicos"), ("vias", "Vías")):
        print(f"  {nombre}: {len(d[k])}  " + " · ".join(f"{h['cuenta']} {h['t']:%d/%m %H:%M}" for h in secciones[k]))
    print("  Descartado: " + ", ".join(f"{v} {k}" for k, v in descartes.items() if v))
    qn = (snap.get("sources") or {}).get("indeci", {}).get("quebradaNotice") or {}
    areas = ((snap.get("sources") or {}).get("indeci", {}).get("quebradaAreas") or {}).get("areas") or []
    fin_q = a_lima(qn.get("fin"))
    if qn.get("number") and fin_q and fin_q >= capt and not areas and not d.get("aviso_dptos"):
        print(f"  [!] Aviso de quebradas {qn['number']} ({(qn.get('level') or '').lower()}) vigente hasta {fin_q:%d/%m %H:%M}, "
              "pero el Observatorio no trae sus zonas: el mapa no lo pinta.")
        print("      Si quieres pintarlo, pega en el txt la lista 'Departamentos bajo aviso' del WhatsApp.")
    if capt.strftime("%Y-%m-%d") != dia:
        print(f"  [!] AVISO: el snapshot es del {capt:%d/%m} y el reporte es del {d['fecha']}.")
    print("-" * 70)
    return d


def _leer_txt_pacifico(ruta):
    """Del txt solo se usa lo de Pacífico. El mensaje del FEN de WhatsApp, si alguien lo pega, se ignora."""
    d = leer_datos_whatsapp(ruta)
    for k in ("lluvias", "vias", "rios", "huaicos", "p1", "p2", "medidas", "recomendaciones"):
        d[k] = []
    d["p1_titulo"] = d["p2_titulo"] = d["horario"] = d["vigencia"] = ""
    return d


def main():
    nombre_txt = os.path.join(CARPETA, "mensaje_observatorio.txt")
    if len(sys.argv) == 1 or sys.argv[1].startswith("-"):
        sys.argv.insert(1, nombre_txt)
    base.VERSION = f"{VERSION_OBS} · diseño {base.VERSION}"
    base.leer_datos = leer_datos_observatorio
    base.main()


if __name__ == "__main__":
    main()
