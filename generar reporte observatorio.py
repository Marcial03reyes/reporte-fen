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

VERSION_OBS = "2026-11-27 (Observatorio · lluvias con mm primero, vías cortadas por inundación)"
CARPETA = os.path.dirname(os.path.abspath(__file__))
LIMA = timezone(timedelta(hours=-5))

# Cuentas institucionales cuyas publicaciones en X se aceptan como hechos (el resto se ignora)
CUENTAS_OFICIALES = {
    "@senamhiperu": "SENAMHI", "@senamhidz05": "SENAMHI", "@coenperu": "COEN", "@indeciperu": "INDECI",
    "@mtc_gobperu": "MTC", "@proviasnac_peru": "Provías", "@anaperu": "ANA", "@bomberos_pe": "Bomberos",
    "@coercallao": "COER Callao", "@naval_peru": "Marina", "@enfenperu": "ENFEN",
    "@munilima": "Municipalidad de Lima", "@dhn_mgp": "DHN", "@dicapiperu": "DICAPI",
}
# Frases que indican que el propio texto no está confirmado
NO_CONFIRMADO = re.compile(r"pendiente|por confirmar|sin confirmar|por corroborar|no confirmad|se viene verificando|"
                           r"en evaluaci[oó]n preliminar|presunt|habr[ií]a\b", re.I)
# Publicaciones que no son hechos del día (campañas, notas de prensa, actividades)
NO_HECHO = re.compile(r"nota de prensa|prep[aá]rate|recomienda|recomendaciones|campa[ñn]a|simulacro|papa\b|le[oó]n xiv|"
                      r"visita|#cuidatupiel|conoce c[oó]mo|infograf|incendio forestal|incendios forestales", re.I)
PRONOSTICO = re.compile(r"^\s*se prev[eé]n?|persistir[aá]n|contin[uú]an? las lluvias|\(\d{1,2}:\d{2} a \d{1,2}:\d{2} ?h\)|"
                        r"toma tus precauciones|aviso (de corto plazo|n[°º.]|especial)|contin[uú]ar[aá]n|se esperan?\b|"
                        r"pr[oó]ximas \d* ?horas|#pron[oó]stico|#aviso\b|se estima|podr[ií]an? alcanzar", re.I)
OBSERVADO = re.compile(r"\b(registr[oóa]\w*|acumul[oó]|acumularon|alcanz[oó]|super[oó]|superando|ocurri[oó]|ocurrieron|activ[oó]|"
                       r"report[oó]|se produjo|provoc[oó]|desbord)", re.I)
TEMPERATURA = re.compile(r"°\s*c|temperatura|calor|radiaci[oó]n", re.I)

SECCION_DE = [  # prioridad: el primer patrón que calce decide la sección
    ("oleaje", re.compile(r"oleaje|marejada|braveza|cierre de puertos|puertos cerrados", re.I)),
    ("huaicos", re.compile(r"huaico|deslizamiento|derrumbe|movimiento en masa|ca[ií]da de rocas", re.I)),
    ("rios", re.compile(r"quebrada|\br[ií]os?\b|caudal|desborde|inundaci|crecida", re.I)),
    ("vias", re.compile(r"(\bv[ií]as?\b|carretera|tr[aá]nsito|puente|tramo).{0,80}(interrump|restring|bloque|cerrad|"
                        r"intransitable)|(interrump|restring|bloque|cerrad).{0,60}(\bv[ií]as?\b|carretera|tr[aá]nsito)",
                        re.I)),
    ("lluvias", re.compile(r"(?<!a consecuencia de las )(?<!a causa de las )(?<!tras las )(?<!tras )"
                           r"(lluvi|precipitaci|acumul|\bmm\b|tormenta|chubasco|granizo|nevada)", re.I)),
    # daños por lluvias sin quebrada/huaico/vía cortada (viviendas, infraestructura): "otras afectaciones"
    ("rios", re.compile(r"da[ñn]os|afect[oóa]|damnificad|colaps|viviendas", re.I)),
]
MAX_POR_SECCION = 3          # las más recientes; así el estatus entra en una diapositiva con letra legible
MINUSCULA = {"huaico", "huaicos", "lluvias", "lluvia", "clima", "tiempo", "deslizamiento", "inundacion"}
META = {"reporte", "radar", "ahora", "minam", "senamhi", "eldato", "clima", "pronostico", "pronóstico", "aviso", "perúresponde",
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


MENCIONES = {"@proviasnac_peru": "Provías Nacional", "@mtc_gobperu": "el MTC", "@coenperu": "el COEN",
              "@indeciperu": "INDECI", "@senamhiperu": "el SENAMHI", "@anaperu": "la ANA"}


def limpiar_post(texto):
    t = re.sub(r"https?://\S+", " ", texto)
    t = re.sub(r"^\W*ante el ni[ñn]o,?\s*el\s+#?per[uú]\s*responde,?\s*", "", t, flags=re.I)   # lema de campaña
    t = re.sub(r"[\U0001F000-\U0001FFFF☀-➿⬀-⯿️‍]", " ", t)
    t = re.sub(r"#(\w+)", _hashtag, t)
    t = re.sub(r"@\w+", lambda m: MENCIONES.get(m.group(0).lower(), ""), t)
    t = re.sub(r"\b(el|la)\s+(el|la)\s+", r"\2 ", t, flags=re.I)                 # "el el MTC" -> "el MTC"
    t = re.sub(r"\b(el|la),\s+(el|la)\b", r"\2", t, flags=re.I)                  # "el, el MTC" -> "el MTC"
    t = re.sub(r"\b(a trav[eé]s de|de la|del|de)\s*([,.;])", r"\2", t)          # "a través de , retiró" -> sin hueco
    t = re.sub(r"^\s*[¡!][^!]{0,140}!\s*", "", t)                               # titular "¡Habilitamos …!" fuera
    t = re.sub(r"\s*(per[uú] ?responde\s*)?ante el fen[oó]meno el ni[ñn]o\.?\s*$", "", t, flags=re.I)   # cola de hashtags
    t = re.sub(r"\s+", " ", t).strip(" |:-–")
    t = re.sub(r"\s+([,.;:])", r"\1", t)                                         # "Nacional , retiró" -> "Nacional, retiró"
    t = re.sub(r"^(\|\s*)+", "", t).strip()
    # "En entrevista con RPP Noticias, la subdirectora del SENAMHI, Rosario Julca, informó que diversas…" -> "Diversas…"
    t = re.sub(r"^en (?:entrevista|declaraciones) (?:con|a) [^,]+,\s*(?:[^,]{0,60},\s*){0,3}"
               r"(?:inform[oó]|señal[oó]|indic[oó]|precis[oó]|explic[oó]|advirti[oó]) que\s+", "", t, flags=re.I)
    # redacción del COEN, más corta
    t = re.sub(r"\b(?:en )?el distrito, provincia y departamento de ([A-ZÁÉÍÓÚÑ]\w+)", r"en la ciudad de \1", t)
    t = re.sub(r"\bdel distrito y provincia de ([A-ZÁÉÍÓÚÑ]\w+)", r"de \1", t)
    t = re.sub(r"\bdistrito y provincia de ([A-ZÁÉÍÓÚÑ]\w+)", r"\1", t)
    t = re.sub(r"\ba consecuencia de (?:las )?lluvias intensas", "por lluvias intensas", t, flags=re.I)
    t = re.sub(r"\bse produjo una inundaci[oó]n que ocasion[oó] da[ñn]os", "una inundación causó daños", t, flags=re.I)
    t = re.sub(r"\bque ocasion[oó] da[ñn]os", "con daños", t, flags=re.I)
    t = re.sub(r"\b(?:la )?infraestructura de transporte\b", "vías", t, flags=re.I)
    t = re.sub(r"\binfraestructura educativa\b", "colegios", t, flags=re.I)
    t = re.sub(r"\blocal(?:es)? p[uú]blicos?\b", "locales públicos", t, flags=re.I)
    return t


def recortar(texto, maximo=210):
    """Hasta la última oración completa (las publicaciones llegan cortadas) y sin pasar de 'maximo'."""
    t = texto.strip()
    # fin de oración: un punto que no cierra una abreviatura de una letra ("C. P.") y va seguido de espacio
    fines = [m.end() for m in re.finditer(r"(?<!\b[A-ZÁÉÍÓÚÑ])\.(?=\s|$)", t)]
    if not t.endswith((".", "!", ")")):
        ok = [f for f in fines if f > 40]
        t = t[:ok[-1]] if ok else t.rstrip(",;: ") + "."
    if len(t) > maximo:
        ok = [f for f in fines if 60 < f <= maximo]
        if ok:
            t = t[:ok[-1]]
        else:                                   # sin fin de oración a mano: se corta en la última coma
            coma = t.rfind(", ", 80, maximo)
            t = (t[:coma] + ".") if coma > 0 else t[:maximo].rsplit(" ", 1)[0].rstrip(",;:") + "…"
    return t[:1].upper() + t[1:]


def quitar_previsiones(texto):
    """Dentro de un hecho, quita las oraciones de pronóstico ('Se prevé que las lluvias continúen…')."""
    frases = re.split(r"(?<=[.!])\s+", texto)
    quedan = [f for f in frases if not re.match(r"^(se prev[eé]n?|se espera|se esperan|podr[ií]a|persistir[aá]n|"
                                                r"al[eé]jate|sigue las|mantente|toma tus|ev[ií]ta|identifica|prepara)", f.strip(), re.I)]
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
def zona_de(texto):
    """'distrito de Querecotillo, provincia de Sullana' -> 'Querecotillo (Sullana)'; 'ciudad de Piura' -> 'la ciudad de Piura'."""
    m = re.search(r"distrito (?:de )?([A-ZÁÉÍÓÚÑ][\wáéíóúñ]+(?: [A-ZÁÉÍÓÚÑ][\wáéíóúñ]+)?)"
                  r"(?:,? (?:en la )?provincia (?:de )?([A-ZÁÉÍÓÚÑ][\wáéíóúñ]+))?", texto)
    if m:
        return m.group(1) + (f" ({m.group(2)})" if m.group(2) and m.group(2) != m.group(1) else "")
    m = re.search(r"ciudad de ([A-ZÁÉÍÓÚÑ][\wáéíóúñ]+)", texto)
    if m:
        return f"la ciudad de {m.group(1)}"
    m = re.search(r"sectores? [^.;]*?,? de ([A-ZÁÉÍÓÚÑ][\wáéíóúñ]+)", texto)
    return m.group(1) if m else ""


def mismo_evento(a, b):
    """Dos publicaciones del mismo lugar que nombran los mismos sitios (p. ej. el COEN repite un reporte)."""
    nombres = lambda t: {w for w in re.findall(r"\b[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{3,}\b|\b\d+ de [A-Z]\w+", t)} - \
        {"Piura", "Lima", "Tumbes", "Durante", "Desde", "Tránsito", "Inundación", "Personal", "Esta", "Asimismo"}
    return a["lugar"] == b["lugar"] and len(nombres(a.get("cuerpo") or a["texto"]) &
                                            nombres(b.get("cuerpo") or b["texto"])) >= 2


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
        if re.search(r"papa\b|le[oó]n xiv|incendio", crudo, re.I) or (NO_HECHO.search(crudo) and not OBSERVADO.search(crudo)):
            descartes["no es un hecho"] += 1
            continue
        texto = limpiar_post(crudo)
        if PRONOSTICO.search(crudo) or PRONOSTICO.search(texto):
            # Un aviso de corto plazo puede traer también un dato ya ocurrido ("El Ciruelo registró 71 mm",
            # "el río Tumbes alcanzó 840 m³/s"): se queda solo con esas oraciones; si no hay ninguna, se descarta.
            obs = [f for f in re.split(r"(?<=[.!])\s+", texto) if OBSERVADO.search(f) and not PRONOSTICO.search(f)]
            if not obs:
                descartes["pronóstico"] += 1
                continue
            texto = " ".join(obs)
        seccion = next((k for k, rx in SECCION_DE if rx.search(texto)), None)
        if not seccion or (seccion == "lluvias" and TEMPERATURA.search(texto) and not re.search(r"\bmm\b|lluvios", texto)):
            descartes["no es un hecho"] += 1
            continue
        lugar, cuerpo = etiqueta(texto, x.get("departamentos") or [])
        # "hoy" de la publicación -> su fecha ("La tarde de hoy" del 17/11 -> "La tarde del 17/11")
        cuerpo = re.sub(r"\b(la (?:tarde|noche|mañana|madrugada)) de hoy\b", lambda m_: f"{m_.group(1)} del {t:%d/%m}",
                        cuerpo, flags=re.I)
        cuerpo = re.sub(r"^hoy,?\s*", f"El {t:%d/%m}, ", cuerpo, flags=re.I)
        cuerpo = re.sub(r"\bhoy\b", f"el {t:%d/%m}", cuerpo, flags=re.I)
        if lugar:                                       # "Piura: En Piura se registraron…" -> "Piura: Se registraron…"
            cuerpo = re.sub(r"^en " + re.escape(lugar) + r",?\s+", "", cuerpo, flags=re.I)
        cuerpo = cuerpo[:1].upper() + cuerpo[1:]
        cuerpo = quitar_previsiones(cuerpo)
        # Sin lugar, o un texto que continúa otra publicación ("Por otro lado…", "Asimismo…"): no se entiende solo
        if not lugar or re.match(r"^(por otro lado|asimismo|adem[aá]s|tambi[eé]n|en tanto|mientras)\b", cuerpo, re.I) \
                or not OBSERVADO.search(cuerpo) and seccion == "lluvias":
            descartes["no es un hecho"] += 1
            continue
        deps = {base._sin_tildes(x_) for x_ in (x.get("departamentos") or [])} or {base._sin_tildes(lugar)}
        if seccion == "rios":       # los acumulados de lluvia ya van en Lluvias: no se repiten aquí
            frs = [f for f in re.split(r"(?<=[.!])\s+", cuerpo)
                   if not (re.search(r"\bmm\b", f) or re.match(r"^lluvias? contin", f, re.I))]
            cuerpo = " ".join(frs) or cuerpo
        texto_final = recortar(cuerpo)
        if seccion == "vias":                           # la oración que dice que la vía está cortada no se pierde
            clave_v = next((f for f in re.split(r"(?<=[.!])\s+", cuerpo)
                            if re.search(r"interrump|restring|bloque|cerrad|intransitable", f, re.I)), "")
            if clave_v and clave_v.rstrip(".") not in texto_final:
                texto_final = recortar(cuerpo, 150).rstrip(".…") + ". " + recortar(clave_v, 120)
        secciones[seccion].append({"t": t, "cuenta": CUENTAS_OFICIALES[cuenta], "lugar": lugar, "deps": deps,
                                   "texto": texto_final, "cuerpo": cuerpo})
    hechos_prensa(snap, desde, hasta, secciones, descartes)
    # Una quebrada o huaico que corta una vía también va (corto) a Vías restringidas
    for k in ("rios", "huaicos"):
        for h in list(secciones[k]):
            m = re.search(r"(interrump|restring|bloque|cerrad)\w*\s+(?:el tr[aá]nsito en\s+)?(?:la |el )?"
                          r"((?:v[ií]a|carretera|tramo|puente|km)\s+[^,.;]+)", h["texto"], re.I)
            if m:
                estado = {"interrump": "interrumpido", "restring": "restringido", "bloque": "bloqueado",
                          "cerrad": "cerrado"}[m.group(1).lower()]
                lugar_v = m.group(2).strip()
                art = "la" if re.match(r"v[ií]a|carretera", lugar_v, re.I) else "el"
                causa = ("por quebradas" if k == "rios" else
                         "por deslizamiento" if re.search(r"desliz", h["texto"], re.I) else "por huaico")
                secciones["vias"].append({**h, "texto": f"Tránsito {estado} en {art} {lugar_v} ({causa})."})
                continue
            # "El tránsito vehicular se encuentra interrumpido" sin nombrar la vía: se usa la zona del hecho
            zona = zona_de(h.get("cuerpo") or h["texto"])
            tr = re.search(r"tr[aá]nsito(?: vehicular)? (?:se encuentra |est[aá] |qued[oó] )?"
                           r"(interrumpido|restringido|bloqueado|cerrado)", h.get("cuerpo") or "", re.I)
            if tr and zona:
                causa = ("por activación de quebrada" if re.search(r"quebrada", h["texto"], re.I) else
                         "por inundación" if re.search(r"inundaci", h["texto"], re.I) else
                         "por deslizamiento" if re.search(r"desliz", h["texto"], re.I) else "por huaico")
                secciones["vias"].append({**h, "texto": f"Tránsito {tr.group(1).lower()} en {zona}, {causa}."})
    # Lluvias: SENAMHI publica varios balances al día; queda el más reciente de cada departamento
    # (una publicación más antigua se descarta si todos sus departamentos ya salen en una más nueva).
    # Ríos, huaicos y vías: cada publicación es un hecho distinto; solo se quitan textos repetidos.
    for k, lista in secciones.items():
        # Lluvias: primero las que traen datos medidos (mm); dentro de eso, la más reciente
        lista.sort(key=lambda h: ((k == "lluvias" and bool(re.search(r"\d\s*mm\b", h["texto"]))), h["t"]),
                   reverse=True)
        cubiertos, textos, unicos = set(), set(), []
        for h in lista:
            if k == "lluvias" and h["deps"] <= cubiertos:
                continue
            if k != "lluvias" and any(mismo_evento(h, u) for u in unicos):
                continue
            clave_txt = re.sub(r"\W+", "", h["texto"].lower())[:120]
            if clave_txt in textos:
                continue
            cubiertos |= h["deps"]
            textos.add(clave_txt)
            unicos.append(h)
        secciones[k] = unicos[:MAX_POR_SECCION]
    return secciones, descartes


TIPO_PRENSA = {"oleaje": "oleaje", "marejada": "oleaje", "deslizamiento": "huaicos", "huaico": "huaicos",
               "derrumbe": "huaicos", "quebrada": "rios", "rio": "rios", "desborde": "rios", "inundacion": "rios",
               "lluvia": "lluvias", "via": "vias", "transito": "vias"}


def hechos_prensa(snap, desde, hasta, secciones, descartes):
    """Notas de prensa que el Observatorio ya revisó y marcó como HECHO (sources.pressEditorialReviews):
    cada una trae los hechos ya extraídos (tipo, departamento, localidad, cuándo ocurrió)."""
    for rv in (snap.get("sources") or {}).get("pressEditorialReviews") or []:
        if (rv.get("decision") or "").lower() != "hecho":
            continue
        for h in rv.get("hechos") or []:
            texto = (h.get("hecho") or "").strip()
            dep = (h.get("departamento") or "").strip()
            if not texto or not dep:
                continue
            oc = h.get("ocurrido") or ""
            t = a_lima(oc) if len(oc) > 10 else None
            if t is None and len(oc) == 10:                 # solo fecha: vale si ese día cae en la ventana
                t = a_lima(rv.get("publicado")) if desde.strftime("%Y-%m-%d") <= oc <= hasta.strftime("%Y-%m-%d") else None
                if t is None:
                    t = desde - timedelta(days=1)
            t = t or a_lima(rv.get("publicado"))            # sin fecha del hecho: la de la nota
            if not t or not (desde <= t <= hasta):
                descartes["fuera de ventana"] += 1
                continue
            tipo = base._sin_tildes((h.get("tipo") or "").lower())
            seccion = next((v for k, v in TIPO_PRENSA.items() if k in tipo), None) or \
                next((k for k, rx in SECCION_DE if rx.search(texto)), "rios")
            secciones.setdefault(seccion, []).append({
                "t": t, "cuenta": "Prensa verificada", "lugar": dep, "deps": {base._sin_tildes(dep)},
                "localidad": (h.get("localidad") or "").split(",")[0].strip(),
                "texto": recortar(texto[:1].upper() + texto[1:])})


SEDES_DPTO = ("Lima", "Piura", "Lambayeque", "La Libertad", "Junín", "Arequipa", "Cusco")


def calor_uv(snap, dia):
    """Sección 'Calor y radiación UV': un ítem por aviso de calor del SENAMHI vigente hoy, con las máximas
    previstas y dónde está en nivel extremo (rojo); si no llega a rojo, dónde está en alto (naranja)."""
    filas = [r for r in (snap.get("sources") or {}).get("senamhi", {}).get("rows", [])
             if r.get("jornada") == dia and not r.get("solo_contexto") and "calor" in (r.get("eventIds") or [])
             and r.get("nivel") in base.NIVEL_ORDEN]
    items = []
    for num in sorted({r.get("numero_aviso") for r in filas if r.get("numero_aviso")}, reverse=True):
        fa = [r for r in filas if r.get("numero_aviso") == num]
        ev = fa[0].get("evento") or ""
        zona = re.search(r"\bEN LA ([A-ZÁÉÍÓÚÑ ]+?)(?:\s*\(|\.|$)", ev.split(".")[0])
        zona = re.sub(r"\s+y en la\s+", " y ", zona.group(1).lower().strip()) if zona else ""
        temps = [int(t_) for t_ in re.findall(r"(\d{2})\s*°\s*C", ev)]
        nocturna = "NOCTURNA" in ev.upper()
        top = max((base.NIVEL_ORDEN[r["nivel"]] for r in fa), default=0)
        nv = next(k for k, v in base.NIVEL_ORDEN.items() if v == top)
        deps = {}
        for r in fa:
            if r["nivel"] == nv:
                for d_ in r.get("departamentos") or []:
                    deps.setdefault(d_, set()).update(p_.title() for p_ in r.get("provincias") or [])
        nombres = sorted(deps, key=base._sin_tildes)
        if len(nombres) <= 2:          # pocos departamentos: se nombran las provincias, como el aviso
            lugares = "; ".join(f"{d_} ({base._unir_y(sorted(deps[d_]))})" for d_ in nombres)
        elif len(nombres) > 6:         # lista larga: el total y los departamentos con sede de Pacífico
            con_sede = [x_ for x_ in SEDES_DPTO if x_ in nombres][:3]
            lugares = f"{len(nombres)} dptos." + (f" (incl. {base._unir_y(con_sede)})" if con_sede else "")
        else:
            lugares = base._unir_y(nombres)
        fin = max((a_lima(r.get("fin_aviso")) for r in fa if a_lima(r.get("fin_aviso"))), default=None)
        txt = (f"{zona.capitalize()}{' (noches)' if nocturna else ''}: " if zona else "") + \
              (f"{'mínimas' if nocturna else 'máximas'} de hasta {max(temps)} °C. " if temps else "") + \
              f"Nivel {base.INTENSIDAD[nv].lower()} ({nv.lower()}) en {lugares}" + \
              (f" (aviso SENAMHI {num}, hasta {fin:%d/%m})." if fin else f" (aviso SENAMHI {num}).")
        items.append((nocturna, txt[:1].upper() + txt[1:]))
    return [t_ for _, t_ in sorted(items, key=lambda x: x[0])]      # primero el calor de día, luego el nocturno


def pronostico(snap, dia, desde):
    """Costa / Sierra / Selva del último pronóstico de SENAMHI en X; si no hay, el aviso de lluvias del día.
    Más los otros avisos vigentes (calor, viento, oleaje)."""
    items = []
    posts = [x for x in (snap.get("datasets") or {}).get("senales_x") or []
             if (x.get("cuenta") or "").lower() == "@senamhiperu" and re.search(r"#pron[oó]stico", x.get("resumen") or "", re.I)
             and (a_lima(x.get("publicado")) or desde) >= desde + timedelta(hours=18)]   # solo el publicado el día del reporte
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
                por_region.setdefault(m.group(1).capitalize(), recortar(m.group(2).strip(), 200))
        else:                                   # texto libre: cada oración va a su región
            txt = re.sub(r"^.*?\|\s*", "", txt)
            for f in re.split(r"(?<=\.)\s+", txt):
                if len(f) > 20:
                    r_ = region_de(f)
                    por_region[r_] = (por_region.get(r_, "") + " " + recortar(limpia(f), 200)).strip()
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
                por_region[reg] = recortar(limpia(f), 200)
    for reg in ("Costa", "Sierra", "Selva"):
        if reg in por_region:
            items.append(f"{reg}: {por_region[reg]}")
    if filas:
        items.append(f"Acumulados previstos (aviso SENAMHI {n}): " + recortar(cuerpo, 170))
    # Departamentos por nivel oficial, por peligro (como "Pronóstico: nivel … hoy" de la página)
    filas_dia = [r for r in (snap.get("sources") or {}).get("senamhi", {}).get("rows", [])
                 if r.get("jornada") == dia and not r.get("solo_contexto") and r.get("nivel") in base.NIVEL_ORDEN]
    for ev, nombre in (("lluvias", "Lluvias"),):        # el calor va como sección propia del estatus
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
        extra = ""
        if ev == "calor":
            # Temperatura máxima prevista por aviso (como el mensaje "Calor y radiación UV" que llega por WhatsApp)
            maximas = []
            for a_ in sorted(a_ for a_ in avisos if a_):
                fila = next((r for r in filas_dia if r.get("numero_aviso") == a_), None)
                if not fila:
                    continue
                temps = [int(t_) for t_ in re.findall(r"(\d{2})\s*°\s*C", fila.get("evento") or "")]
                zona = re.search(r"\bEN LA ([A-ZÁÉÍÓÚÑ ]+?)(?:\s*\(|\.|$)", (fila.get("evento") or "").split(".")[0])
                if temps and zona:
                    z_ = re.sub(r"\s+y en la\s+", " y ", zona.group(1).lower().strip())
                    maximas.append(f"{max(temps)} °C en la {z_}")
            if maximas:
                extra = " Máximas previstas de hasta " + base._unir_y(maximas) + "."
        items.append(f"{nombre}: nivel " + "; ".join(partes) + f" (aviso{'s' if len(avisos) > 1 else ''} SENAMHI {num})." + extra)
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
    # el oleaje (DHN) va en su propia sección del estatus, no aquí
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
    # Departamentos con quebradas activadas o huaicos reportados (COEN/INDECI): ícono de quebrada en el mapa
    d["quebradas_hechos"] = sorted({dep for k in ("rios", "huaicos") for h in secciones[k]
                                    if re.search(r"quebrada|huaico", h["texto"], re.I)
                                    for dep in ([h["lugar"]] if h["lugar"] and " y " not in h["lugar"]
                                                else re.split(r",\s*|\s+y\s+", h["lugar"])) if dep})
    adic = " ".join(d.get("adicionales") or []).lower()
    for k, lista in secciones.items():
        # un hecho de prensa que ya está en [ALERTAS ADICIONALES] del txt (p. ej. Costa Verde) no se repite
        lista[:] = [h for h in lista if not (h.get("localidad") and h["localidad"].lower() in adic)]
        d[k] = [(f"{h['lugar']}: " if h["lugar"] else "") + h["texto"] for h in lista]
    d["origen_observatorio"] = True
    d["p1_titulo"] = "HOY"
    d["p1"] = pronostico(snap, dia, desde)
    d["calor"] = calor_uv(snap, dia)
    dhn = []
    for no in (snap.get("sources") or {}).get("dhn", {}).get("notices", []):
        ini, fin = a_lima(no.get("inicio")), a_lima(no.get("fin"))
        if ini and fin and ini.strftime("%Y-%m-%d") <= dia <= fin.strftime("%Y-%m-%d"):
            dhn.append((str(no.get("number")), ini, fin))
    dhn.sort(key=lambda x: x[0])
    if len(dhn) == 1:
        n_, ini, fin = dhn[0]
        d["oleaje"].append(f"Litoral: aviso especial de oleaje anómalo N.° {n_} de la DHN, vigente del {ini:%d/%m} al {fin:%d/%m}.")
    elif dhn:                          # varios avisos vigentes: una sola línea
        d["oleaje"].append(f"Litoral: avisos especiales de oleaje anómalo N.° {base._unir_y([x[0] for x in dhn])} de la DHN, "
                           f"vigentes hasta el {max(x[2] for x in dhn):%d/%m}.")
    d["nota"] = d.get("nota") or "Se mantiene el monitoreo."
    base.sintetizar(d)

    print("-" * 70)
    print(f"[OK] OBSERVATORIO: datos del {capt:%d/%m %H:%M} · hechos de {desde:%d/%m %H:%M} a {hasta:%d/%m %H:%M}")
    for k, nombre in (("lluvias", "Lluvias"), ("rios", "Ríos y quebradas"), ("huaicos", "Huaicos"), ("vias", "Vías"),
                      ("oleaje", "Oleaje (hechos)")):
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
    for k in ("lluvias", "vias", "rios", "huaicos", "calor", "oleaje", "p1", "p2", "medidas", "recomendaciones"):
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
