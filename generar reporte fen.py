"""
Genera el Reporte FEN de Pacífico Seguros en PowerPoint (.pptx) y en imagen (.png).

Uso:
    python generar_reporte_fen.py                      (usa mensaje_whatsapp.txt; sirve el botón ▶ de VS Code)
    python generar_reporte_fen.py datos_fen.txt
    python generar_reporte_fen.py datos_fen.txt --borrador   (genera aunque falten confirmaciones)
    python generar_reporte_fen.py datos_fen.txt --sin-correo (no pregunta por el correo)

Correo: si existe destinatarios.txt junto al script, al final se abre el reporte para
revisarlo y se pide el OK antes de enviarlo por Outlook (o abrir el borrador).

Imágenes de marca en "recursos": logo_pacifico.png, torito.png y logo_continuidad.png.
Fuente: Nunito (redondeada, afín al logo de Pacífico), en recursos/fuentes.
    En Windows el script la instala sola para tu usuario la primera vez.

Requisitos:
    pip install python-pptx pillow
    Para el PNG (uno de los dos):
      - Windows con PowerPoint instalado + pip install pywin32   (recomendado)
      - LibreOffice instalado (cualquier sistema)
"""
import argparse
import io
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

# ---------------------------------------------------------------- colores
TEAL = RGBColor(0x0B, 0x9D, 0xB0)
TEAL_BAND = RGBColor(0x15, 0x97, 0xA8)
PINK = RGBColor(0xE6, 0x00, 0x7E)
NAVY = RGBColor(0x1B, 0x2A, 0x4E)
GREY = RGBColor(0x6B, 0x76, 0x90)
LINE = RGBColor(0xE1, 0xE5, 0xEC)
PAGE = RGBColor(0xF7, 0xF8, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PILL = RGBColor(0xE6, 0xF4, 0xF6)
RED = RGBColor(0xC4, 0x12, 0x3F)

VERSION = "2026-11-13b (letra más grande y PNG 4K)"

FONT_XB = "Nunito ExtraBold"   # títulos con más peso (recursos/fuentes/Nunito-ExtraBold.ttf)
FONT = "Nunito"   # tipografía redondeada, afín al logo de Pacífico (incluida en recursos/fuentes)


# ================================================================ 1. Leer datos
SECCIONES = {
    "LLUVIAS REPORTADAS": "lluvias",
    "VIAS RESTRINGIDAS": "vias",
    "VÍAS RESTRINGIDAS": "vias",
    "RIOS Y OTRAS AFECTACIONES": "rios",
    "RÍOS Y OTRAS AFECTACIONES": "rios",
    "PRONOSTICO 1": "p1",
    "PRONÓSTICO 1": "p1",
    "PRONOSTICO 2": "p2",
    "PRONÓSTICO 2": "p2",
    "NOTA": "nota",
    "CONFIRMACIONES": "conf",
    "IMPACTO PACIFICO": "impacto",
    "IMPACTO PACÍFICO": "impacto",
    "PIE DE PAGINA": "pie",
    "PIE DE PÁGINA": "pie",
    "MEDIDAS": "medidas",
    "COLABORADORES AFECTADOS": "colab_casos",
    "COLABORADORES": "colab_casos",
    "AFECTACION DE COLABORADORES": "colab_casos",
    "AFECTACIÓN DE COLABORADORES": "colab_casos",
    "RECOMENDACIONES": "recomendaciones",
    "RECOMENDACIONES PREVENTIVAS": "recomendaciones",
    "AGENCIAS": "agencias",
    "AGENCIAS PACIFICO": "agencias",
    "AGENCIAS PACÍFICO": "agencias",
}


def _vacio():
    return {"numero": "", "fecha": "", "horario": "", "lluvias": [], "vias": [], "rios": [],
            "p1_titulo": "", "p1": [], "p2_titulo": "", "p2": [], "nota": "",
            "conf": [], "sedes": "", "colab": "", "pie": [], "medidas": [], "recomendaciones": [], "vigencia": "",
            "colab_casos": [], "colab_nivel": "",
            "agencias": [], "ag_fecha": "", "ag_corte": "", "ag_resumen": ""}


SIN_AFECTACION = re.compile(r"^(sin\s+(afectaci|impacto|novedad|da[ñn]o|incidencia)|operativ|normal)", re.I)


LEVE = re.compile(r"controlad|leve|menor|pequeñ|pequen|m[ií]nim|sin impacto en la (operaci|atenci)|"
                  r"operativ|atenci[oó]n normal", re.I)


def nivel(valor):
    """'ok' (sin afectación), 'leve' (afectación menor o controlada) o 'afectado'."""
    if not es_afectado(valor):
        return "ok"
    return "leve" if LEVE.search(valor) else "afectado"


def es_afectado(valor):
    """True si el texto describe una afectación (no empieza con 'Sin afectación', 'Sin impacto', etc.)."""
    return bool(valor.strip()) and not SIN_AFECTACION.match(valor.strip())


def _linea_colaborador(d, linea):
    """Una línea de la sección de colaboradores afectados. Formato recomendado:  Lugar: detalle
    (p. ej. 'Lima: 5 colaboradores con filtraciones en sus viviendas, en evaluación')."""
    linea = linea.strip().rstrip(",")
    if not linea or re.match(r"^(saludos|atentamente|atte|gracias|buen(os|as) (d[ií]as|tardes|noches)\W*$)", linea, re.I):
        return
    linea = re.sub(r"^buen(os|as) (d[ií]as|tardes|noches)\s*[,.:]\s*", "", linea, flags=re.I)
    if linea and linea[0].islower():
        linea = linea[0].upper() + linea[1:]
    if re.fullmatch(r"(s/?n|sin afectaci[oó]n.*|ninguno|ninguna|-+)\.?", linea, re.I):
        return
    # Término genérico: "colaboradora(s)" -> "colaborador(es)"
    linea = re.sub(r"\bcolaboradoras\b", "colaboradores", linea, flags=re.I)
    linea = re.sub(r"\bcolaboradora\b", "colaborador", linea, flags=re.I)
    d["colab_casos"].append(linea)


NUMEROS = {"un": 1, "una": 1, "uno": 1, "dos": 2, "tres": 3, "cuatro": 4, "cinco": 5, "seis": 6,
           "siete": 7, "ocho": 8, "nueve": 9, "diez": 10}


def resolver_colaboradores(d):
    """Si hay casos en [COLABORADORES AFECTADOS], arma el texto de 'Colaboradores' y su nivel."""
    if not d["colab_casos"]:
        return d
    resumen, total, todos_contados = [], 0, True
    niveles = []
    for caso in d["colab_casos"]:
        lugar, sep, detalle = caso.partition(":")
        if not sep or len(lugar) > 40:
            lugar, detalle = "", caso
        m = re.search(r"\b(\d+|un|una|uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez)\b\s+colaborador",
                      detalle, re.I)
        n = (int(m.group(1)) if m.group(1).isdigit() else NUMEROS[m.group(1).lower()]) if m else None
        if n is None:
            todos_contados = False
        else:
            total += n
        if lugar:
            resumen.append(f"{lugar.strip()} ({n})" if n else lugar.strip())
        niv = nivel(detalle)
        niveles.append("afectado" if niv == "ok" else niv)
    d["colab_nivel"] = "afectado" if "afectado" in niveles else "leve"
    lugares = ", ".join(resumen[:-1]) + " y " + resumen[-1] if len(resumen) > 1 else (resumen[0] if resumen else "")
    if todos_contados and total:
        d["colab"] = (f"{total} en total" if total > 1 else "1 colaborador") + (f": {lugares}." if lugares else ".")
    else:
        d["colab"] = (f"Casos en {lugares}." if lugares else f"{len(d['colab_casos'])} caso(s) reportado(s).")
    return d


def _linea_agencia(d, linea):
    """Procesa una línea del mensaje de agencias de Pacífico (ya sin emojis ni asteriscos)."""
    k, sep, valor = linea.partition(":")
    kl = k.strip().lower()
    if kl == "fecha":
        d["ag_fecha"] = valor.strip()
    elif kl in ("corte", "hora", "horario"):
        d["ag_corte"] = valor.strip()
    elif kl == "resumen":
        d["ag_resumen"] = valor.strip()
    elif re.match(r"^(saludos|atentamente|atte|gracias)", kl):
        pass
    elif "impacto" in kl and "agencia" in kl:
        pass                                    # es el título del mensaje
    elif sep and valor.strip():
        d["agencias"].append((k.strip(), valor.strip()))


def _es_seccion(linea):
    m = re.match(r"^\[(.+)\]$", linea.strip())
    return m.group(1).strip().upper() if m else None


def _leer_estructurado(lineas, d):
    """Formato con secciones entre corchetes: [LLUVIAS REPORTADAS], [CONFIRMACIONES], etc."""
    actual = None
    for raw in lineas:
        linea = raw.strip()
        if not linea or linea.startswith("#"):
            continue
        sec = _es_seccion(linea)
        if sec:
            actual = SECCIONES.get(sec)
            if actual is None:
                sys.exit(f"Sección desconocida en el archivo de datos: [{sec}]")
            continue
        linea = re.sub(r"^[•\-\*·–▪️\s]+", "", linea)
        clave, sep, valor = linea.partition(":")
        k = clave.strip().upper()
        if k in ("NUMERO", "NÚMERO") and actual in (None, "conf", "impacto"):
            d["numero"] = valor.strip()
        elif actual is None:
            if k == "FECHA":
                d["fecha"] = valor.strip()
            elif k == "HORARIO":
                d["horario"] = valor.strip()
        elif actual in ("p1", "p2") and k in ("TITULO", "TÍTULO"):
            d[actual + "_titulo"] = valor.strip()
        elif actual == "nota":
            d["nota"] = (d["nota"] + " " + linea).strip()
        elif actual == "conf":
            d["conf"].append(valor.strip() if sep else linea)
        elif actual == "impacto":
            if k == "SEDES":
                d["sedes"] = valor.strip()
            elif k == "COLABORADORES":
                d["colab"] = valor.strip()
        elif actual == "pie":
            d["pie"].append(linea)
        elif actual == "agencias":
            _linea_agencia(d, _limpiar(raw))
        elif actual == "colab_casos":
            _linea_colaborador(d, _limpiar(raw))
        else:
            d[actual].append(linea)
    return d


def _limpiar(linea):
    """Quita emojis, viñetas y caracteres invisibles del inicio de una línea de WhatsApp."""
    linea = linea.replace("​", "").replace("️", "").strip()
    linea = re.sub(r"^[^\wÁÉÍÓÚÑáéíóúñ¿¡(]+", "", linea)
    return linea.replace("*", "").replace("_", " ").strip()


def _etiqueta_bonita(linea):
    """'SIERRA: Se prevén...' -> 'Sierra: ...';  'LIMA METROPOLITANA: ...' -> 'Lima Metropolitana: ...'"""
    etq, sep, resto = linea.partition(":")
    if sep and etq.isupper() and len(etq) <= 40:
        menores = {"y", "e", "de", "del", "la", "las", "los", "el", "en", "a"}
        palabras = etq.strip().lower().split()
        etq = " ".join(w if (i and w in menores) else w[:1].upper() + w[1:] for i, w in enumerate(palabras))
        return etq + ": " + resto.strip()
    return linea


VINETAS = ("▪", "-", "•", "·", "–", "◦", "▫", "→", ">")


def _es_vineta(raw):
    """True si la línea original empieza con una viñeta (▪️, -, •), es decir, es un punto y no un título."""
    t = raw.replace("​", "").replace("️", "").lstrip()
    return t.startswith(VINETAS)


def _tipo_seccion(titulo):
    """Decide a qué parte del reporte va un título del mensaje, por palabras clave.
    Devuelve 'lluvias', 'rios', 'vias', 'pronostico', 'ignorar' o None si no lo reconoce."""
    t = titulo.lower()

    def hay(patron):                # compara palabras completas ("lluvias" no contiene la palabra "vías")
        return re.search(r"\b(" + patron + r")", t) is not None

    if hay(r"colaborador") and hay(r"afectaci|afectad") and not hay(r"credicorp|infraestructura"):
        return "colab_casos"        # casos de colaboradores afectados (alimenta "Colaboradores")
    if hay(r"agencias?") and hay(r"pac[ií]fico"):
        return "agencias"           # mensaje de agencias de Pacífico (alimenta "Sedes")
    if hay(r"pron[oó]stico"):
        return "pronostico"
    if hay(r"recomendaci[oó]n|recomendaciones|prevenci[oó]n|preventivas?"):
        return "recomendaciones"
    if hay(r"medidas|acciones|disposiciones"):
        return "medidas"
    if hay(r"impacto|credicorp|infraestructura|colaborador|agencia"):
        return "ignorar"            # el impacto lo completa Pacífico
    if hay(r"v[ií]as?\b|carreteras?\b|tr[aá]nsito|tramos?\b"):
        return "vias"
    if hay(r"r[ií]os?\b|quebradas?|crecida|huaicos?|deslizamientos?|desbordes?|caudal|otras afectaciones|inundaci"):
        return "rios"
    if hay(r"lluvias?|precipitaci|tormentas?"):
        return "lluvias"
    return None


def _leer_whatsapp(lineas, d):
    """Mensaje de WhatsApp pegado tal cual (con emojis, negritas *...* y viñetas ▪️ o -)."""
    actual = None
    for raw in lineas:
        if raw.strip().startswith("#"):
            continue
        linea = _limpiar(raw)
        if not linea:
            continue
        low = linea.lower()
        k, sep, valor = linea.partition(":")
        kl = k.strip().lower()
        if re.match(r"^(saludos|atentamente|atte|gracias)\b", low):
            continue

        # Dentro del mensaje de agencias de Pacífico, fecha/corte/resumen son de ese mensaje
        # Dos puntos "de etiqueta" (Tumbes: ...), sin contar los de una hora como 13:00
        sep_etq = re.search(r"(?<!\d):(?!\d)", linea.rstrip(":")) is not None

        titulo_posible = (not _es_vineta(raw)) and len(linea) <= 90 and (
            linea.endswith(":") or not sep_etq or re.match(r"pron[oó]stico\b", low) is not None)
        if actual == "agencias" and not (titulo_posible and _tipo_seccion(linea) not in (None, "agencias", "ignorar")):
            _linea_agencia(d, linea)
            continue
        if actual == "colab_casos" and not (titulo_posible and _tipo_seccion(linea) not in (None, "colab_casos")):
            _linea_colaborador(d, linea)
            continue

        # Datos de cabecera
        if kl in ("numero", "número", "n°", "nro"):
            d["numero"] = valor.strip()
            continue
        # La fecha y el corte del reporte son los del PRIMER mensaje (el Reporte FEN);
        # si después aparece otra fecha (p. ej. del mensaje de agencias), no la reemplaza.
        if kl == "fecha":
            if not d["fecha"]:
                d["fecha"] = valor.strip()
            continue
        if kl in ("corte", "horario", "hora"):
            if not d["horario"] or d["horario"].startswith("Vigencia"):
                d["horario"] = linea.strip() if kl == "corte" else valor.strip()
            continue
        if kl.startswith("vigencia"):
            # "Del 05/11/2026 a las 13:00 h al 06/11/2026 a las 13:00 h." -> "Vigencia: 05/11 13:00 h – 06/11 13:00 h"
            v = valor.strip().rstrip(".")
            m = re.findall(r"(\d{1,2}/\d{1,2})(?:/\d{2,4})?\D{0,12}?(\d{1,2}:\d{2})", v)
            corto = " – ".join(f"{f} {h} h" for f, h in m[:2]) if len(m) >= 2 else v
            d["vigencia"] = v
            if not d["horario"]:
                d["horario"] = "Vigencia: " + corto
            continue
        if "reporte fenomenol" in low:
            # A veces el título trae todo en una línea: "... NACIONAL 📅 Fecha: 10/11/2026 🕗 Corte: 07:00 horas"
            mf = re.search(r"fecha\s*:\s*(\d{1,2}/\d{1,2}/\d{2,4})", linea, re.I)
            mc = re.search(r"corte\s*:\s*(\d{1,2}:\d{2}(?:\s*(?:horas|hrs?|h)\b\.?)?)", linea, re.I)
            if mf and not d["fecha"]:
                d["fecha"] = mf.group(1)
            if mc and (not d["horario"] or d["horario"].startswith("Vigencia")):
                d["horario"] = "Corte: " + mc.group(1).strip()
            continue
        if low.startswith("se mantiene el monitoreo"):
            d["nota"] = linea
            actual = None
            continue

        # ¿Es un título de sección? (no lleva viñeta y no tiene texto después de los dos puntos)
        es_titulo = titulo_posible
        if es_titulo:
            tipo = _tipo_seccion(linea)
            if tipo == "pronostico":
                actual = "p1" if not d["p1"] else "p2"
                titulo = re.sub(r"^pron[oó]stico\s*[–\-:]*\s*", "", linea, flags=re.I).rstrip(":").strip()
                titulo = re.sub(r"^de\s+", "", titulo, flags=re.I)          # "Pronóstico de hoy" -> "HOY"
                titulo = re.sub(r"\s+", " ", re.sub(r"(\S)\(", r"\1 (", titulo))   # "HOY(Hasta" -> "HOY (Hasta"
                d[actual + "_titulo"] = titulo.upper() if titulo else "PRONÓSTICO"
            elif tipo:
                actual = tipo
            else:
                actual = "ignorar"
                print(f"  Aviso: no reconocí la sección «{linea.rstrip(':')}»; sus líneas no se incluyen.")
            continue

        if actual and actual != "ignorar":
            if re.fullmatch(r"(s/?n|n/?a|sin novedad(es)?|ninguno|ninguna|sin reportes?|-+)\.?", low.strip()):
                continue                                  # "s/n" = sin novedades: la sección queda en "Sin reportes."
            d[actual].append(_etiqueta_bonita(linea))
    return d


def leer_datos(ruta):
    with open(ruta, encoding="utf-8-sig") as f:
        lineas = f.read().splitlines()
    d = _vacio()
    secciones = [_es_seccion(l) for l in lineas]
    clima = {"LLUVIAS REPORTADAS", "VIAS RESTRINGIDAS", "VÍAS RESTRINGIDAS", "RIOS Y OTRAS AFECTACIONES",
             "RÍOS Y OTRAS AFECTACIONES", "PRONOSTICO 1", "PRONÓSTICO 1"}
    if any(s in clima for s in secciones if s):
        return resolver_colaboradores(resolver_sedes(_leer_estructurado(lineas, d)))
    # Mensaje de WhatsApp arriba + datos de Pacífico entre corchetes abajo
    corte = next((i for i, s in enumerate(secciones) if s), len(lineas))
    _leer_whatsapp(lineas[:corte], d)
    d = resolver_colaboradores(resolver_sedes(_leer_estructurado(lineas[corte:], d)))
    _avisar_colaboradores_sueltos(lineas, d)
    return d


def _avisar_colaboradores_sueltos(lineas, d):
    """Si el .txt menciona casos de colaboradores (p. ej. '5 colaboradores', 'una colaboradora') fuera de
    [COLABORADORES AFECTADOS], avisa: esas líneas no se usan y el reporte diría 'Sin afectación'."""
    if d["colab_casos"]:
        return
    patron = re.compile(r"\b(\d+|una?|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez)\s+colaborador", re.I)
    sueltas = [l.strip() for l in lineas if patron.search(l) and not l.strip().startswith("#")]
    if sueltas:
        print("  AVISO: hay casos de colaboradores fuera de la sección [COLABORADORES AFECTADOS]; no se incluyen.")
        print(f"         Ej.: «{sueltas[0][:80]}»")
        print("         Pásalos a esa sección, uno por línea (Lugar: detalle).")


def locales_sede(estado):
    """El mensaje de agencias a veces agrupa dos locales en una línea ('... ambos locales')."""
    return 1   # cada línea del mensaje de agencias = una sede (Piura tiene una sola, aunque diga "ambos locales")


def lista_sedes(d):
    """[(ciudad, n_locales, nivel)] en el orden del mensaje de agencias; Arequipa (Yanahuara) y
    Arequipa (Cayma) se juntan como 'Arequipa' con 2 locales."""
    orden_ = {"ok": 0, "leve": 1, "afectado": 2}
    res = {}
    for n_, e_ in (d["agencias"] or [(s_, "Sin afectación") for s_ in SEDES_POR_DEFECTO]):
        base = re.sub(r"\s*\(.*\)$", "", n_).strip()
        base = "Lima" if _sin_tildes(base) == "edificios lima" else base
        cant, nv = res.get(base, (0, "ok"))
        nv_e = nivel(e_)
        res[base] = (cant + locales_sede(e_), nv_e if orden_[nv_e] > orden_[nv] else nv)
    return [(k, c, n) for k, (c, n) in res.items()]


def resolver_sedes(d):
    """Si se pegó el mensaje de agencias de Pacífico, el estado de 'Sedes' sale de ahí.
    Lo que escribas a mano en 'Sedes:' manda solo si describe una afectación."""
    if not d["agencias"] or es_afectado(d["sedes"]):
        return d
    afectadas = [(n, e) for n, e in d["agencias"] if es_afectado(e)]
    total = sum(locales_sede(e) for _, e in d["agencias"])
    if not afectadas:
        d["sedes"] = f"Sin afectación en las {total} sedes revisadas." if total > 1 else \
            "Sin afectación en la sede revisada."
    elif len(afectadas) == 1:
        n, e = afectadas[0]
        d["sedes"] = f"{n} – {e[:1].lower() + e[1:]}"
    else:
        nombres = [n for n, _ in afectadas]
        d["sedes"] = ", ".join(nombres[:-1]) + " y " + nombres[-1] + f" ({len(afectadas)} de {total})."
    return d


# ================================================================ 2. Medir texto
def _cargar_fuente(bold=False):
    propia = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recursos", "fuentes",
                          "Nunito-Bold.ttf" if bold else "Nunito-Regular.ttf")
    if os.path.exists(propia):
        return ImageFont.truetype(propia, 100)
    nombres = (["arialbd.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"] if bold
               else ["arial.ttf", "Arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf"])
    carpetas = ["", "C:/Windows/Fonts", "/Library/Fonts", "/System/Library/Fonts/Supplemental",
                "/usr/share/fonts/truetype/liberation", "/usr/share/fonts/truetype/dejavu"]
    for n in nombres:
        for c in carpetas:
            try:
                return ImageFont.truetype(os.path.join(c, n) if c else n, 100)
            except OSError:
                continue
    return None


F_REG, F_BOLD = _cargar_fuente(), _cargar_fuente(True)


def ancho_pt(texto, size_pt, bold=False):
    f = F_BOLD if bold else F_REG
    if f is None:
        return len(texto) * size_pt * 0.52
    return f.getlength(texto) * size_pt / 100.0


def contar_lineas(etiqueta, texto, size_pt, ancho_disp_pt):
    """Cuenta líneas de un párrafo con etiqueta en negrita + texto normal (ajuste por palabras)."""
    palabras = [(w, True) for w in (etiqueta + ":").split()] if etiqueta else []
    palabras += [(w, False) for w in texto.split()]
    lineas, x = 1, 0.0
    esp = ancho_pt(" ", size_pt)
    for w, b in palabras:
        a = ancho_pt(w, size_pt, b) * 1.03
        if x and x + esp + a > ancho_disp_pt:
            lineas += 1
            x = a
        else:
            x += (esp if x else 0) + a
    return lineas


def separar(linea):
    m = re.match(r"^([^:]{1,60}):\s*(.*)$", linea)
    return (m.group(1).strip(), m.group(2).strip()) if m else ("", linea)


# ================================================================ 3. Íconos (dibujados con Pillow)
def _icono(dibujar, color, size=256):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    dr = ImageDraw.Draw(img)
    dibujar(dr, color, size)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def _lluvia(dr, c, s):
    w = s // 18
    dr.ellipse((s*.10, s*.30, s*.48, s*.62), outline=c, width=w)
    dr.ellipse((s*.30, s*.14, s*.74, s*.56), outline=c, width=w)
    dr.ellipse((s*.56, s*.30, s*.90, s*.62), outline=c, width=w)
    dr.rectangle((s*.16, s*.47, s*.84, s*.60), fill=(0, 0, 0, 0))
    dr.line((s*.24, s*.62, s*.80, s*.62), fill=c, width=w)
    for x in (.30, .50, .70):
        dr.line((s*x, s*.72, s*(x-.06), s*.88), fill=c, width=w)


def _barrera(dr, c, s):
    w = s // 18
    dr.rectangle((s*.08, s*.30, s*.92, s*.50), outline=c, width=w)
    for x in (.18, .40, .62):
        dr.line((s*x, s*.31, s*(x+.14), s*.49), fill=c, width=w)
    for x in (.20, .80):
        dr.line((s*x, s*.50, s*x, s*.86), fill=c, width=w)
        dr.line((s*x, s*.14, s*x, s*.30), fill=c, width=w)
    dr.line((s*.10, s*.86, s*.30, s*.86), fill=c, width=w)
    dr.line((s*.70, s*.86, s*.90, s*.86), fill=c, width=w)


def _rio(dr, c, s):
    w = s // 18
    for y in (.52, .74):
        for i in range(3):
            x0 = s*(.08 + i*.28)
            dr.arc((x0, s*y - s*.08, x0 + s*.28, s*y + s*.08), 180, 360, fill=c, width=w)
    dr.arc((s*.30, s*.10, s*.70, s*.42), 200, 340, fill=c, width=w)
    dr.ellipse((s*.46, s*.20, s*.56, s*.30), fill=c)


def _edificio(dr, c, s):
    w = s // 18
    dr.polygon([(s*.10, s*.36), (s*.50, s*.12), (s*.90, s*.36)], outline=c, width=w)
    for x in (.20, .40, .60, .80):
        dr.line((s*x, s*.42, s*x, s*.78), fill=c, width=w)
    dr.line((s*.08, s*.86, s*.92, s*.86), fill=c, width=w)
    dr.line((s*.12, s*.40, s*.88, s*.40), fill=c, width=w)


def _personas(dr, c, s):
    w = s // 18
    dr.ellipse((s*.18, s*.16, s*.42, s*.40), outline=c, width=w)
    dr.ellipse((s*.56, s*.16, s*.80, s*.40), outline=c, width=w)
    dr.arc((s*.06, s*.50, s*.54, s*1.10), 180, 360, fill=c, width=w)
    dr.arc((s*.46, s*.50, s*.94, s*1.10), 180, 360, fill=c, width=w)


def _radar(dr, c, s):
    w = s // 20
    dr.ellipse((s*.08, s*.08, s*.92, s*.92), outline=c, width=w)
    dr.ellipse((s*.28, s*.28, s*.72, s*.72), outline=c, width=w)
    dr.line((s*.50, s*.50, s*.80, s*.20), fill=c, width=w)
    dr.ellipse((s*.45, s*.45, s*.55, s*.55), fill=c)


def _alerta(dr, c, s):
    """Triángulo de alerta con signo de exclamación."""
    w = s // 16
    dr.polygon([(s*.50, s*.10), (s*.92, s*.86), (s*.08, s*.86)], outline=c, width=w)
    dr.line((s*.50, s*.36, s*.50, s*.62), fill=c, width=w + 2)
    dr.ellipse((s*.45, s*.70, s*.55, s*.80), fill=c)


def _megafono(dr, c, s):
    w = s // 18
    dr.polygon([(s*.14, s*.40), (s*.34, s*.40), (s*.74, s*.18), (s*.74, s*.82), (s*.34, s*.60), (s*.14, s*.60)],
               outline=c, width=w)
    dr.line((s*.34, s*.40, s*.34, s*.60), fill=c, width=w)
    dr.line((s*.26, s*.60, s*.34, s*.86), fill=c, width=w)
    dr.arc((s*.70, s*.36, s*.94, s*.64), 300, 60, fill=c, width=w)


C_TEAL, C_PINK = (11, 157, 176, 255), (230, 0, 126, 255)


# ================================================================ 4. Helpers de PowerPoint
def caja(slide, x, y, w, h, color=None, linea=None, forma=MSO_SHAPE.RECTANGLE, nombre=None):
    s = slide.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
    if color is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = color
    if linea is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = linea
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    sin_estilo(s)
    if nombre:
        s.name = nombre
    return s


def sin_estilo(shape):
    """Quita el estilo del tema (evita sombras automáticas distintas en PowerPoint y LibreOffice)."""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def sombra_suave(shape):
    spPr = shape._element.spPr
    ef = spPr.find(qn("a:effectLst"))
    if ef is None:
        ef = etree.SubElement(spPr, qn("a:effectLst"))
    for ch in list(ef):
        ef.remove(ch)
    sh = etree.SubElement(ef, qn("a:outerShdw"), blurRad="76200", dist="19050", dir="5400000",
                          algn="t", rotWithShape="0")
    clr = etree.SubElement(sh, qn("a:srgbClr"), val="1E2A44")
    etree.SubElement(clr, qn("a:alpha"), val="14000")


def degradado(shape):
    shape.fill.gradient()
    shape.fill.gradient_angle = 0
    st = shape.fill.gradient_stops
    st[0].position, st[0].color.rgb = 0.0, TEAL_BAND
    st[1].position, st[1].color.rgb = 1.0, PINK


def texto(slide, x, y, w, h, partes, size, color=NAVY, bold=False, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP, nombre=None):
    """partes: str o lista de párrafos; cada párrafo es str o lista de (texto, bold, color)."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if nombre:
        tb.name = nombre
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    parrafos = [partes] if isinstance(partes, str) else partes
    for i, par in enumerate(parrafos):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = Pt(round(size * 1.2, 1))   # interlineado fijo (Nunito trae uno muy amplio)
        runs = [(par, bold, color)] if isinstance(par, str) else par
        for t, b, c in runs:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(size)
            r.font.bold = b is True
            r.font.color.rgb = c
            r.font.name = FONT_XB if b == "xb" else FONT
    return tb


def vinetas(slide, x, y, w, h, items, size, nombre, esp=3):
    """Lista con viñetas reales de PowerPoint; la etiqueta antes de ':' va en negrita."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = nombre
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, linea in enumerate(items):
        etq, txt = separar(linea)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.JUSTIFY
        p.space_after = Pt(esp)
        p.line_spacing = Pt(round(size * 1.15, 1))  # interlineado fijo: igual en PowerPoint y LibreOffice
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(0.13)))
        pPr.set("indent", str(-Inches(0.13)))
        bf = etree.SubElement(pPr, qn("a:buFont"))
        bf.set("typeface", FONT)
        bc = etree.SubElement(pPr, qn("a:buChar"))
        bc.set("char", "•")
        if etq:
            r = p.add_run()
            r.text = etq + ": "
            r.font.bold = True
            r.font.size, r.font.color.rgb, r.font.name = Pt(size), NAVY, FONT
        r = p.add_run()
        r.text = txt
        r.font.size, r.font.color.rgb, r.font.name = Pt(size), NAVY, FONT
    return tb


def alto_lista(items, size, ancho_in, esp=3):
    disp = ancho_in * 72 - 0.13 * 72
    total = 0.0
    for linea in items:
        etq, txt = separar(linea)
        n = contar_lineas(etq, txt, size, disp)
        total += n * size * 1.15 + esp
    return total / 72.0


def icono(slide, buf, x, y, lado):
    slide.shapes.add_picture(buf, Inches(x), Inches(y), Inches(lado), Inches(lado))


def linea_h(slide, x1, x2, y):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y), Inches(x2), Inches(y))
    c.line.color.rgb = RGBColor(0xDD, 0xE6, 0xF0)
    c.line.width = Pt(0.75)
    c.shadow.inherit = False
    sin_estilo(c)


# ================================================================ 4b. Radar (Observatorio Pacífico)
RADAR_URL = "https://observatorio-pacifico.pacificoabba0.chatgpt.site/api/radar/snapshot"
NIVEL_ORDEN = {"Rojo": 3, "Naranja": 2, "Amarillo": 1}
NIVEL_RGB = {"Rojo": RGBColor(0xC4, 0x12, 0x3F), "Naranja": RGBColor(0xE0, 0x6C, 0x00),
             "Amarillo": RGBColor(0xC9, 0xA2, 0x00), "": RGBColor(0x1E, 0x7A, 0x45)}
NIVEL_FONDO = {"Rojo": RGBColor(0xFD, 0xE7, 0xEE), "Naranja": RGBColor(0xFF, 0xEC, 0xDA),
               "Amarillo": RGBColor(0xFF, 0xF7, 0xD1), "": RGBColor(0xE4, 0xF4, 0xEA)}
# Ciudad de la sede -> departamento (para cruzar sedes con avisos)
CIUDAD_DEPTO = {
    "lima": "Lima", "edificios lima": "Lima", "callao": "Callao", "piura": "Piura", "sullana": "Piura",
    "chiclayo": "Lambayeque", "trujillo": "La Libertad", "huancayo": "Junín", "arequipa": "Arequipa",
    "cusco": "Cusco", "ica": "Ica", "tumbes": "Tumbes", "cajamarca": "Cajamarca", "iquitos": "Loreto",
    "tarapoto": "San Martín", "huanuco": "Huánuco", "tacna": "Tacna", "puno": "Puno", "juliaca": "Puno",
    "chimbote": "Áncash", "huaraz": "Áncash", "ayacucho": "Ayacucho", "pucallpa": "Ucayali",
    "moquegua": "Moquegua", "ilo": "Moquegua", "abancay": "Apurímac", "huancavelica": "Huancavelica",
}
# Ciudad de la sede -> provincia (los avisos SENAMHI vienen por provincia)
PROV_SEDE = {"edificios lima": "Lima", "lima": "Lima", "sullana": "Sullana", "chiclayo": "Chiclayo",
             "trujillo": "Trujillo", "huancayo": "Huancayo", "arequipa": "Arequipa", "cusco": "Cusco",
             "piura": "Piura", "juliaca": "San Roman", "chimbote": "Santa", "tarapoto": "San Martin",
             "iquitos": "Maynas", "pucallpa": "Coronel Portillo", "ayacucho": "Huamanga"}
SEDES_POR_DEFECTO = ["Piura", "Chiclayo", "Trujillo", "Huancayo", "Arequipa (Yanahuara)", "Arequipa (Cayma)",
                     "Cusco", "Edificios Lima"]


def _sin_tildes(t):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", t or "") if unicodedata.category(c) != "Mn").lower()


def cargar_radar(carpetas):
    """Intenta descargar el snapshot del Radar; si no se puede (p. ej. Cloudflare), usa un
    snapshot.json guardado a mano en alguna de las carpetas. Devuelve (datos, origen) o (None, motivo)."""
    import json
    import urllib.request
    try:
        req = urllib.request.Request(RADAR_URL, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/131.0 Safari/537.36 Edg/131.0",
            "Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=20) as r:
            datos = json.loads(r.read().decode("utf-8"))
        for c in carpetas[:1]:                              # copia local por si mañana falla la descarga
            try:
                with open(os.path.join(c, "snapshot.json"), "w", encoding="utf-8") as f:
                    json.dump(datos, f, ensure_ascii=False)
            except OSError:
                pass
        return datos, "descargado del Observatorio"
    except Exception as e:  # noqa: BLE001
        motivo = f"no se pudo descargar ({e.__class__.__name__})"
    for c in carpetas:
        ruta = os.path.join(c, "snapshot.json")
        if os.path.exists(ruta):
            try:
                with open(ruta, encoding="utf-8-sig") as f:
                    return json.load(f), f"archivo snapshot.json ({datetime.fromtimestamp(os.path.getmtime(ruta)):%d/%m %H:%M})"
            except Exception:  # noqa: BLE001
                continue
    return None, motivo + " y no hay snapshot.json en la carpeta"


def fin_local(iso):
    """Fecha (AAAA-MM-DD, hora de Lima) en que termina un aviso; '' si no se sabe."""
    from datetime import timedelta, timezone
    try:
        t = datetime.fromisoformat((iso or "").replace("Z", "+00:00"))
    except ValueError:
        return ""
    if t.tzinfo:
        t = t.astimezone(timezone(timedelta(hours=-5)))
    return t.strftime("%Y-%m-%d")


def resumir_radar(datos, fecha_reporte, sedes=None):
    """Extrae lo que NO viene por WhatsApp: nivel oficial de avisos de lluvia y de activación de
    quebradas por departamento (para la fecha del reporte), estado ENFEN y sedes en zona con aviso."""
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", fecha_reporte or "")
    dia = f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}" if m else datetime.now().strftime("%Y-%m-%d")
    src, ds = datos.get("sources", {}), datos.get("datasets", {})
    res = {"dia": dia, "lluvias": {}, "lluvias_avisos": [], "quebradas": {}, "quebradas_aviso": "",
           "enfen": None, "sedes": []}

    # Lluvias: avisos SENAMHI con evento "lluvias", jornada = fecha del reporte
    filas = [r for r in src.get("senamhi", {}).get("rows", [])
             if "lluvias" in (r.get("eventIds") or []) and not r.get("solo_contexto")]
    jornadas = sorted({r.get("jornada") for r in filas if r.get("jornada") and r.get("jornada") <= dia})
    jornada = dia if dia in jornadas else None      # solo avisos cuya jornada es el día del reporte
    if jornada:
        for r in filas:
            if r.get("jornada") != jornada:
                continue
            nv = r.get("nivel")
            if nv not in NIVEL_ORDEN:
                continue
            res["lluvias_avisos"].append(r.get("numero_aviso"))
            for dep in r.get("departamentos") or []:
                if NIVEL_ORDEN[nv] > NIVEL_ORDEN.get(res["lluvias"].get(dep), 0):
                    res["lluvias"][dep] = nv
    res["lluvias_avisos"] = sorted({a for a in res["lluvias_avisos"] if a})

    # Activación de quebradas (aviso SENAMHI integrado por INDECI)
    q = src.get("indeci", {})
    aviso_q = q.get("quebradaNotice") or {}
    areas = (q.get("quebradaAreas") or {}).get("areas") or []
    dias_q = sorted({a.get("day") for a in areas if a.get("day") and a.get("day") <= dia})
    dia_q = dia if dia in dias_q else (dias_q[-1] if dias_q else None)
    # Si el mapa de quebradas es de un día anterior, solo vale si el aviso sigue vigente el día del reporte
    if dia_q and dia_q != dia and fin_local(aviso_q.get("fin")) < dia:
        dia_q = None
    if dia_q and (datetime.strptime(dia, "%Y-%m-%d") - datetime.strptime(dia_q, "%Y-%m-%d")).days <= 1:
        for a in areas:
            if a.get("day") != dia_q:
                continue
            nv = (a.get("level") or "").capitalize()
            dep = a.get("department")
            if nv in NIVEL_ORDEN and dep and NIVEL_ORDEN[nv] > NIVEL_ORDEN.get(res["quebradas"].get(dep), 0):
                res["quebradas"][dep] = nv
        if aviso_q.get("number"):
            res["quebradas_aviso"] = str(aviso_q["number"])

    # ENFEN
    est = (ds.get("estado") or [None])[0]
    if est:
        prox = est.get("proximo") or ""
        mp = re.match(r"(\d{4})-(\d{2})-(\d{2})", prox)
        res["enfen"] = {"estado": est.get("estado", ""), "comunicado": est.get("comunicado", ""),
                        "proximo": f"{mp.group(3)}/{mp.group(2)}" if mp else prox}

    # Nivel por provincia (para las sedes): lluvias y quebradas
    prov_ll, prov_qb = {}, {}
    if jornada:
        for r in filas:
            if r.get("jornada") != jornada or r.get("nivel") not in NIVEL_ORDEN:
                continue
            for dep in r.get("departamentos") or []:
                for prov in r.get("provincias") or []:
                    clave = (_sin_tildes(dep), _sin_tildes(prov))
                    if NIVEL_ORDEN[r["nivel"]] > NIVEL_ORDEN.get(prov_ll.get(clave), 0):
                        prov_ll[clave] = r["nivel"]
    if dia_q:
        for a in areas:
            nv = (a.get("level") or "").capitalize()
            if a.get("day") == dia_q and nv in NIVEL_ORDEN:
                clave = (_sin_tildes(a.get("department")), _sin_tildes(a.get("province")))
                if NIVEL_ORDEN[nv] > NIVEL_ORDEN.get(prov_qb.get(clave), 0):
                    prov_qb[clave] = nv

    # Sedes Pacífico en zona con aviso: nivel de la PROVINCIA donde está la sede
    lista = sedes or SEDES_POR_DEFECTO
    vistos = {}
    for s in lista:
        base = re.sub(r"\s*\(.*\)$", "", s).strip()
        nombre = "Lima" if _sin_tildes(base) == "edificios lima" else base
        vistos[nombre] = vistos.get(nombre, 0) + 1
    for nombre, n in vistos.items():
        dep = CIUDAD_DEPTO.get(_sin_tildes(nombre), nombre)
        prov = PROV_SEDE.get(_sin_tildes(nombre), nombre)
        clave = (_sin_tildes(dep), _sin_tildes(prov))
        nl, nq = prov_ll.get(clave, ""), prov_qb.get(clave, "")
        nivel_max = max((nl, nq), key=lambda x: NIVEL_ORDEN.get(x, 0))
        motivos = [t for t, nv in (("lluvias", nl), ("quebradas", nq)) if nv]
        res["sedes"].append({"nombre": nombre + (f" ({n})" if n > 1 else ""), "nivel": nivel_max,
                             "motivo": " + ".join(motivos) if motivos else "sin aviso"})
    return res


def _lista_corta(deps, maximo=4):
    deps = sorted(deps)
    return ", ".join(deps) if len(deps) <= maximo else ", ".join(deps[:maximo]) + f" y {len(deps) - maximo} más"


def lineas_radar_texto(rz):
    """Propuesta A en texto plano (sirve para medir alto)."""
    out = []
    if rz["enfen"]:
        out.append(f"ENFEN: {rz['enfen']['estado']} · próximo comunicado {rz['enfen']['proximo']}")
    for clave, titulo in (("lluvias", "Lluvias"), ("quebradas", "Quebradas")):
        niv = rz[clave]
        if niv:
            partes = [f"● {_lista_corta([d for d, v in niv.items() if v == n])}"
                      for n in ("Rojo", "Naranja", "Amarillo") if any(v == n for v in niv.values())]
            out.append(f"{titulo}: " + "   ".join(partes))
    expuestas = [s for s in rz["sedes"] if s["nivel"]]
    if expuestas:
        out.append("Sedes Pacífico en zona con aviso: " + "   ".join(f"● {s['nombre']}" for s in expuestas))
    return out


def alto_radar(rz, estilo, ancho_in, fs=8.5):
    """Alto (pulgadas) que ocupa el bloque del Radar en la columna derecha."""
    if not rz:
        return 0.0
    if estilo == "mapa":
        return 2.45
    if estilo == "sedes":
        filas = (len(rz["sedes"]) + 3) // 4
        return 0.30 + filas * 0.40 + 0.22 + 0.20 + 0.10
    lineas = sum(contar_lineas("", t, fs, (ancho_in - 0.3) * 72) for t in lineas_radar_texto(rz))
    return 0.30 + lineas * fs * 1.2 / 72 + 0.06 * len(lineas_radar_texto(rz)) + 0.20 + 0.08


def dibujar_radar_mapa(slide, rz, x, y, w, h, impacto_sedes):
    """Propuesta 'mapa': silueta del Perú (color = alerta oficial) + sedes (punto = estado de la sede)."""
    caja_r = caja(slide, x, y, w, h, P_TINT, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Radar mapa")
    caja_r.adjustments[0] = 0.06
    texto(slide, x + 0.15, y + 0.07, w - 0.3, 0.24, "ALERTAS OFICIALES Y SEDES PACÍFICO", 10.5, P_BLUE_DK, True)
    niveles = dict(rz["lluvias"])
    for dep, nv in rz["quebradas"].items():
        if NIVEL_ORDEN[nv] > NIVEL_ORDEN.get(niveles.get(dep), 0):
            niveles[dep] = nv
    pines = [(nombre, nivel_imp) for nombre, nivel_imp in impacto_sedes]
    res = dibujar_mapa_png(niveles, pines)
    mx = x + 0.12
    mh = h - 0.62
    if res:
        buf, ratio = res
        mw = mh * ratio
        slide.shapes.add_picture(buf, Inches(mx), Inches(y + 0.34), Inches(mw), Inches(mh))
    else:
        mw = 0
    lx = mx + mw + 0.12
    lw = x + w - 0.12 - lx

    def sw(color_rgb):
        return ("■ ", True, RGBColor(*color_rgb))
    p = []
    p.append([("Alerta oficial (lluvias o quebradas)", True, P_INK)])
    p.append([sw(FILL_NIVEL["Rojo"]), ("Rojo  ", False, P_GREY), sw(FILL_NIVEL["Naranja"]), ("Naranja", False, P_GREY)])
    p.append([sw(FILL_NIVEL["Amarillo"]), ("Amarillo  ", False, P_GREY), sw(FILL_NIVEL[""]), ("Sin aviso", False, P_GREY)])
    p.append([("Sedes Pacífico", True, P_INK)])
    p.append([("● ", True, RGBColor(*PIN_NIVEL["ok"])), ("Sin afectación  ", False, P_GREY),
              ("● ", True, RGBColor(*PIN_NIVEL["leve"])), ("Leve", False, P_GREY)])
    p.append([("● ", True, RGBColor(*PIN_NIVEL["afectado"])), ("Con afectación", False, P_GREY)])
    expuestas = [s for s in rz["sedes"] if s["nivel"]]
    if expuestas:
        runs = [("Sedes en provincia con aviso hoy: ", True, P_INK)]
        for nv in ("Rojo", "Naranja", "Amarillo"):
            nombres = [s["nombre"] for s in expuestas if s["nivel"] == nv]
            if nombres:
                runs += [("■\u00a0", True, RGBColor(*FILL_NIVEL[nv])), (", ".join(nombres) + "  ", False, P_INK)]
    else:
        runs = [("Sedes en provincia con aviso hoy: ", True, P_INK), ("ninguna", False, P_INK)]
    p.append(runs)
    if rz["enfen"]:
        p.append([("ENFEN: ", True, P_INK), (f"{rz['enfen']['estado']} · próx. comunicado {rz['enfen']['proximo']}",
                                            False, P_INK)])
    tb = texto(slide, lx, y + 0.38, lw, h - 0.62, p, 8, nombre="Radar leyenda")
    for i, par in enumerate(tb.text_frame.paragraphs):
        par.space_after = Pt(4 if i in (2, 5, 6) else 1)
    texto(slide, x + 0.15, y + h - 0.21, w - 0.3, 0.16,
          "Fuente: Observatorio Pacífico / Torre de Control. Color del departamento = nivel más alto en "
          "alguna de sus provincias.", 6.5, P_GREY, False, nombre="Radar fuente")


def dibujar_radar(slide, rz, estilo, x, y, w, h, fs=8.5):
    """Dibuja el bloque del Radar (propuesta A: 'avisos'; propuesta B: 'sedes')."""
    caja_r = caja(slide, x, y, w, h, P_TINT, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Radar")
    caja_r.adjustments[0] = 0.08
    titulo = "AVISOS OFICIALES VIGENTES" if estilo != "sedes" else "SEDES PACÍFICO FRENTE A AVISOS OFICIALES"
    texto(slide, x + 0.15, y + 0.07, w - 0.3, 0.24, titulo, 10.5, P_BLUE_DK, True, nombre="Radar titulo")
    cy = y + 0.33
    if estilo != "sedes":
        parrafos = []
        if rz["enfen"]:
            parrafos.append([("ENFEN: ", True, P_INK),
                             (f"{rz['enfen']['estado']} · próximo comunicado {rz['enfen']['proximo']}", False, P_INK)])
        for clave, etq in (("lluvias", "Lluvias"), ("quebradas", "Quebradas")):
            niv = rz[clave]
            if not niv:
                continue
            aviso = (f" (aviso {', '.join(str(a) for a in rz['lluvias_avisos'])})" if clave == "lluvias"
                     and rz["lluvias_avisos"] else (f" (aviso {rz['quebradas_aviso']})" if clave == "quebradas"
                                                    and rz["quebradas_aviso"] else ""))
            runs = [(f"{etq}{aviso}: ", True, P_INK)]
            for n in ("Rojo", "Naranja", "Amarillo"):
                deps = [d for d, v in niv.items() if v == n]
                if deps:
                    runs += [("●\u00a0", True, NIVEL_RGB[n]), (_lista_corta(deps).replace(" ", "\u00a0") + "   ", False, P_INK)]
            parrafos.append(runs)
        expuestas = [s for s in rz["sedes"] if s["nivel"]]
        if expuestas:
            runs = [("Sedes Pacífico en zona con aviso: ", True, P_INK)]
            for s in sorted(expuestas, key=lambda s: -NIVEL_ORDEN[s["nivel"]]):
                runs += [("●\u00a0", True, NIVEL_RGB[s["nivel"]]), (s["nombre"].replace(" ", "\u00a0") + "   ", False, P_INK)]
            parrafos.append(runs)
        tb = texto(slide, x + 0.15, cy, w - 0.3, h - 0.55, parrafos, fs, nombre="Radar detalle")
        for p in tb.text_frame.paragraphs:
            p.space_after = Pt(3)
    else:
        sedes = rz["sedes"]
        por_fila = 4
        gap = 0.07
        cw = (w - 0.3 - gap * (por_fila - 1)) / por_fila
        for i, s in enumerate(sedes):
            fx, fy = x + 0.15 + (i % por_fila) * (cw + gap), cy + (i // por_fila) * 0.40
            ch = caja(slide, fx, fy, cw, 0.34, NIVEL_FONDO[s["nivel"]], None, MSO_SHAPE.ROUNDED_RECTANGLE,
                      "Sede " + s["nombre"])
            ch.adjustments[0] = 0.3
            texto(slide, fx + 0.06, fy + 0.01, cw - 0.1, 0.32,
                  [[("● ", True, NIVEL_RGB[s["nivel"]]), (s["nombre"], True, P_INK)],
                   [(s["motivo"] if s["nivel"] else "sin aviso", False, NIVEL_RGB[s["nivel"]])]],
                  7.5, anchor=MSO_ANCHOR.MIDDLE)
        cy += ((len(sedes) + por_fila - 1) // por_fila) * 0.40 + 0.02
        leyenda = [("● ", True, NIVEL_RGB["Rojo"]), ("Rojo   ", False, P_GREY),
                   ("● ", True, NIVEL_RGB["Naranja"]), ("Naranja   ", False, P_GREY),
                   ("● ", True, NIVEL_RGB["Amarillo"]), ("Amarillo   ", False, P_GREY),
                   ("● ", True, NIVEL_RGB[""]), ("Sin aviso", False, P_GREY)]
        if rz["enfen"]:
            leyenda = [("ENFEN: ", True, P_INK), (f"{rz['enfen']['estado']} · próximo comunicado "
                                                  f"{rz['enfen']['proximo']}", False, P_INK)]
            texto(slide, x + 0.15, cy, w - 0.3, 0.2, [leyenda], 8, nombre="Radar ENFEN")
    texto(slide, x + 0.15, y + h - 0.21, w - 0.3, 0.16, "Fuente: Observatorio Pacífico / Torre de Control",
          7, P_GREY, False, nombre="Radar fuente")



# ---------------- Mapa (propuesta "mapa")
COORD_SEDE = {  # lat, lon aproximadas de la ciudad
    "Piura": (-5.19, -80.63), "Chiclayo": (-6.77, -79.84), "Trujillo": (-8.11, -79.03),
    "Huancayo": (-12.07, -75.21), "Arequipa": (-16.40, -71.54), "Cusco": (-13.53, -71.97),
    "Lima": (-12.05, -77.04), "Ica": (-14.07, -75.73), "Tumbes": (-3.57, -80.45), "Cajamarca": (-7.16, -78.51),
    "Iquitos": (-3.75, -73.25), "Tarapoto": (-6.49, -76.37), "Tacna": (-18.01, -70.25), "Puno": (-15.84, -70.02),
    "Juliaca": (-15.50, -70.13), "Chimbote": (-9.07, -78.59), "Huaraz": (-9.53, -77.53), "Pucallpa": (-8.38, -74.55),
    "Ayacucho": (-13.16, -74.22), "Huánuco": (-9.93, -76.24), "Sullana": (-4.90, -80.69),
}
# hacia dónde va la etiqueta de cada ciudad (para que no se encimen)
LADO_ETIQUETA = {"Huancayo": "der", "Cusco": "der", "Arequipa": "der", "Iquitos": "der", "Tarapoto": "der",
                 "Puno": "der", "Juliaca": "der", "Pucallpa": "der", "Ayacucho": "der", "Huánuco": "der"}
FILL_NIVEL = {"Rojo": (214, 48, 79), "Naranja": (240, 140, 50), "Amarillo": (246, 210, 80), "": (226, 232, 240)}
INTENSIDAD = {"Rojo": "Extremo", "Naranja": "Alto", "Amarillo": "Moderado", "": "Sin información"}
PIN_NIVEL = {"ok": (30, 122, 69), "leve": (214, 128, 0), "afectado": (196, 18, 63)}


def ruta_geojson():
    return recurso("peru_departamentos.geojson")


def dibujar_mapa_png(niveles_dep, pines, alto_px=1500):
    """Silueta del Perú por departamentos coloreada por nivel oficial, con las sedes como puntos.
    niveles_dep: {"Lima": "Naranja", ...};  pines: [(nombre, nivel_impacto)]  -> BytesIO PNG."""
    import json
    import math
    ruta = ruta_geojson()
    if not ruta:
        return None
    with open(ruta, encoding="utf-8") as f:
        geo = json.load(f)
    lon0, lon1, lat0, lat1 = -81.45, -68.6, -18.45, 0.1
    k = math.cos(math.radians(9.5))
    escala = alto_px / (lat1 - lat0)
    margen_izq = int(alto_px * 0.30)                  # espacio en el mar para las etiquetas
    ancho_px = int((lon1 - lon0) * k * escala) + margen_izq + int(alto_px * 0.12)

    def xy(lon, lat):
        return (margen_izq + (lon - lon0) * k * escala, (lat1 - lat) * escala)

    img = Image.new("RGBA", (ancho_px, alto_px), (0, 0, 0, 0))
    dr = ImageDraw.Draw(img)
    norm = {_sin_tildes(k_): v for k_, v in niveles_dep.items()}
    for feat in geo["features"]:
        nombre = feat["properties"].get("NOMBDEP", "")
        nivel = norm.get(_sin_tildes(nombre), "")
        g = feat["geometry"]
        polys = g["coordinates"] if g["type"] == "MultiPolygon" else [g["coordinates"]]
        for poly in polys:
            for anillo in poly[:1]:
                dr.polygon([xy(x, y) for x, y in anillo], fill=FILL_NIVEL[nivel] + (255,))
    for feat in geo["features"]:                       # bordes blancos encima
        g = feat["geometry"]
        polys = g["coordinates"] if g["type"] == "MultiPolygon" else [g["coordinates"]]
        for poly in polys:
            for anillo in poly[:1]:
                dr.line([xy(x, y) for x, y in anillo] + [xy(*anillo[0])], fill=(255, 255, 255, 255),
                        width=max(2, alto_px // 400))
    fuente = ImageFont.truetype(recurso(os.path.join("fuentes", "Nunito-Bold.ttf")), int(alto_px * 0.040)) \
        if recurso(os.path.join("fuentes", "Nunito-Bold.ttf")) else None
    r = int(alto_px * 0.016)
    for nombre, nivel_imp in pines:
        base = re.sub(r"\s*\(.*\)$", "", nombre).strip()
        if base not in COORD_SEDE:
            continue
        x, y = xy(COORD_SEDE[base][1], COORD_SEDE[base][0])
        if fuente:                                   # solo el nombre, centrado en la ciudad
            w = dr.textlength(nombre, font=fuente)
            lado = "izq" if base in ("Lima", "Piura", "Chiclayo", "Trujillo", "Tumbes", "Chimbote", "Ica") \
                else LADO_ETIQUETA.get(base, "centro")
            tx = x - w - 4 if lado == "izq" else (x + 4 if lado == "der" else x - w / 2)
            ty = y - fuente.size * 0.62
            dr.text((tx, ty), nombre, font=fuente, fill=(30, 42, 68, 255),
                    stroke_width=max(3, alto_px // 300), stroke_fill=(255, 255, 255, 255))
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf, img.width / img.height



def pines_sedes(d):
    """Estado de cada sede según el mensaje de agencias (el peor por ciudad): [(nombre, 'ok'|'leve'|'afectado')]."""
    peor, orden_ = {}, {"ok": 0, "leve": 1, "afectado": 2}
    for n_, e_ in (d["agencias"] or [(s_, "Sin afectación") for s_ in SEDES_POR_DEFECTO]):
        base = re.sub(r"\s*\(.*\)$", "", n_).strip()
        base = "Lima" if _sin_tildes(base) == "edificios lima" else base
        nv_ = nivel(e_)
        if orden_[nv_] >= orden_.get(peor.get(base, "ok"), 0):
            peor[base] = nv_
    return list(peor.items())


def dibujar_panel_mapa(slide, d, rz, impactos, px, py, pw, ph):
    """Panel 'Impacto y exposición': estados de Sedes y Colaboradores a la izquierda,
    mapa del Perú (alerta oficial por departamento + sedes) a la derecha."""
    panel = caja(slide, px, py, pw, ph, P_TINT, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Panel impacto")
    panel.adjustments[0] = 0.06
    niveles = dict(rz["lluvias"])
    for dep, nv in rz["quebradas"].items():
        if NIVEL_ORDEN[nv] > NIVEL_ORDEN.get(niveles.get(dep), 0):
            niveles[dep] = nv
    res = dibujar_mapa_png(niveles, pines_sedes(d), alto_px=1600)
    mh = ph - 0.24
    mw = 0.0
    sw = lambda nv: ("■\u00a0", True, RGBColor(*FILL_NIVEL[nv]))
    if res:
        buf, ratio = res
        mw = mh * ratio
        mx = px + pw - 0.15 - mw
        slide.shapes.add_picture(buf, Inches(mx), Inches(py + 0.12), Inches(mw), Inches(mh))
        # leyenda de intensidad en el mar, abajo a la izquierda del mapa
        ley = [[sw(nv), (INTENSIDAD[nv], False, P_INK)] for nv in ("Rojo", "Naranja", "Amarillo", "")]
        texto(slide, mx - 0.35, py + ph - 0.80, 1.3, 0.72, ley, 8, nombre="Leyenda intensidad")
        mw += 0.35
    lx = px + 0.2
    la = pw - 0.35 - mw - 0.25
    texto(slide, lx, py + 0.1, la, 0.26, "IMPACTO Y EXPOSICIÓN · PACÍFICO SEGUROS", 11.5, P_BLUE_DK, "xb")
    y = py + 0.38
    for etq, ic, col, valor in impactos:
        v = valor or "Sin afectación reportada."
        af = nivel(v)
        if etq == "Colaboradores" and d.get("colab_nivel"):
            af = d["colab_nivel"]
        if af != "ok":
            ic, col = _alerta, (G_AMBER if af == "leve" else G_RED)
            etq = "Sedes afectadas" if etq == "Sedes" else "Colaboradores afectados"
        icono(slide, icono_circulo(ic, col), lx, y + 0.02, 0.38)
        ancho_chip = la - 0.48
        fs_chip = 10.5
        while fs_chip > 7 and contar_lineas(etq, v, fs_chip, (ancho_chip - 0.3) * 72) * fs_chip * 1.2 / 72 > 0.34:
            fs_chip -= 0.5
        chip(slide, lx + 0.48, y, ancho_chip, 0.36, etq, v, af, fs_chip)
        y += 0.41
        if etq.startswith("Sedes"):                 # cuadrícula: ✓ Ciudad (n)
            sedes_l = lista_sedes(d)
            por_fila = 4
            cw = (ancho_chip - 0.15) / por_fila
            for k_, (ciudad, cant, nv) in enumerate(sedes_l):
                marca = ("✓\u00a0", True, OK_FG) if nv == "ok" else ("⚠\u00a0", True, COLORES_NIVEL[nv][1])
                nom = ciudad + (f" ({cant})" if cant > 1 else "")
                texto(slide, lx + 0.62 + (k_ % por_fila) * cw, y - 0.04 + (k_ // por_fila) * 0.2, cw, 0.2,
                      [[marca, (nom, nv != "ok", P_INK if nv == "ok" else COLORES_NIVEL[nv][1])]], 9.5,
                      nombre="Sede " + ciudad)
            y += -(-len(sedes_l) // por_fila) * 0.2 + 0.04
    if d.get("colab_casos"):
        partes = []
        for caso in d["colab_casos"]:
            niv = nivel(caso)
            partes.append([("●\u00a0", True, COLORES_NIVEL["afectado" if niv == "ok" else niv][1]), (caso, False, P_INK)])
        texto(slide, lx + 0.48, y - 0.04, la - 0.48, 0.34, partes, 7, nombre="Colaboradores afectados")
        y += 0.34
    texto(slide, lx, py + ph - 0.22, la, 0.16, "Fuente: Observatorio Pacífico / Torre de Control",
          7, P_GREY, False, nombre="Fuente mapa")


# ================================================================ 5. Construir la diapositiva
# Paleta tomada de la plantilla corporativa de Pacífico (degradado azul → morado → magenta)
P_BLUE = RGBColor(0x07, 0x98, 0xD6)
P_BLUE_DK = RGBColor(0x0A, 0x5F, 0xA8)
P_PURPLE = RGBColor(0x7B, 0x5B, 0xA8)
P_PINK = RGBColor(0xD9, 0x2F, 0x83)
P_INK = RGBColor(0x1E, 0x2A, 0x44)
P_GREY = RGBColor(0x6E, 0x74, 0x80)
P_BG = RGBColor(0xF3, 0xF6, 0xFA)
P_TINT = RGBColor(0xE9, 0xF3, 0xFB)
P_LINE = RGBColor(0xDD, 0xE6, 0xF0)
OK_BG, OK_FG = RGBColor(0xE4, 0xF4, 0xEA), RGBColor(0x1E, 0x7A, 0x45)
BAD_BG, BAD_FG = RGBColor(0xFD, 0xE7, 0xEE), RGBColor(0xC4, 0x12, 0x3F)
WARN_BG, WARN_FG = RGBColor(0xFF, 0xF1, 0xD6), RGBColor(0xA3, 0x5A, 0x00)
COLORES_NIVEL = {"ok": (OK_BG, OK_FG), "leve": (WARN_BG, WARN_FG), "afectado": (BAD_BG, BAD_FG)}

G_BLUE, G_PURPLE, G_PINK = (7, 152, 214), (123, 91, 168), (217, 47, 131)
G_AMBER, G_RED = (214, 128, 0), (196, 18, 63)


def recurso(nombre):
    """Busca una imagen en la carpeta 'recursos' junto al script. Si no está, devuelve None."""
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recursos", nombre)
    return ruta if os.path.exists(ruta) else None


def _mezcla(c1, c2, t):
    return tuple(round(a + (b - a) * t) for a, b in zip(c1, c2))


def fondo_cabecera(ancho_px=2400, alto_px=300):
    """Degradado horizontal azul → morado → magenta con borde inferior ondulado, como la plantilla."""
    import math
    grad = Image.new("RGBA", (ancho_px, alto_px))
    px = grad.load()
    for x in range(ancho_px):
        t = x / (ancho_px - 1)
        c = _mezcla(G_BLUE, G_PURPLE, t / 0.55) if t < 0.55 else _mezcla(G_PURPLE, G_PINK, (t - 0.55) / 0.45)
        for y in range(alto_px):
            k = 1 - 0.08 * (y / alto_px)          # leve sombreado vertical
            px[x, y] = (round(c[0] * k), round(c[1] * k), round(c[2] * k), 255)
    mascara = Image.new("L", (ancho_px, alto_px), 0)
    md = ImageDraw.Draw(mascara)
    borde = [(x, alto_px * 0.80 + alto_px * 0.11 * math.sin(2 * math.pi * x / ancho_px * 1.1 + 0.6))
             for x in range(0, ancho_px + 1, 8)]
    md.polygon([(0, 0), (ancho_px, 0)] + borde[::-1], fill=255)
    grad.putalpha(mascara)
    # brillo ondulado translúcido
    capa = Image.new("RGBA", (ancho_px, alto_px), (0, 0, 0, 0))
    cd = ImageDraw.Draw(capa)
    ola = [(x, alto_px * 0.62 + alto_px * 0.09 * math.sin(2 * math.pi * x / ancho_px * 0.9 + 2.2))
           for x in range(0, ancho_px + 1, 8)]
    cd.polygon(ola + [(ancho_px, alto_px), (0, alto_px)], fill=(255, 255, 255, 34))
    capa.putalpha(Image.composite(capa.getchannel("A"), Image.new("L", capa.size, 0), mascara))
    grad.alpha_composite(capa)
    buf = io.BytesIO()
    grad.save(buf, "PNG")
    buf.seek(0)
    return buf


def icono_circulo(dibujar, fondo, size=256):
    """Ícono blanco dentro de un círculo de color (motivo visual del reporte)."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(img).ellipse((0, 0, size - 1, size - 1), fill=fondo + (255,))
    inner = int(size * 0.62)
    ic = Image.new("RGBA", (inner, inner), (0, 0, 0, 0))
    dibujar(ImageDraw.Draw(ic), (255, 255, 255, 255), inner)
    img.alpha_composite(ic, ((size - inner) // 2, (size - inner) // 2))
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def imagen_ajustada(slide, ruta, x, y, alto=None, ancho=None):
    with Image.open(ruta) as im:
        r = im.width / im.height
    if alto is not None:
        ancho = alto * r
    else:
        alto = ancho / r
    slide.shapes.add_picture(ruta, Inches(x), Inches(y), Inches(ancho), Inches(alto))
    return ancho, alto


def chip(slide, x, y, w, h, etiqueta, valor, afectado, size):
    """afectado: True/False o un nivel ('ok', 'leve', 'afectado')."""
    niv = afectado if isinstance(afectado, str) else ("afectado" if afectado else "ok")
    fondo, tinta = COLORES_NIVEL[niv]
    afectado = niv != "ok"
    c = caja(slide, x, y, w, h, fondo, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Estado " + etiqueta.lower())
    c.adjustments[0] = 0.5
    partes = [[(etiqueta + ": ", True, P_INK), (valor, afectado, tinta)]] if etiqueta else [[(valor, True, tinta)]]
    texto(slide, x + 0.16, y, w - 0.3, h, partes, size, anchor=MSO_ANCHOR.MIDDLE)


def construir(d, salida_pptx, borrador=False):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = P_BG

    # ---- Cabecera con degradado ondulado
    slide.shapes.add_picture(fondo_cabecera(), 0, 0, Inches(13.333), Inches(1.72))
    texto(slide, 0.6, 0.24, 7.5, 0.62, "REPORTE FEN " + d["numero"], 32, WHITE, "xb", nombre="Titulo")
    texto(slide, 0.62, 0.86, 7.0, 0.3, "Reporte fenomenológico nacional", 14, WHITE, True, nombre="Subtitulo")
    if recurso("logo_continuidad.png"):
        imagen_ajustada(slide, recurso("logo_continuidad.png"), 7.55, 0.30, alto=0.62)
    fb = caja(slide, 10.05, 0.24, 2.78, 0.9, WHITE, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Caja fecha")
    fb.adjustments[0] = 0.18
    texto(slide, 10.3, 0.40, 2.4, 0.2, "FECHA", 9, P_GREY, True)
    texto(slide, 10.3, 0.62, 2.4, 0.36, d["fecha"], 18, P_BLUE_DK, "xb")
    if borrador:
        b = caja(slide, 4.6, 0.36, 2.7, 0.42, BAD_FG, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Marca borrador")
        b.adjustments[0] = 0.5
        texto(slide, 4.6, 0.36, 2.7, 0.42, "BORRADOR · SIN CONFIRMAR", 11, WHITE, True,
              PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

    # ---- Contenido de cada columna
    secs = [("LLUVIAS REPORTADAS", _lluvia, G_BLUE, d["lluvias"]),
            ("VÍAS RESTRINGIDAS", _barrera, G_PURPLE, d["vias"]),
            ("RÍOS, QUEBRADAS Y OTRAS AFECTACIONES", _rio, G_PINK, d["rios"])]
    secs = [(t, ic, col, items or ["Sin reportes."]) for t, ic, col, items in secs]
    bloques = [(d["p1_titulo"] or "PRONÓSTICO", P_BLUE_DK, d["p1"] or ["Sin pronóstico."])]
    if d["p2"]:
        bloques.append((d["p2_titulo"] or "NOCHE Y MADRUGADA", P_PINK, d["p2"]))
    secs_base, bloques_base = secs, bloques
    # Secciones opcionales: aparecen solo si el mensaje las trae
    extras = []
    if d["medidas"]:
        extras.append((("MEDIDAS", _megafono, G_PURPLE, d["medidas"]), ("MEDIDAS", P_PURPLE, d["medidas"])))
    if d.get("recomendaciones"):
        extras.append((("RECOMENDACIONES PREVENTIVAS", _alerta, G_AMBER, d["recomendaciones"]),
                       ("RECOMENDACIONES PREVENTIVAS", WARN_FG, d["recomendaciones"])))
    hay_agencias = bool(d["agencias"]) and not (d.get("radar") and d.get("radar_estilo") == "mapa")
    hay_casos = bool(d.get("colab_casos"))

    TOP, BOTTOM = 1.42, 6.86
    HEAD = 0.5
    y0 = TOP + HEAD + 0.14
    EXTRA_AG = 0.27 if hay_agencias else 0.0      # fila de agencias dentro del panel de impacto
    EXTRA_CO = 0.30 if hay_casos else 0.0         # fila de colaboradores afectados
    EXTRA_AG_TOTAL = EXTRA_AG + EXTRA_CO
    es_mapa = bool(d.get("radar")) and d.get("radar_estilo") == "mapa"
    if es_mapa:                                   # panel "Impacto y exposición" con el mapa
        EXTRA_AG_TOTAL = 2.12 + (0.34 if hay_casos else 0.0) - 0.6
    Y1_NORMAL, Y1_COMPACTO = 5.42, 6.08 - EXTRA_AG_TOTAL
    ry0, ry1 = TOP + HEAD + 0.16, 6.02
    MIN_FS = 7.5

    def alto_tit(t, ancho, s=13):
        return contar_lineas("", t.upper(), s, ancho * 72 * 0.995) * s * 1.2 / 72 + 0.08

    def fs_izq(LW, y1):
        lw = LW - 0.86 - 0.28
        need = lambda f: sum(ESP[0] + alto_lista(it, f, lw, ESP_IZQ) + ESP[1] for *_, it in secs)
        fs = 12.5
        while fs > MIN_FS and need(fs) > y1 - y0:
            fs -= 0.5
        return fs, need(fs) <= y1 - y0, max(0.0, need(fs) - (y1 - y0))

    rz, rz_estilo = d.get("radar"), d.get("radar_estilo", "avisos")

    def alto_bloque_radar(RW):
        return (alto_radar(rz, rz_estilo, RW - 0.44) + 0.10) if (rz and rz_estilo != "mapa") else 0.0

    ESP_DER = 9          # puntos entre viñetas del pronóstico
    ESP_IZQ = 2          # puntos entre viñetas del estatus (poco: así la letra puede ser más grande)

    def fs_der(RW):
        rw = RW - 0.56
        y1d = ry1 - alto_bloque_radar(RW)           # el bloque del Radar ocupa la parte baja de la columna
        need = lambda f: sum(alto_tit(t, rw) + alto_lista(it, f, rw, ESP_DER) + 0.22 for t, _, it in bloques)
        fs = 12.5
        while fs > MIN_FS and need(fs) > y1d - ry0:
            fs -= 0.5
        return fs, need(fs) <= y1d - ry0, max(0.0, need(fs) - (y1d - ry0))

    LX, GAP, ANCHO_TOTAL = 0.5, 0.22, 12.33

    def elegir(y1):
        mejor = None
        # Reparte el ancho entre columnas: primero que todo entre; luego, la letra más grande;
        # si no entra en ningún reparto, el que menos se desborde.
        for LW in (7.75, 7.9, 8.05, 8.2, 8.35, 8.5, 8.65, 7.5, 7.25, 7.0, 6.75, 6.5):
            RW = ANCHO_TOTAL - LW - GAP
            if RW < 3.45:                      # por debajo, el título del pronóstico ya no entra legible
                continue
            (a, ok_a, ov_a), (b, ok_b, ov_b) = fs_izq(LW, y1), fs_der(RW)
            ok = ok_a and ok_b
            clave = (ok, min(a, b) if ok else -(ov_a + ov_b), a + b, -abs(LW - 7.75))
            if mejor is None or clave > mejor[0]:
                mejor = (clave, LW, RW, a, b)
        return mejor

    compacto = True                     # un solo diseño: impacto en una fila
    # Espacio de cada sección de la izquierda: (título → lista, entre secciones).
    # Si con el espaciado normal no entra, se prueba uno más compacto.
    ESPACIADOS = ((0.34, 0.22), (0.32, 0.14), (0.30, 0.07)) if es_mapa else ((0.34, 0.12), (0.30, 0.07))
    ESP = ESPACIADOS[0]
    # Cada sección opcional va a la columna (izquierda o derecha) donde deje la letra más grande
    import itertools
    opciones = []
    for lados in itertools.product(("izq", "der"), repeat=len(extras)):
        secs_o = secs_base + [e[0] for e, lado in zip(extras, lados) if lado == "izq"]
        bloques_o = bloques_base + [e[1] for e, lado in zip(extras, lados) if lado == "der"]
        opciones.append((secs_o, bloques_o))
    # Con el mapa, si la columna izquierda no entra, el panel de impacto se hace más bajo.
    alturas = [EXTRA_AG_TOTAL - r for r in (0.0, 0.08, 0.16, 0.24)] if es_mapa else [EXTRA_AG_TOTAL]
    for EXTRA_AG_TOTAL in alturas:
        Y1_COMPACTO = 6.08 - EXTRA_AG_TOTAL
        mejor = None
        for ESP in ESPACIADOS:
            for secs, bloques in opciones:
                m = elegir(Y1_COMPACTO)
                if mejor is None or m[0] > mejor[0][0]:
                    mejor = (m, secs, bloques, ESP)
            if mejor[0][0][0] and mejor[0][0][1] >= 9:      # entra con letra legible: no compactar
                break
        if mejor[0][0][0]:                                  # entra: no achicar más el panel
            break
    (m, secs, bloques, ESP) = mejor
    (cabe, *_), LW, RW, fs, fs2 = m
    RX = LX + LW + GAP

    # ---- Tarjetas blancas con encabezado tenue
    for x, w, titulo, nom, color in ((LX, LW, "ESTATUS DEL DÍA", "Tarjeta estatus", P_BLUE_DK),
                                     (RX, RW, "PRONÓSTICO DE LLUVIAS", "Tarjeta pronostico", P_PINK)):
        c = caja(slide, x, TOP, w, BOTTOM - TOP, WHITE, P_LINE, MSO_SHAPE.ROUNDED_RECTANGLE, nom)
        c.adjustments[0] = 0.035
        sombra_suave(c)
        h = caja(slide, x, TOP, w, HEAD, P_TINT, None, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, nom + " encabezado")
        h.adjustments[0] = 0.32
        h.adjustments[1] = 0.0
        fs_t = 16.0
        while fs_t > 13 and ancho_pt(titulo, fs_t, True) * 1.06 > (w - 0.5) * 72:
            fs_t -= 0.5
        texto(slide, x + 0.3, TOP, w - 0.5, HEAD, titulo, fs_t, color, "xb", anchor=MSO_ANCHOR.MIDDLE)

    # ---- Columna izquierda: estatus del día
    lista_x, lista_w = LX + 0.86, LW - 0.86 - 0.28
    y = y0
    usado = sum(ESP[0] + alto_lista(it, fs, lista_w, ESP_IZQ) + ESP[1] for *_, it in secs)
    aire = min(0.28, max(0.0, (Y1_COMPACTO - y0 - usado) / max(1, len(secs))))   # reparte el espacio libre
    for i, (titulo, ic, col, items) in enumerate(secs):
        icono(slide, icono_circulo(ic, col), LX + 0.3, y - 0.02, 0.4 if ESP[0] >= 0.34 else 0.36)
        texto(slide, lista_x, y + 0.02, lista_w, 0.26, titulo, 12.5, P_INK, "xb")
        h = alto_lista(items, fs, lista_w, ESP_IZQ)
        vinetas(slide, lista_x, y + ESP[0], lista_w, h + 0.02, items, fs, "Lista " + titulo.lower(), esp=ESP_IZQ)
        y += ESP[0] + h + ESP[1] + aire
        if i < len(secs) - 1:
            linea_h(slide, lista_x, LX + LW - 0.28, y - (ESP[1] + aire) / 2)

    # ---- Impacto en Pacífico (panel tenue con estados en chip)
    impactos = (("Sedes", _edificio, G_BLUE, d["sedes"]),
                ("Colaboradores", _personas, G_PINK, d["colab"]))

    if not compacto:
        px, py, pw, ph = LX + 0.22, 5.52, LW - 0.44, 1.2
        panel = caja(slide, px, py, pw, ph, P_TINT, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Panel impacto")
        panel.adjustments[0] = 0.12
        texto(slide, px + 0.22, py + 0.1, pw - 0.4, 0.3, "IMPACTO EN PACÍFICO SEGUROS", 13, P_BLUE_DK, True)
        col_w = (pw - 0.44) / 2
        for j, (etq, ic, col, valor) in enumerate(impactos):
            cx = px + 0.22 + j * col_w
            icono(slide, icono_circulo(ic, col), cx, py + 0.52, 0.48)
            texto(slide, cx + 0.6, py + 0.48, col_w - 0.75, 0.24, etq.upper(), 11, P_INK, True)
            v = valor or "Sin afectación reportada."
            af = es_afectado(v)
            chip(slide, cx + 0.6, py + 0.76, col_w - 0.75, 0.34, "", v, af, 10 if len(v) < 45 else 8.5)
    else:
        px, py, pw, ph = LX + 0.22, 6.14 - EXTRA_AG_TOTAL, LW - 0.44, 0.6 + EXTRA_AG_TOTAL
    if es_mapa:
        dibujar_panel_mapa(slide, d, d["radar"], impactos, px, py, pw, ph)
    elif compacto:
        panel = caja(slide, px, py, pw, ph, P_TINT, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Panel impacto")
        panel.adjustments[0] = 0.25 if not (hay_agencias or hay_casos) else 0.14
        texto(slide, px + 0.2, py, 2.1, 0.6, [[("IMPACTO EN", True, P_BLUE_DK)], [("PACÍFICO SEGUROS", True, P_BLUE_DK)]],
              11, anchor=MSO_ANCHOR.MIDDLE)
        resto = pw - 2.35
        col_w = resto / 2
        for j, (etq, ic, col, valor) in enumerate(impactos):
            cx = px + 2.25 + j * col_w
            v = valor or "Sin afectación reportada."
            af = nivel(v)
            if etq == "Colaboradores" and d.get("colab_nivel"):
                af = d["colab_nivel"]                    # nivel calculado de los casos reportados
            if af != "ok":
                # Hay afectación: ícono de alerta (ámbar si es leve, rojo si es seria) y etiqueta explícita
                ic, col = _alerta, (G_AMBER if af == "leve" else G_RED)
                etq = "Sedes afectadas" if etq == "Sedes" else "Colaboradores afectados"
            icono(slide, icono_circulo(ic, col), cx, py + 0.11, 0.38)
            ancho_chip = col_w - 0.56
            fs_chip = 9.5
            while fs_chip > 7 and contar_lineas(etq, v, fs_chip, (ancho_chip - 0.3) * 72) * fs_chip * 1.2 / 72 > 0.34:
                fs_chip -= 0.5
            chip(slide, cx + 0.46, py + 0.1, ancho_chip, 0.4, etq, v, af, fs_chip)
        if hay_agencias:
            # Fila con las agencias/sedes SIN afectación (las afectadas ya se muestran arriba).
            # La fecha y el corte del mensaje de agencias quedan en las notas del PPT.
            ok = [n for n, e in d["agencias"] if nivel(e) == "ok"]
            total = len(d["agencias"])
            if not ok:
                partes = [("Ninguna agencia o sede sin afectación.", True, BAD_FG)]
            else:
                rotulo = f"Sin afectación ({len(ok)}):  "
                partes = [(rotulo, True, OK_FG)]
                for n in ok:
                    partes += [("●\u00a0", False, OK_FG), (n + "   ", False, P_INK)]
            ancho_ag = (pw - 0.4) * 72 * 0.97
            fs_ag = 9.5
            while fs_ag > 6.5 and sum(ancho_pt(t, fs_ag, b) for t, b, _ in partes) > ancho_ag:
                fs_ag -= 0.5
            texto(slide, px + 0.2, py + 0.58, pw - 0.4, 0.22, [partes], fs_ag, anchor=MSO_ANCHOR.MIDDLE,
                  nombre="Agencias y sedes")
        if hay_casos:
            # Fila con el detalle de cada caso de colaboradores afectados
            tinta_rot = COLORES_NIVEL[d.get("colab_nivel") or "afectado"][1]
            partes = [("Colaboradores afectados:  ", True, tinta_rot)]
            for caso in d["colab_casos"]:
                niv = nivel(caso)
                tinta = COLORES_NIVEL["afectado" if niv == "ok" else niv][1]
                partes += [("●\u00a0", False, tinta), (caso + "    ", False, P_INK)]
            ancho_co = (pw - 0.4) * 72 * 0.97
            largo = lambda f: sum(ancho_pt(t, f, b) for t, b, _ in partes)
            fs_co = 9.5
            while fs_co > 7 and largo(fs_co) > ancho_co:
                fs_co -= 0.5
            if largo(fs_co) > ancho_co:                  # no entra en una línea: dos líneas a 7 pt
                fs_co = 7
            texto(slide, px + 0.2, py + 0.58 + EXTRA_AG, pw - 0.4, 0.28, [partes], fs_co,
                  anchor=MSO_ANCHOR.MIDDLE, nombre="Colaboradores afectados")

    # ---- Columna derecha: pronóstico
    rl_x, rl_w = RX + 0.28, RW - 0.56
    y = ry0
    for i, (t, col, items) in enumerate(bloques):
        ht = alto_tit(t, rl_w)
        texto(slide, rl_x, y, rl_w, ht, t.upper(), 13, col, "xb")
        h = alto_lista(items, fs2, rl_w, ESP_DER)
        vinetas(slide, rl_x, y + ht, rl_w, h + 0.02, items, fs2,
                t.capitalize() if i >= len(bloques_base) else f"Pronostico {i + 1}", esp=ESP_DER)
        y += ht + h + 0.22
        if i < len(bloques) - 1:
            linea_h(slide, rl_x, RX + RW - 0.28, y - 0.11)
    if rz and rz_estilo != "mapa":
        hr = alto_radar(rz, rz_estilo, RW - 0.44)
        dibujar_radar(slide, rz, rz_estilo, RX + 0.22, 6.14 - 0.08 - hr, RW - 0.44, hr)
    nb = caja(slide, RX + 0.22, 6.14, RW - 0.44, 0.6, P_TINT, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Nota monitoreo")
    nb.adjustments[0] = 0.25
    icono(slide, icono_circulo(_radar, G_PURPLE), RX + 0.34, 6.25, 0.38)
    nota = d["nota"] or "Se mantiene el monitoreo."
    ancho_nota = RW - 1.2
    fs_nota = 11.0
    while fs_nota > 7.5 and contar_lineas("", nota, fs_nota, ancho_nota * 72, ) * fs_nota * 1.2 / 72 > 0.54:
        fs_nota -= 0.5
    texto(slide, RX + 0.86, 6.14, ancho_nota, 0.6, nota, fs_nota, P_INK, True, anchor=MSO_ANCHOR.MIDDLE)
    if not cabe:
        print("  Aviso: hay mucho texto; revisa en el PPT que nada quede cortado.")

    # ---- Pie: "Pacífico Seguros" a la izquierda; torito y logo de Pacífico a la derecha
    # (la sección [PIE DE PAGINA] del .txt ya no se usa; se acepta para no romper archivos antiguos)
    texto(slide, 0.6, 6.98, 8.5, 0.42, "Pacífico Seguros", 10, P_GREY, True, anchor=MSO_ANCHOR.MIDDLE,
          nombre="Pie")
    x_logo = 12.83
    if recurso("logo_pacifico.png"):
        w, _ = imagen_ajustada(slide, recurso("logo_pacifico.png"), 0, 7.0, alto=0.4)
        slide.shapes[-1].left = Inches(x_logo - w)
        x_logo -= w + 0.18
    if recurso("torito.png"):
        w, _ = imagen_ajustada(slide, recurso("torito.png"), 0, 6.88, alto=0.62)
        slide.shapes[-1].left = Inches(x_logo - w)

    # Sin notas en la diapositiva (el PPT se comparte tal cual)

    prs.save(salida_pptx)


# ================================================================ 6. Exportar PNG
def exportar_png(pptx_path, png_path):
    pptx_path, png_path = os.path.abspath(pptx_path), os.path.abspath(png_path)

    # Opción A: PowerPoint en Windows (resultado idéntico al PPT)
    if platform.system() == "Windows":
        tmpdir = tempfile.mkdtemp(prefix="reporte_fen_")
        pres = None
        try:
            import win32com.client  # pip install pywin32
            # Se exporta desde una copia con nombre único: si PowerPoint todavía tiene abierta
            # una versión anterior del mismo reporte, no la reutiliza y la imagen sale actualizada.
            copia = os.path.join(tmpdir, f"reporte_{datetime.now():%H%M%S%f}.pptx")
            shutil.copy2(pptx_path, copia)
            png_tmp = os.path.join(tmpdir, "reporte.png")
            app = win32com.client.Dispatch("PowerPoint.Application")
            pres = app.Presentations.Open(copia, ReadOnly=True, WithWindow=False)
            pres.Slides(1).Export(png_tmp, "PNG", 3840, 2160)   # 4K: se ve nítido aunque WhatsApp lo comprima
            pres.Close()
            pres = None
            if os.path.exists(png_path):
                os.remove(png_path)
            shutil.move(png_tmp, png_path)
            return "PowerPoint"
        except PermissionError:
            print(f"  No se pudo reemplazar {os.path.basename(png_path)}: ciérralo si está abierto y vuelve a ejecutar.")
            return None
        except Exception as e:  # noqa: BLE001
            print(f"  (No se pudo usar PowerPoint: {e}. Probando con LibreOffice...)")
        finally:
            if pres is not None:
                try:
                    pres.Close()
                except Exception:  # noqa: BLE001
                    pass
            shutil.rmtree(tmpdir, ignore_errors=True)

    # Opción B: LibreOffice
    soffice = os.environ.get("FEN_SOFFICE")
    cmd = soffice.split() if soffice else None
    if not cmd:
        exe = shutil.which("soffice") or shutil.which("libreoffice")
        for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                  "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
            if not exe and os.path.exists(c):
                exe = c
        cmd = [exe] if exe else None
    if not cmd:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(cmd + ["--headless", "--convert-to", "pdf", "--outdir", tmp, pptx_path],
                       check=True, capture_output=True, timeout=180)
        pdf = os.path.join(tmp, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
        if shutil.which("pdftoppm"):
            base = os.path.join(tmp, "img")
            subprocess.run(["pdftoppm", "-png", "-r", "288", "-singlefile", pdf, base], check=True)
            shutil.move(base + ".png", png_path)
        else:
            subprocess.run(cmd + ["--headless", "--convert-to", "png", "--outdir", tmp, pptx_path],
                           check=True, capture_output=True, timeout=180)
            shutil.move(os.path.join(tmp, os.path.splitext(os.path.basename(pptx_path))[0] + ".png"), png_path)
    return "LibreOffice"


# ================================================================ 6b. Correo (Outlook)
def leer_destinatarios(ruta):
    """Lee destinatarios.txt: líneas 'PARA: a@x.com; b@x.com' y 'CC: ...' (se pueden repetir).
    Opcionales: 'SALUDO: ...' y 'FIRMA: ...'. Las líneas con # se ignoran."""
    conf = {"desde": "", "para": [], "cc": [], "saludo": "Estimados:", "firma": "Saludos,"}
    if not os.path.exists(ruta):
        return conf
    with open(ruta, encoding="utf-8-sig") as f:
        for raw in f:
            linea = raw.strip()
            if not linea or linea.startswith("#"):
                continue
            k, sep, v = linea.partition(":")
            k = k.strip().upper()
            if not sep:
                continue
            if k in ("DESDE", "DE", "FROM", "BUZON", "BUZÓN"):
                conf["desde"] = v.strip()
            elif k in ("PARA", "TO"):
                conf["para"] += [x.strip() for x in re.split(r"[;,]", v) if x.strip()]
            elif k == "CC":
                conf["cc"] += [x.strip() for x in re.split(r"[;,]", v) if x.strip()]
            elif k == "SALUDO" and v.strip():
                conf["saludo"] = v.strip()
            elif k == "FIRMA" and v.strip():
                conf["firma"] = v.strip()
    return conf


def fecha_puntos(fecha):
    """'4/11/2026' -> '04.11.2026'"""
    m = re.match(r"^\s*(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})\s*$", fecha or "")
    return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}.{m.group(3)}" if m else (fecha or "")


def saludo_por_hora(ahora=None):
    h = (ahora or datetime.now()).hour
    return "Estimados, buenos días:" if h < 12 else ("Estimados, buenas tardes:" if h < 19
                                                     else "Estimados, buenas noches:")


def armar_correo(d, conf, img_src="cid:reporte_fen"):
    """Devuelve (asunto, cuerpo_html) con el formato del correo del Comité FEN.
    img_src: 'cid:...' (Outlook) o 'data:image/...;base64,...' (Power Automate)."""
    import html
    asunto = f"Reporte FEN {d['numero']} | {fecha_puntos(d['fecha'])}"
    saludo = conf.get("saludo") or ""
    if not saludo or saludo == "Estimados:":
        saludo = saludo_por_hora()
    firma = conf.get("firma") or "Saludos,"
    punto = {"ok": "🟢", "leve": "🟠", "afectado": "🔴"}

    def estado(etq, valor):
        v = valor or "Sin afectación reportada."
        niv = d["colab_nivel"] if (etq == "Colaboradores" and d.get("colab_nivel")) else nivel(v)
        linea = f'<p style="margin:6px 0">{etq}: {punto[niv]} {html.escape(v)}</p>'
        if etq == "Colaboradores" and d.get("colab_casos"):
            linea += ('<ul style="margin:2px 0 6px 0">'
                      + "".join(f"<li>{html.escape(c)}</li>" for c in d["colab_casos"]) + "</ul>")
        return linea

    cuerpo = f"""<div style="font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#000000">
<p>{html.escape(saludo)}</p>
<p>Según lo coordinado en el último Comité FEN, compartimos el estado de afectación asociados en los frentes de
Infraestructura y Colaboradores, con el objetivo de mantener visibilidad sobre cualquier impacto que pudiera presentarse.</p>
<p>&nbsp;</p>
<p style="margin:6px 0">Estado actual:</p>
{estado("Infraestructura", d["sedes"])}
{estado("Colaboradores", d["colab"])}
<p><img src="{img_src}" alt="Reporte FEN {html.escape(d['numero'])}" width="900" style="width:900px;max-width:100%;height:auto"></p>
<p>{html.escape(firma)}</p>
</div>"""
    return asunto, cuerpo


def imagen_para_correo_b64(png_path, ancho=1920):
    """Versión liviana del PNG para incrustar en el cuerpo del correo (JPEG nítido, ~200 KB)."""
    import base64
    with Image.open(png_path) as im:
        im = im.convert("RGB")
        if im.width > ancho:
            im = im.resize((ancho, round(im.height * ancho / im.width)), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=86, subsampling=0, optimize=True)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def enviar_outlook(asunto, cuerpo, para, cc, png_path, pptx_path, solo_borrador=False, desde=""):
    """Crea el correo en Outlook con la imagen incrustada y el PPT adjunto. Envía o abre el borrador.
    desde: buzón remitente (p. ej. el de Continuidad). Si ese buzón está agregado como cuenta en
    Outlook se usa esa cuenta; si no, se envía "en nombre de" / "como" el buzón (requiere permiso)."""
    import win32com.client  # pip install pywin32
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    if desde:
        cuenta = None
        for acc in outlook.Session.Accounts:
            try:
                if (acc.SmtpAddress or "").lower() == desde.lower():
                    cuenta = acc
                    break
            except Exception:  # noqa: BLE001
                continue
        if cuenta is not None:
            try:
                mail.SendUsingAccount = cuenta
            except Exception:  # noqa: BLE001  (algunas versiones de pywin32 lo requieren así)
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta))
        else:
            mail.SentOnBehalfOfName = desde
    mail.To = "; ".join(para)
    mail.CC = "; ".join(cc)
    mail.Subject = asunto
    if png_path and os.path.exists(png_path):
        adj = mail.Attachments.Add(os.path.abspath(png_path))
        # Content-ID para mostrar la imagen dentro del correo
        adj.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "reporte_fen")
    mail.Attachments.Add(os.path.abspath(pptx_path))
    mail.HTMLBody = cuerpo
    if solo_borrador:
        mail.Display()
    else:
        mail.Send()


def abrir_para_revisar(ruta):
    """Abre el PNG (o el PPT) con el programa predeterminado para que lo revises."""
    try:
        if platform.system() == "Windows":
            os.startfile(ruta)  # noqa: S606
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:  # noqa: BLE001
        print(f"  (No se pudo abrir automáticamente; ábrelo tú: {ruta})")


def paso_correo(d, pptx_path, png_path, carpeta_script, carpeta_salida, auto=None):
    """Paso de validación: muestra el reporte, pide el OK y recién entonces envía el correo."""
    conf = leer_destinatarios(os.path.join(carpeta_script, "destinatarios.txt"))
    if not conf["para"]:
        print("\nCorreo: no se envía porque destinatarios.txt no existe o no tiene ninguna línea 'PARA:'.")
        return
    if not conf["desde"]:
        print("\nCorreo: no se envía porque falta la línea 'DESDE:' (buzón de Continuidad) en destinatarios.txt.")
        return
    if platform.system() != "Windows":
        print("\nCorreo: el envío usa Outlook de Windows; en este equipo no está disponible.")
        return

    asunto, cuerpo = armar_correo(d, conf)
    revisar = png_path if (png_path and os.path.exists(png_path)) else pptx_path
    print("\n" + "=" * 70)
    print("REVISIÓN ANTES DE ENVIAR")
    print(f"  Se abrió el reporte para que lo revises: {os.path.basename(revisar)}")
    print(f"  Desde:  {conf['desde']}")
    print(f"  Asunto: {asunto}")
    print(f"  Para:   {'; '.join(conf['para'])}")
    if conf["cc"]:
        print(f"  CC:     {'; '.join(conf['cc'])}")
    print("=" * 70)
    abrir_para_revisar(revisar)

    if auto is None:
        resp = ""
        while resp not in ("s", "b", "n"):
            resp = input("¿Das el OK?  [s] enviar ahora   [b] abrir borrador en Outlook   [n] no enviar: ")
            resp = resp.strip().lower()[:1]
    else:
        resp = auto

    if resp == "n":
        print("No se envió el correo. El PPT y el PNG quedan en 'Reportes generados'.")
        return
    try:
        enviar_outlook(asunto, cuerpo, conf["para"], conf["cc"], png_path, pptx_path,
                       solo_borrador=(resp == "b"), desde=conf["desde"])
    except Exception as e:  # noqa: BLE001
        print(f"No se pudo crear el correo en Outlook: {e}")
        print("Revisa que Outlook esté abierto y con tu cuenta iniciada, y vuelve a intentarlo.")
        return
    accion = "Correo ENVIADO" if resp == "s" else "Borrador abierto en Outlook (envíalo desde ahí)"
    print(f"{accion}: {asunto}  (desde {conf['desde']})")
    if resp == "s":
        print("  Si en unos minutos te llega un aviso de 'no tiene permiso para enviar en nombre de…', "
              "tu usuario aún no tiene permiso sobre el buzón (ver destinatarios.txt).")
    # Registro de envíos
    try:
        with open(os.path.join(carpeta_salida, "registro_envios.txt"), "a", encoding="utf-8") as f:
            conf_por = " y ".join(c for c in d["conf"] if c)
            f.write(f"{datetime.now():%d/%m/%Y %H:%M} | {accion.split(' (')[0]} | {asunto} | "
                    f"Desde: {conf['desde']} | Para: {'; '.join(conf['para'])} | CC: {'; '.join(conf['cc'])} | Confirmado por: {conf_por}\n")
    except OSError:
        pass


# ================================================================ 6c. Power Automate
def leer_configuracion(carpeta_script):
    """configuracion.txt (opcional) junto al script:
        CARPETA_SALIDA: ruta donde guardar PPT y PNG (p. ej. una carpeta sincronizada con OneDrive)
        MODO_CORREO: automate | outlook | ninguno
    Sin el archivo: se guarda en 'Reportes generados' y se usa el modo outlook (destinatarios.txt)."""
    conf = {"salida": "", "modo": "outlook", "radar": "si", "radar_estilo": "mapa"}
    ruta = os.path.join(carpeta_script, "configuracion.txt")
    if not os.path.exists(ruta):
        return conf
    with open(ruta, encoding="utf-8-sig") as f:
        for raw in f:
            linea = raw.strip()
            if not linea or linea.startswith("#"):
                continue
            k, sep, v = linea.partition(":")
            if not sep:
                continue
            k, v = k.strip().upper(), v.strip().strip('"')
            if k == "CARPETA_SALIDA" and v:
                conf["salida"] = os.path.expandvars(os.path.expanduser(v))
            elif k == "MODO_CORREO" and v:
                conf["modo"] = v.lower()
            elif k == "RADAR" and v:
                conf["radar"] = v.lower()
            elif k == "RADAR_ESTILO" and v:
                conf["radar_estilo"] = v.lower()
    return conf


def revisar_imagen(png_path):
    """Devuelve '' si el PNG está completo; si no, una descripción corta del problema."""
    try:
        with Image.open(png_path) as im:
            im.verify()                      # detecta archivos cortados o dañados
        with Image.open(png_path) as im:
            im.load()
            ancho, alto = im.size
            if ancho < 800 or alto < 400:
                return f"tamaño {ancho}x{alto}"
            # una imagen "vacía" (casi de un solo color) también es sospechosa
            colores = im.convert("RGB").resize((64, 36)).getcolors(64 * 36)
            if colores is not None and len(colores) < 20:
                return "imagen en blanco"
    except Exception as e:  # noqa: BLE001
        return f"no se pudo leer: {e.__class__.__name__}"
    return ""


def ruta_en_onedrive(ruta_local):
    """Convierte una ruta local dentro de OneDrive en la ruta que usa Power Automate
    ('/Reporte FEN/Reportes generados/archivo.pptx'). Devuelve '' si no está en OneDrive."""
    ruta = os.path.abspath(ruta_local)
    for var in ("OneDriveCommercial", "OneDrive"):
        raiz = os.environ.get(var)
        if raiz and os.path.normcase(ruta).startswith(os.path.normcase(os.path.abspath(raiz)) + os.sep):
            return "/" + os.path.relpath(ruta, raiz).replace("\\", "/")
    # Sin variables de entorno: busca la carpeta "OneDrive - ..." en la ruta
    partes = re.split(r"[\\/]", ruta)
    for i, p in enumerate(partes):
        if p.lower().startswith("onedrive"):
            return "/" + "/".join(partes[i + 1:])
    return ""


def buscar_salida_onedrive():
    """Si CARPETA_SALIDA está vacía, busca en tu OneDrive la carpeta que ya usa el flujo
    (la que contiene '_para_envio'). Devuelve la ruta o ''."""
    candidatas = []
    for var in ("OneDriveCommercial", "OneDrive"):
        raiz = os.environ.get(var)
        if not raiz or not os.path.isdir(raiz):
            continue
        base_prof = raiz.rstrip("\\/").count(os.sep)
        for dirpath, dirnames, _ in os.walk(raiz):
            if dirpath.count(os.sep) - base_prof >= 4:
                dirnames[:] = []
                continue
            if "_para_envio" in dirnames:
                candidatas.append(dirpath)
            dirnames[:] = [n for n in dirnames if not n.startswith((".", "_vista"))]
    if not candidatas:
        return ""
    return max(candidatas, key=lambda c: os.path.getmtime(os.path.join(c, "_para_envio")))


def preparar_envio_automate(d, pptx_path, png_path, carpeta_salida):
    """Deja en <salida>/_para_envio un .json con todo lo que el flujo de Power Automate necesita
    (asunto, cuerpo, resumen e imagen y PPT en base64). El flujo se dispara con ese archivo."""
    import base64
    import html
    import json
    carpeta = os.path.join(carpeta_salida, "_para_envio")
    os.makedirs(carpeta, exist_ok=True)
    corte = re.sub(r"^corte\s*:\s*", "", d["horario"], flags=re.I).strip()
    img_b64 = imagen_para_correo_b64(png_path) if png_path and os.path.exists(png_path) else ""
    asunto, cuerpo = armar_correo(d, {}, img_src=f"data:image/jpeg;base64,{img_b64}")
    # HTML en una sola línea y sin espacios en los extremos (el editor de Power Automate lo trata mejor)
    cuerpo = re.sub(r">\s+<", "><", re.sub(r"\s*\n\s*", " ", cuerpo)).strip()
    s_val = d["sedes"] or "Sin afectación reportada."
    c_val = d["colab"] or "Sin afectación reportada."
    s_niv, c_niv = nivel(s_val), (d.get("colab_nivel") or nivel(c_val))
    punto = {"ok": "🟢", "leve": "🟠", "afectado": "🔴"}
    resumen = f"**{asunto}**\n\n- Infraestructura: {punto[s_niv]} {s_val}\n- Colaboradores: {punto[c_niv]} {c_val}"
    if d.get("colab_casos"):
        resumen += "\n\nDetalle de colaboradores:" + "".join(f"\n- {c}" for c in d["colab_casos"])
    resumen += f"\n\nConfirmado por: {' y '.join(c for c in d['conf'] if c)}"

    def b64(ruta):
        if not ruta or not os.path.exists(ruta):
            return ""
        with open(ruta, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")

    datos = {
        "numero": d["numero"], "fecha": d["fecha"], "corte": corte, "asunto": asunto,
        "cuerpo_html": cuerpo, "resumen_aprobacion": resumen,
        "sedes": s_val, "sedes_nivel": s_niv, "colaboradores": c_val, "colaboradores_nivel": c_niv,
        "confirmado_por": " y ".join(c for c in d["conf"] if c),
        "archivo_png": os.path.basename(png_path) if png_path else "",
        "archivo_pptx": os.path.basename(pptx_path),
        "ruta_pptx_onedrive": ruta_en_onedrive(pptx_path),
        "ruta_png_onedrive": ruta_en_onedrive(png_path) if png_path else "",
        "png_base64": b64(png_path), "pptx_base64": b64(pptx_path),
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    # Nombre único: si regeneras el mismo reporte, el flujo se vuelve a disparar
    nombre = os.path.splitext(os.path.basename(pptx_path))[0] + f"_{datetime.now():%Y%m%d-%H%M%S}.json"
    destino = os.path.join(carpeta, nombre)
    tmp = destino + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False)
    os.replace(tmp, destino)          # aparece completo de una vez (OneDrive no sube un archivo a medias)

    # Vista previa del correo (fuera de _para_envio para no disparar el flujo)
    carpeta_prev = os.path.join(carpeta_salida, "_vista_previa_correo")
    os.makedirs(carpeta_prev, exist_ok=True)
    previa = os.path.join(carpeta_prev, os.path.splitext(os.path.basename(pptx_path))[0] + "_correo.html")
    with open(previa, "w", encoding="utf-8") as f:
        f.write('<!doctype html><html><head><meta charset="utf-8"><title>' + html.escape(asunto) + '</title></head>'
                '<body style="margin:24px;background:#fff">'
                '<div style="font:600 18px Segoe UI,Arial;margin-bottom:4px">Asunto: ' + html.escape(asunto) + '</div>'
                '<div style="font:12px Segoe UI,Arial;color:#555;margin-bottom:16px">Adjunto: '
                + html.escape(os.path.basename(pptx_path)) + ' · Desde: buzón de Continuidad (según el flujo)</div>'
                + cuerpo + '</body></html>')
    print("\n" + "-" * 70)
    if datos["ruta_pptx_onedrive"]:
        print("[OK] ONEDRIVE / POWER AUTOMATE: PPT, PNG y JSON guardados en OneDrive")
        print(f"  PPT en OneDrive: {datos['ruta_pptx_onedrive']}")
        print(f"  JSON que dispara el flujo: {destino}")
        print("  Espera a que OneDrive termine de sincronizar (check verde) y te llegará la aprobación.")
    else:
        print("[X] ONEDRIVE / POWER AUTOMATE: LOS ARCHIVOS NO ESTÁN EN ONEDRIVE")
        print(f"  Se guardaron en: {carpeta_salida}")
        print("  El flujo NO se va a disparar. En configuracion.txt escribe la carpeta de OneDrive, por ejemplo:")
        print("  CARPETA_SALIDA: C:\\Users\\<usuario>\\OneDrive - Pacifico Seguros\\Reporte FEN\\Reportes generados")
    print("-" * 70)
    print(f"  Vista previa del correo: {previa}")
    print("  En unos minutos te llegará la aprobación (Teams/Outlook). Al aprobarla, se envía el correo "
          "desde el buzón de Continuidad.")
    if platform.system() == "Windows":
        try:
            os.startfile(previa)  # noqa: S606  abre la vista previa en el navegador
        except Exception:  # noqa: BLE001
            pass


# ================================================================ 7. Principal
def instalar_fuentes():
    """En Windows, instala Nunito para el usuario actual (sin permisos de administrador) si aún no está."""
    if platform.system() != "Windows":
        return
    carpeta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recursos", "fuentes")
    if not os.path.isdir(carpeta):
        return
    import ctypes
    import winreg
    destino = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
    os.makedirs(destino, exist_ok=True)
    nuevas = 0
    for archivo, nombre in (("Nunito-Regular.ttf", "Nunito (TrueType)"), ("Nunito-Bold.ttf", "Nunito Bold (TrueType)"),
                            ("Nunito-ExtraBold.ttf", "Nunito ExtraBold (TrueType)")):
        origen, final = os.path.join(carpeta, archivo), os.path.join(destino, archivo)
        if not os.path.exists(origen) or os.path.exists(final):
            continue
        shutil.copy2(origen, final)
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Fonts",
                            0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, nombre, 0, winreg.REG_SZ, final)
        ctypes.windll.gdi32.AddFontResourceW(final)
        nuevas += 1
    if nuevas:
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001D, 0, 0, 0, 1000, None)
        print("Fuente Nunito instalada. Si PowerPoint estaba abierto, ciérralo y vuelve a ejecutar el script.")


def siguiente_numero(carpeta):
    """Busca el último Reporte_FEN_###_... en la carpeta de salida y devuelve el siguiente."""
    nums = []
    if os.path.isdir(carpeta):
        for f in os.listdir(carpeta):
            m = re.match(r"Reporte_FEN_(\d+)_", f)
            if m and "BORRADOR" not in f:
                nums.append(int(m.group(1)))
    return f"{max(nums) + 1:03d}" if nums else ""


def main():
    ap = argparse.ArgumentParser(description="Genera el Reporte FEN de Pacífico en PPTX y PNG.")
    carpeta_script = os.path.dirname(os.path.abspath(__file__))
    # Si no se indica archivo (p. ej. con el botón ▶ de VS Code), usa mensaje_whatsapp.txt junto al script
    ap.add_argument("datos", nargs="?", default=os.path.join(carpeta_script, "mensaje_whatsapp.txt"),
                    help="archivo de datos (por defecto: mensaje_whatsapp.txt junto al script)")
    ap.add_argument("--borrador", action="store_true", help="generar aunque falten las 2 confirmaciones")
    ap.add_argument("--sin-correo", action="store_true", help="solo generar el PPT y el PNG, sin el paso de correo")
    config = leer_configuracion(carpeta_script)
    salida_detectada = ""
    if config["modo"] == "automate" and not config["salida"]:
        salida_detectada = buscar_salida_onedrive()
        config["salida"] = salida_detectada
    ap.add_argument("--salida", default=config["salida"] or os.path.join(carpeta_script, "Reportes generados"),
                    help="carpeta donde guardar los PPT y PNG (por defecto: configuracion.txt o 'Reportes generados')")
    a = ap.parse_args()

    print(f"Reporte FEN · versión {VERSION}")
    print(f"  Script: {os.path.abspath(__file__)}")
    if not os.path.exists(a.datos):
        sys.exit(f"No encuentro el archivo de datos: {os.path.abspath(a.datos)}\n"
                 "Ponle ese nombre a tu .txt o indica cuál usar al final del comando.")
    print(f"  Datos:  {os.path.abspath(a.datos)}")
    print(f"  Salida: {os.path.abspath(a.salida)}   (correo: {config['modo']})")
    if config["modo"] == "automate":
        if salida_detectada:
            print("  [!] CARPETA_SALIDA está vacía en configuracion.txt; uso la carpeta de OneDrive del flujo:")
            print(f"      {salida_detectada}")
            print("      Escríbela en configuracion.txt para no depender de la búsqueda.")
        if ruta_en_onedrive(os.path.join(a.salida, "x")):
            print("  [OK] La salida está en OneDrive: el flujo de Power Automate se disparará.")
        else:
            print("  [X] La salida NO está en OneDrive: el flujo de Power Automate no se disparará.")
            print("      Revisa CARPETA_SALIDA en configuracion.txt (debe ser tu carpeta de OneDrive).")
    d = leer_datos(a.datos)
    print(f"Leído: {len(d['lluvias'])} lluvias · {len(d['vias'])} vías · {len(d['rios'])} ríos/quebradas · "
          f"{len(d['p1']) + len(d['p2'])} pronósticos"
          + (f" · {len(d['medidas'])} medidas" if d["medidas"] else "")
          + (f" · {len(d['recomendaciones'])} recomendaciones" if d["recomendaciones"] else ""))
    if d["colab_casos"]:
        print(f"Colaboradores afectados: {d['colab']}  ({len(d['colab_casos'])} caso(s))")
    if d["agencias"]:
        af = [n for n, e in d["agencias"] if es_afectado(e)]
        print(f"Agencias/sedes Pacífico: {len(d['agencias'])} revisadas, "
              + (f"{len(af)} con afectación ({', '.join(af)})." if af else "ninguna con afectación."))
        if d["ag_fecha"] and d["fecha"] and \
                re.sub(r"\D", "", d["ag_fecha"])[:4] != re.sub(r"\D", "", d["fecha"])[:4]:
            print(f"  Aviso: el mensaje de agencias es del {d['ag_fecha']} y el reporte es del {d['fecha']}. "
                  "Pega el mensaje de agencias de hoy.")
    for campo, nombre in (("lluvias", "LLUVIAS"), ("vias", "VÍAS RESTRINGIDAS"), ("rios", "RÍOS Y QUEBRADAS")):
        if not d[campo]:
            print(f"  Aviso: la sección {nombre} quedó vacía. Revisa que el mensaje la traiga.")
    if not d["numero"]:
        d["numero"] = siguiente_numero(a.salida)
        if d["numero"]:
            print(f"Número de reporte no indicado: se usa el siguiente al último generado ({d['numero']}).")
    print("-" * 70)
    if config["radar"] not in ("si", "sí", "yes", "1"):
        print("[--] OBSERVATORIO PACÍFICO: DESACTIVADO (configuracion.txt dice RADAR: no).")
        print("  El reporte sale sin el mapa. Para activarlo escribe  RADAR: si")
    else:
        datos_radar, origen = cargar_radar([carpeta_script, a.salida])
        if not datos_radar:
            print("[X] OBSERVATORIO PACÍFICO: SIN ACCESO")
            print(f"  {origen}.")
            print(f"  -> Guarda el snapshot.json de hoy en: {carpeta_script}")
            print("  El reporte sale sin el mapa (con la fila de agencias de siempre).")
        else:
            d["radar"] = resumir_radar(datos_radar, d["fecha"], [n for n, _ in d["agencias"]] or None)
            d["radar_estilo"] = ("mapa" if config["radar_estilo"].startswith("mapa") else
                                 "sedes" if config["radar_estilo"].startswith("sede") else "avisos")
            rz = d["radar"]
            descargado = origen.startswith("descargado")
            print("[OK] OBSERVATORIO PACÍFICO: " + ("datos descargados de la página" if descargado
                  else "la página no respondió; se usan los datos del snapshot.json guardado"))
            capt = fin_local(datos_radar.get("at"))
            if capt:
                print(f"  Datos del {capt[8:]}/{capt[5:7]}/{capt[:4]}"
                      + ("" if descargado else f" (archivo: {os.path.join(carpeta_script, 'snapshot.json')})"))
            print(f"  Lluvias en {len(rz['lluvias'])} dptos · quebradas en {len(rz['quebradas'])} dptos · "
                  f"{sum(1 for s_ in rz['sedes'] if s_['nivel'])} sede(s) en zona con aviso")
            if capt and capt != rz["dia"]:
                print(f"  [!] AVISO: el snapshot es del {capt[8:]}/{capt[5:7]} y el reporte es del "
                      f"{rz['dia'][8:]}/{rz['dia'][5:7]}. Guarda el snapshot.json de hoy y vuelve a ejecutar.")
            if not rz["lluvias"] and not rz["quebradas"]:
                print("  [!] AVISO: no hay avisos de lluvias ni quebradas para la fecha del reporte; el mapa sale en gris.")
            if d["radar_estilo"] == "mapa" and not ruta_geojson():
                print("  [!] AVISO: falta recursos\\peru_departamentos.geojson; el mapa no se puede dibujar.")
    print("-" * 70)
    faltan = [c for c in ("numero", "fecha") if not d[c]]
    if faltan:
        sys.exit("Falta completar en el archivo de datos: " + ", ".join(f.upper() for f in faltan)
                 + ". Agrega por ejemplo la línea  NUMERO: 013")

    confirmados = [c for c in d["conf"] if c]
    if len(confirmados) < 2 and not a.borrador:
        sys.exit(f"Solo hay {len(confirmados)} de 2 confirmaciones en [CONFIRMACIONES]. "
                 "Completa los nombres cuando confirmen por WhatsApp, o usa --borrador para una vista previa.")

    # Si la carpeta viene de configuracion.txt, su carpeta "madre" debe existir (evita crear rutas inventadas)
    if config["salida"] and os.path.abspath(a.salida) == os.path.abspath(config["salida"]):
        madre = os.path.dirname(os.path.abspath(a.salida))
        if "..." in a.salida or not os.path.isdir(madre):
            sys.exit(f"La CARPETA_SALIDA de configuracion.txt no existe:\n  {a.salida}\n"
                     "Copia la ruta real desde la barra de direcciones del Explorador de archivos "
                     "(sin '...') o deja la línea vacía para guardar en 'Reportes generados'.")
    os.makedirs(a.salida, exist_ok=True)
    base = f"Reporte_FEN_{d['numero']}_Pacifico_{d['fecha'].replace('/', '-')}"
    if len(confirmados) < 2:
        base += "_BORRADOR"
    pptx_path = os.path.join(a.salida, base + ".pptx")
    png_path = os.path.join(a.salida, base + ".png")

    try:
        instalar_fuentes()
    except Exception as e:  # noqa: BLE001
        print(f"  (No se pudo instalar la fuente Nunito: {e}. Instálala a mano desde recursos/fuentes.)")
    try:
        construir(d, pptx_path, borrador=len(confirmados) < 2)
    except PermissionError:
        sys.exit(f"No se pudo guardar {os.path.basename(pptx_path)}: está abierto en PowerPoint. "
                 "Ciérralo y vuelve a ejecutar.")
    print(f"PPT creado: {pptx_path}")
    print("  (No abras el PPT ni el PNG hasta que el script termine.)")
    motor = exportar_png(pptx_path, png_path)
    if motor:
        print(f"PNG creado ({motor}): {png_path}")
    else:
        print("No se generó el PNG: instala PowerPoint + pywin32 (Windows) o LibreOffice. "
              "Mientras tanto, en PowerPoint: Archivo > Exportar > PNG.")

    print(f"\nArchivos guardados en: {os.path.abspath(a.salida)}")
    if (a.sin_correo or config["modo"] != "outlook") and platform.system() == "Windows":
        try:
            os.startfile(os.path.abspath(a.salida))  # noqa: S606  abre la carpeta en el Explorador
        except Exception:  # noqa: BLE001
            pass

    # Paso de validación y envío por correo
    if a.sin_correo:
        return
    if len(confirmados) < 2:
        print("\nEs un BORRADOR (faltan confirmaciones): no se ofrece el envío por correo.")
        return
    if config["modo"] == "automate":
        if not motor:
            print("\nNo se preparó el envío: falta el PNG.")
            return
        problema = revisar_imagen(png_path)
        if problema:
            print(f"\nNo se preparó el envío: la imagen del reporte salió incompleta ({problema}).")
            print("Suele pasar si se abre el PPT o el PNG mientras el script corre. Cierra PowerPoint y "
                  "vuelve a ejecutar sin abrir nada hasta que termine.")
            return
        preparar_envio_automate(d, pptx_path, png_path, a.salida)
    elif config["modo"] == "outlook":
        paso_correo(d, pptx_path, png_path if motor else None, carpeta_script, a.salida)
    else:
        print("\nCorreo desactivado (MODO_CORREO: ninguno).")


if __name__ == "__main__":
    main()
