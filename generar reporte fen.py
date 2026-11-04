"""
Genera el Reporte FEN de Pacífico Seguros en PowerPoint (.pptx) y en imagen (.png).

Uso:
    python generar_reporte_fen.py datos_fen.txt
    python generar_reporte_fen.py datos_fen.txt --borrador   (genera aunque falten confirmaciones)

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
    "AGENCIAS": "agencias",
    "AGENCIAS PACIFICO": "agencias",
    "AGENCIAS PACÍFICO": "agencias",
}


def _vacio():
    return {"numero": "", "fecha": "", "horario": "", "lluvias": [], "vias": [], "rios": [],
            "p1_titulo": "", "p1": [], "p2_titulo": "", "p2": [], "nota": "",
            "conf": [], "sedes": "", "colab": "", "pie": [], "medidas": [],
            "agencias": [], "ag_fecha": "", "ag_corte": "", "ag_resumen": ""}


SIN_AFECTACION = re.compile(r"^(sin\s+(afectaci|impacto|novedad|da[ñn]o|incidencia)|operativ|normal)", re.I)


def es_afectado(valor):
    """True si el texto describe una afectación (no empieza con 'Sin afectación', 'Sin impacto', etc.)."""
    return bool(valor.strip()) and not SIN_AFECTACION.match(valor.strip())


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
        else:
            d[actual].append(linea)
    return d


def _limpiar(linea):
    """Quita emojis, viñetas y caracteres invisibles del inicio de una línea de WhatsApp."""
    linea = linea.replace("​", "").replace("️", "").strip()
    linea = re.sub(r"^[^\wÁÉÍÓÚÑáéíóúñ¿¡(]+", "", linea)
    return linea.replace("*", "").replace("_", " ").strip()


def _etiqueta_bonita(linea):
    """'SIERRA: Se prevén...' -> 'Sierra: Se prevén...'"""
    etq, sep, resto = linea.partition(":")
    if sep and etq.isupper() and len(etq) <= 40:
        return etq.strip().capitalize() + ": " + resto.strip()
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

    if hay(r"agencias?") and hay(r"pac[ií]fico"):
        return "agencias"           # mensaje de agencias de Pacífico (alimenta "Sedes")
    if hay(r"pron[oó]stico"):
        return "pronostico"
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
        if actual == "agencias" and not (not _es_vineta(raw) and _tipo_seccion(linea) not in (None, "agencias")
                                         and (linea.endswith(":") or not sep)):
            _linea_agencia(d, linea)
            continue

        # Datos de cabecera
        if kl in ("numero", "número", "n°", "nro"):
            d["numero"] = valor.strip()
            continue
        if kl == "fecha":
            d["fecha"] = valor.strip()
            continue
        if kl in ("corte", "horario", "hora"):
            d["horario"] = linea.strip() if kl == "corte" else valor.strip()
            continue
        if "reporte fenomenol" in low:
            continue
        if low.startswith("se mantiene el monitoreo"):
            d["nota"] = linea
            actual = None
            continue

        # ¿Es un título de sección? (no lleva viñeta y no tiene texto después de los dos puntos)
        es_titulo = (not _es_vineta(raw)) and len(linea) <= 90 and (linea.endswith(":") or not sep)
        if es_titulo:
            tipo = _tipo_seccion(linea)
            if tipo == "pronostico":
                actual = "p1" if not d["p1"] else "p2"
                titulo = re.sub(r"^pron[oó]stico\s*[–\-:]*\s*", "", linea, flags=re.I).rstrip(":").strip()
                d[actual + "_titulo"] = titulo.upper() if titulo else "PRONÓSTICO"
            elif tipo:
                actual = tipo
            else:
                actual = "ignorar"
                print(f"  Aviso: no reconocí la sección «{linea.rstrip(':')}»; sus líneas no se incluyen.")
            continue

        if actual and actual != "ignorar":
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
        return resolver_sedes(_leer_estructurado(lineas, d))
    # Mensaje de WhatsApp arriba + datos de Pacífico entre corchetes abajo
    corte = next((i for i, s in enumerate(secciones) if s), len(lineas))
    _leer_whatsapp(lineas[:corte], d)
    return resolver_sedes(_leer_estructurado(lineas[corte:], d))


def resolver_sedes(d):
    """Si se pegó el mensaje de agencias de Pacífico, el estado de 'Sedes' sale de ahí.
    Lo que escribas a mano en 'Sedes:' manda solo si describe una afectación."""
    if not d["agencias"] or es_afectado(d["sedes"]):
        return d
    afectadas = [(n, e) for n, e in d["agencias"] if es_afectado(e)]
    total = len(d["agencias"])
    if not afectadas:
        d["sedes"] = f"Sin afectación en las {total} agencias de provincia." if total > 1 else \
            "Sin afectación en agencias de provincia."
    elif len(afectadas) == 1:
        d["sedes"] = f"{afectadas[0][0]}: {afectadas[0][1]}"
    else:
        nombres = [n for n, _ in afectadas]
        d["sedes"] = f"Afectación en {len(afectadas)} de {total} agencias: " + ", ".join(nombres[:-1]) + \
            " y " + nombres[-1] + "."
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
            r.font.bold = b
            r.font.color.rgb = c
            r.font.name = FONT
    return tb


def vinetas(slide, x, y, w, h, items, size, nombre):
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
        p.space_after = Pt(3)
        p.line_spacing = Pt(round(size * 1.2, 1))   # interlineado fijo: igual en PowerPoint y LibreOffice
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


def alto_lista(items, size, ancho_in):
    disp = ancho_in * 72 - 0.13 * 72
    total = 0.0
    for linea in items:
        etq, txt = separar(linea)
        n = contar_lineas(etq, txt, size, disp)
        total += n * size * 1.2 + 3
    return total / 72.0


def icono(slide, buf, x, y, lado):
    slide.shapes.add_picture(buf, Inches(x), Inches(y), Inches(lado), Inches(lado))


def linea_h(slide, x1, x2, y):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y), Inches(x2), Inches(y))
    c.line.color.rgb = RGBColor(0xDD, 0xE6, 0xF0)
    c.line.width = Pt(0.75)
    c.shadow.inherit = False
    sin_estilo(c)


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

G_BLUE, G_PURPLE, G_PINK = (7, 152, 214), (123, 91, 168), (217, 47, 131)


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
    fondo, tinta = (BAD_BG, BAD_FG) if afectado else (OK_BG, OK_FG)
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
    texto(slide, 0.6, 0.24, 7.5, 0.62, "REPORTE FEN " + d["numero"], 32, WHITE, True, nombre="Titulo")
    texto(slide, 0.62, 0.86, 7.0, 0.3, "Reporte fenomenológico nacional", 14, WHITE, nombre="Subtitulo")
    if recurso("logo_continuidad.png"):
        imagen_ajustada(slide, recurso("logo_continuidad.png"), 7.55, 0.30, alto=0.62)
    fb = caja(slide, 10.05, 0.24, 2.78, 0.9, WHITE, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Caja fecha")
    fb.adjustments[0] = 0.18
    texto(slide, 10.3, 0.32, 2.4, 0.2, "FECHA", 9, P_GREY, True)
    texto(slide, 10.3, 0.52, 2.4, 0.32, d["fecha"], 17, P_BLUE_DK, True)
    texto(slide, 10.3, 0.86, 2.4, 0.2, d["horario"], 9, P_GREY, True)
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
    sec_medidas = ("MEDIDAS", _megafono, G_PURPLE, d["medidas"]) if d["medidas"] else None
    blq_medidas = ("MEDIDAS", P_PURPLE, d["medidas"]) if d["medidas"] else None
    hay_agencias = bool(d["agencias"])

    TOP, BOTTOM = 1.42, 6.86
    HEAD = 0.5
    y0 = TOP + HEAD + 0.14
    EXTRA_AG = 0.27 if hay_agencias else 0.0      # fila de agencias dentro del panel de impacto
    Y1_NORMAL, Y1_COMPACTO = 5.42, 6.08 - EXTRA_AG
    ry0, ry1 = TOP + HEAD + 0.16, 6.02
    MIN_FS = 8.0

    def alto_tit(t, ancho, s=13):
        return contar_lineas("", t.upper(), s, ancho * 72 * 0.995) * s * 1.2 / 72 + 0.08

    def fs_izq(LW, y1):
        lw = LW - 0.86 - 0.28
        need = lambda f: sum(0.34 + alto_lista(it, f, lw) + 0.12 for *_, it in secs)
        fs = 12.0
        while fs > MIN_FS and need(fs) > y1 - y0:
            fs -= 0.5
        return fs, need(fs) <= y1 - y0

    def fs_der(RW):
        rw = RW - 0.56
        need = lambda f: sum(alto_tit(t, rw) + alto_lista(it, f, rw) + 0.22 for t, _, it in bloques)
        fs = 12.0
        while fs > MIN_FS and need(fs) > ry1 - ry0:
            fs -= 0.5
        return fs, need(fs) <= ry1 - ry0

    LX, GAP, ANCHO_TOTAL = 0.5, 0.22, 12.33

    def elegir(y1):
        mejor = None
        for LW in (7.75, 8.15, 8.55, 8.95, 7.35, 6.95, 6.55):
            RW = ANCHO_TOTAL - LW - GAP
            (a, ok_a), (b, ok_b) = fs_izq(LW, y1), fs_der(RW)
            clave = (ok_a and ok_b, min(a, b), -abs(LW - 7.75))
            if mejor is None or clave > mejor[0]:
                mejor = (clave, LW, RW, a, b)
        return mejor

    compacto = True                     # un solo diseño: impacto en una fila
    # "Medidas" (cuando viene) va a la columna donde deje la letra más grande
    opciones = [(secs_base, bloques_base)]
    if sec_medidas:
        opciones = [(secs_base + [sec_medidas], bloques_base), (secs_base, bloques_base + [blq_medidas])]
    mejor = None
    for secs, bloques in opciones:
        m = elegir(Y1_COMPACTO)
        if mejor is None or m[0] > mejor[0][0]:
            mejor = (m, secs, bloques)
    (m, secs, bloques) = mejor
    (cabe, _, _), LW, RW, fs, fs2 = m
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
        texto(slide, x + 0.3, TOP, w - 0.5, HEAD, titulo, 16, color, True, anchor=MSO_ANCHOR.MIDDLE)

    # ---- Columna izquierda: estatus del día
    lista_x, lista_w = LX + 0.86, LW - 0.86 - 0.28
    y = y0
    for i, (titulo, ic, col, items) in enumerate(secs):
        icono(slide, icono_circulo(ic, col), LX + 0.3, y - 0.02, 0.4)
        texto(slide, lista_x, y + 0.02, lista_w, 0.26, titulo, 12.5, P_INK, True)
        h = alto_lista(items, fs, lista_w)
        vinetas(slide, lista_x, y + 0.34, lista_w, h + 0.02, items, fs, "Lista " + titulo.lower())
        y += 0.34 + h + 0.12
        if i < len(secs) - 1:
            linea_h(slide, lista_x, LX + LW - 0.28, y - 0.06)

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
        px, py, pw, ph = LX + 0.22, 6.14 - EXTRA_AG, LW - 0.44, 0.6 + EXTRA_AG
        panel = caja(slide, px, py, pw, ph, P_TINT, None, MSO_SHAPE.ROUNDED_RECTANGLE, "Panel impacto")
        panel.adjustments[0] = 0.25 if not hay_agencias else 0.18
        texto(slide, px + 0.2, py, 2.1, 0.6, [[("IMPACTO EN", True, P_BLUE_DK)], [("PACÍFICO SEGUROS", True, P_BLUE_DK)]],
              11, anchor=MSO_ANCHOR.MIDDLE)
        resto = pw - 2.35
        col_w = resto / 2
        for j, (etq, ic, col, valor) in enumerate(impactos):
            cx = px + 2.25 + j * col_w
            icono(slide, icono_circulo(ic, col), cx, py + 0.11, 0.38)
            v = valor or "Sin afectación reportada."
            af = es_afectado(v)
            ancho_chip = col_w - 0.56
            fs_chip = 9.5
            while fs_chip > 7 and contar_lineas(etq, v, fs_chip, (ancho_chip - 0.3) * 72) * fs_chip * 1.2 / 72 > 0.36:
                fs_chip -= 0.5
            chip(slide, cx + 0.46, py + 0.1, ancho_chip, 0.4, etq, v, af, fs_chip)
        if hay_agencias:
            # Fila con cada agencia revisada: punto verde (sin afectación) o rojo (con afectación)
            f_ag = re.sub(r"^(\d{1,2}/\d{1,2})/\d{2,4}$", r"\1", d["ag_fecha"].strip())
            c_ag = re.sub(r"\s*(horas|hrs?|h)\.?$", "", d["ag_corte"].strip(), flags=re.I)
            extra = ", ".join(x for x in (f_ag, c_ag) if x)
            partes = [("Agencias provincia" + (f" ({extra})" if extra else "") + ":  ", True, P_BLUE_DK)]
            for n, e in d["agencias"]:
                af = es_afectado(e)
                partes += [("●\u00a0", False, BAD_FG if af else OK_FG), (n + "   ", af, BAD_FG if af else P_INK)]
            ancho_ag = (pw - 0.4) * 72 * 0.97
            fs_ag = 9.5
            while fs_ag > 6.5 and sum(ancho_pt(t, fs_ag, b) for t, b, _ in partes) > ancho_ag:
                fs_ag -= 0.5
            texto(slide, px + 0.2, py + 0.58, pw - 0.4, 0.22, [partes], fs_ag, anchor=MSO_ANCHOR.MIDDLE,
                  nombre="Agencias provincia")

    # ---- Columna derecha: pronóstico
    rl_x, rl_w = RX + 0.28, RW - 0.56
    y = ry0
    for i, (t, col, items) in enumerate(bloques):
        ht = alto_tit(t, rl_w)
        texto(slide, rl_x, y, rl_w, ht, t.upper(), 13, col, True)
        h = alto_lista(items, fs2, rl_w)
        vinetas(slide, rl_x, y + ht, rl_w, h + 0.02, items, fs2, t.capitalize() if t == "MEDIDAS" else f"Pronostico {i + 1}")
        y += ht + h + 0.22
        if i < len(bloques) - 1:
            linea_h(slide, rl_x, RX + RW - 0.28, y - 0.11)
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

    # Notas: trazabilidad de confirmaciones
    conf = [c for c in d["conf"] if c]
    slide.notes_slide.notes_text_frame.text = (
        ("Confirmado sin afectados por: " + " y ".join(conf)) if conf else "Sin confirmaciones registradas."
    ) + (("\nAgencias de provincia" + (f" (corte {d['ag_corte']}, {d['ag_fecha']})" if d["ag_corte"] else "") + ":\n"
          + "\n".join(f"  - {n}: {e}" for n, e in d["agencias"])) if d["agencias"] else "") \
      + f"\nGenerado el {datetime.now():%d/%m/%Y %H:%M}."

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
            pres.Slides(1).Export(png_tmp, "PNG", 1920, 1080)
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
            subprocess.run(["pdftoppm", "-png", "-r", "144", "-singlefile", pdf, base], check=True)
            shutil.move(base + ".png", png_path)
        else:
            subprocess.run(cmd + ["--headless", "--convert-to", "png", "--outdir", tmp, pptx_path],
                           check=True, capture_output=True, timeout=180)
            shutil.move(os.path.join(tmp, os.path.splitext(os.path.basename(pptx_path))[0] + ".png"), png_path)
    return "LibreOffice"


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
    for archivo, nombre in (("Nunito-Regular.ttf", "Nunito (TrueType)"), ("Nunito-Bold.ttf", "Nunito Bold (TrueType)")):
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
    ap.add_argument("datos", help="archivo de datos (ver datos_fen.txt)")
    ap.add_argument("--borrador", action="store_true", help="generar aunque falten las 2 confirmaciones")
    carpeta_script = os.path.dirname(os.path.abspath(__file__))
    ap.add_argument("--salida", default=os.path.join(carpeta_script, "Reportes generados"),
                    help="carpeta donde guardar los PPT y PNG (por defecto: 'Reportes generados' junto al script)")
    a = ap.parse_args()

    d = leer_datos(a.datos)
    print(f"Leído: {len(d['lluvias'])} lluvias · {len(d['vias'])} vías · {len(d['rios'])} ríos/quebradas · "
          f"{len(d['p1']) + len(d['p2'])} pronósticos"
          + (f" · {len(d['medidas'])} medidas" if d["medidas"] else ""))
    if d["agencias"]:
        af = [n for n, e in d["agencias"] if es_afectado(e)]
        print(f"Agencias Pacífico: {len(d['agencias'])} revisadas, "
              + (f"{len(af)} con afectación ({', '.join(af)})." if af else "ninguna con afectación."))
    for campo, nombre in (("lluvias", "LLUVIAS"), ("vias", "VÍAS RESTRINGIDAS"), ("rios", "RÍOS Y QUEBRADAS")):
        if not d[campo]:
            print(f"  Aviso: la sección {nombre} quedó vacía. Revisa que el mensaje la traiga.")
    if not d["numero"]:
        d["numero"] = siguiente_numero(a.salida)
        if d["numero"]:
            print(f"Número de reporte no indicado: se usa el siguiente al último generado ({d['numero']}).")
    faltan = [c for c in ("numero", "fecha") if not d[c]]
    if faltan:
        sys.exit("Falta completar en el archivo de datos: " + ", ".join(f.upper() for f in faltan)
                 + ". Agrega por ejemplo la línea  NUMERO: 013")

    confirmados = [c for c in d["conf"] if c]
    if len(confirmados) < 2 and not a.borrador:
        sys.exit(f"Solo hay {len(confirmados)} de 2 confirmaciones en [CONFIRMACIONES]. "
                 "Completa los nombres cuando confirmen por WhatsApp, o usa --borrador para una vista previa.")

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
    motor = exportar_png(pptx_path, png_path)
    if motor:
        print(f"PNG creado ({motor}): {png_path}")
    else:
        print("No se generó el PNG: instala PowerPoint + pywin32 (Windows) o LibreOffice. "
              "Mientras tanto, en PowerPoint: Archivo > Exportar > PNG.")


if __name__ == "__main__":
    main()
