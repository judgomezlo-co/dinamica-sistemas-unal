"""
Constructor de documentos Word sobre la plantilla institucional
"Formato para presentacion de trabajos ASI - Word.docx".

Conserva de la plantilla: portada (logo UNAL, bloque institucional), estilos
(Arial 11, titulos de 20/16/14 pt), numeracion de capitulos, encabezados pares/impares
con numero de pagina, tamano carta y margenes (sup/inf 2.54 cm, izq 3.6 cm, der 2.54 cm),
y la seccion independiente de Bibliografia con su propio encabezado.

Las ecuaciones se insertan como objetos nativos de Word (OMML) convirtiendo
LaTeX -> MathML (latex2mathml) -> OMML (MML2OMML.XSL incluido con Microsoft Office).
"""

import copy
import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from latex2mathml.converter import convert as latex_a_mathml
from lxml import etree

ANCHO_TEXTO = 12240 - 2041 - 1440  # twips disponibles entre margenes (8759)
XSL_CANDIDATOS = [
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
]
NUM_VINETAS = 2      # numId de la plantilla: vinetas (circulo negro)
NUM_CAPITULOS = 3    # numId de la plantilla: 1. / 1.1 / 1.1.1
PH_DOLAR = "\uE000"  # marcador para "\$" literal en el texto


# ----------------------------------------------------------------------------
# Utilidades XML
# ----------------------------------------------------------------------------
def E(tag, attrs=None, *hijos):
    el = OxmlElement(tag)
    for k, v in (attrs or {}).items():
        el.set(qn(k), str(v))
    for h in hijos:
        if h is not None:
            el.append(h)
    return el


def rpr(size=22, bold=False, italic=False, color="000000", sup=False, font=None):
    r = E("w:rPr")
    if font:
        r.append(E("w:rFonts", {"w:ascii": font, "w:hAnsi": font, "w:cs": font, "w:eastAsia": font}))
    if bold:
        r.append(E("w:b"))
        r.append(E("w:bCs"))
    if italic:
        r.append(E("w:i"))
        r.append(E("w:iCs"))
    if color:
        r.append(E("w:color", {"w:val": color}))
    r.append(E("w:sz", {"w:val": size}))
    r.append(E("w:szCs", {"w:val": size}))
    if sup:
        r.append(E("w:vertAlign", {"w:val": "superscript"}))
    return r


def run(texto, **kw):
    r = E("w:r", None, rpr(**kw))
    partes = texto.split("\t")
    for i, parte in enumerate(partes):
        if i:
            r.append(E("w:tab"))
        if parte:
            t = E("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = parte
            r.append(t)
    return r


def ppr(jc="both", line=360, before=0, after=240, ind=None, keep_next=False, keep_lines=False,
        page_break=False, style=None, numpr=None, tabs=None, widow=True):
    p = E("w:pPr")
    if style:
        p.append(E("w:pStyle", {"w:val": style}))
    if keep_next:
        p.append(E("w:keepNext"))
    if keep_lines:
        p.append(E("w:keepLines"))
    if page_break:
        p.append(E("w:pageBreakBefore"))
    if not widow:
        p.append(E("w:widowControl", {"w:val": "0"}))
    if numpr is not None:
        ilvl, num = numpr
        p.append(E("w:numPr", None, E("w:ilvl", {"w:val": ilvl}), E("w:numId", {"w:val": num})))
    if tabs:
        t = E("w:tabs")
        for val, pos, leader in tabs:
            a = {"w:val": val, "w:pos": pos}
            if leader:
                a["w:leader"] = leader
            t.append(E("w:tab", a))
        p.append(t)
    sp = {"w:before": before, "w:after": after}
    if line:
        sp["w:line"] = line
        sp["w:lineRule"] = "auto"
    p.append(E("w:spacing", sp))
    if ind:
        p.append(E("w:ind", {f"w:{k}": v for k, v in ind.items()}))
    if jc:
        p.append(E("w:jc", {"w:val": jc}))
    return p


# ----------------------------------------------------------------------------
# Ecuaciones (LaTeX -> OMML)
# ----------------------------------------------------------------------------
class Matematica:
    def __init__(self):
        xsl = next((x for x in XSL_CANDIDATOS if Path(x).exists()), None)
        if xsl is None:
            raise FileNotFoundError("No se encontro MML2OMML.XSL (Microsoft Office).")
        self.xslt = etree.XSLT(etree.parse(xsl))

    @staticmethod
    def preparar(latex):
        latex = latex.replace(PH_DOLAR, r"\$")
        # MML2OMML descarta los espacios de LaTeX: se convierten en espacios Unicode explicitos
        latex = (latex.replace(r"\qquad", "\\text{\u2003\u2003}")
                      .replace(r"\quad", "\\text{\u2003}")
                      .replace(r"\;", "\\text{\u2004}")
                      .replace(r"\,", "\\text{\u2009}"))
        # subindices de 2 o mas letras en texto recto (P_{mut} -> P_{\text{mut}})
        latex = re.sub(r"_\{([A-Za-zÁÉÍÓÚáéíóúñ]{2,})\}", r"_{\\text{\1}}", latex)
        latex = re.sub(r"\^\{([A-Za-zÁÉÍÓÚáéíóúñ]{2,})\}", r"^{\\text{\1}}", latex)
        latex = latex.replace(r"\max", r"\mathrm{max}").replace(r"\min", r"\mathrm{min}")
        return latex

    @staticmethod
    def _unir_delimitadores(om):
        """MML2OMML parte el contenido de los parentesis en varios m:e usando el primer
        operador como 'separador' (p. ej. (1-Q) -> e1='1', sep='-', e2='Q'), lo que produce
        espacios y simbolos erroneos. Se reunen en un solo m:e con el separador explicito."""
        m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        for d in list(om.iter(m + "d")):
            es = d.findall(m + "e")
            if len(es) < 2:
                continue
            dpr = d.find(m + "dPr")
            sep = dpr.find(m + "sepChr") if dpr is not None else None
            sepval = sep.get(m + "val") if sep is not None else "|"
            nuevo = etree.Element(m + "e")
            for i, e in enumerate(es):
                if i and sepval:
                    r = etree.SubElement(nuevo, m + "r")
                    t = etree.SubElement(r, m + "t")
                    t.text = sepval
                for hijo in list(e):
                    nuevo.append(hijo)
                d.remove(e)
            d.append(nuevo)
            if sep is not None:
                dpr.remove(sep)

    def omml(self, latex, size=22):
        mml = latex_a_mathml(self.preparar(latex))
        om = self.xslt(etree.fromstring(mml.encode("utf-8"))).getroot()
        om = copy.deepcopy(om)
        self._unir_delimitadores(om)
        for mr in om.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}r"):
            w_rpr = E("w:rPr", None,
                      E("w:rFonts", {"w:ascii": "Cambria Math", "w:hAnsi": "Cambria Math"}),
                      E("w:color", {"w:val": "000000"}),
                      E("w:sz", {"w:val": size}), E("w:szCs", {"w:val": size}))
            m_rpr = mr.find("{http://schemas.openxmlformats.org/officeDocument/2006/math}rPr")
            if m_rpr is not None:
                m_rpr.addnext(w_rpr)
            else:
                mr.insert(0, w_rpr)
        return om


# ----------------------------------------------------------------------------
# Constructor del informe
# ----------------------------------------------------------------------------
class Informe:
    def __init__(self, plantilla):
        self.doc = Document(str(plantilla))
        self.body = self.doc.element.body
        self.sect_final = self.body.find(qn("w:sectPr"))
        self.mat = Matematica()
        self.refs = {}
        self.n_fig = 0
        self.n_tab = 0
        self.sect_seccion1 = None

    # ------------------------------------------------------------ plantilla
    def preparar_plantilla(self, titulo, autores, linea_curso, anio, titulo_corto):
        hijos = list(self.body.iterchildren())
        # sectPr de la seccion 1 (va en el parrafo de "Conclusiones" de la plantilla)
        for h in hijos:
            s = h.find(qn("w:pPr") + "/" + qn("w:sectPr")) if h.tag == qn("w:p") else None
            if s is not None:
                self.sect_seccion1 = copy.deepcopy(s)
        # se eliminan todos los elementos posteriores a la portada (indices >= 24)
        for h in hijos[24:]:
            if h is not self.sect_final:
                self.body.remove(h)
        portada = list(self.body.iterchildren())[:24]

        def fijar_texto(p, lineas, **kw):
            for r in p.findall(qn("w:r")):
                p.remove(r)
            ppr_el = p.find(qn("w:pPr"))
            for i, linea in enumerate(lineas):
                r = run(linea, **kw)
                if i:
                    r.insert(1, E("w:br"))
                p.append(r)
            return ppr_el

        fijar_texto(portada[3], [titulo], size=48, bold=True)
        fijar_texto(portada[10], autores, size=32, bold=True)
        p_curso = portada[12]
        for h in list(p_curso):
            p_curso.remove(h)
        p_curso.append(ppr(jc="center", line=240, after=0))
        p_curso.append(run(linea_curso, size=24))
        fijar_texto(portada[23], [anio])
        # se retiran parrafos vacios para compensar las lineas extra de titulo y autores
        for idx in (4, 5, 6, 15, 16):
            self.body.remove(portada[idx])

        # encabezados: titulo corto del trabajo
        for rel in self.doc.part.rels.values():
            if rel.reltype.endswith("/header"):
                for t in rel.target_part.element.iter(qn("w:t")):
                    if t.text and t.text.strip() in ("Título del trabajo",
                                                      "Título de la tesis o trabajo de investigación"):
                        t.text = titulo_corto
        # relaciones huerfanas de la plantilla (imagen de ejemplo e hipervinculos)
        for rid, rel in list(self.doc.part.rels.items()):
            if rel.reltype.endswith("/hyperlink") or (rel.reltype.endswith("/image") and rid == "rId7"):
                del self.doc.part.rels[rid]
        self._estilos_toc()

    def _estilos_toc(self):
        estilos = self.doc.styles.element
        for nivel, ind, bold, before in ((1, 0, True, 120), (2, 440, False, 40)):
            st = E("w:style", {"w:type": "paragraph", "w:styleId": f"TOC{nivel}"},
                   E("w:name", {"w:val": f"toc {nivel}"}),
                   E("w:basedOn", {"w:val": "Normal"}),
                   E("w:next", {"w:val": "Normal"}),
                   E("w:uiPriority", {"w:val": 39}),
                   E("w:unhideWhenUsed"))
            pp = E("w:pPr", None,
                   E("w:tabs", None, E("w:tab", {"w:val": "right", "w:leader": "dot", "w:pos": ANCHO_TEXTO})),
                   E("w:spacing", {"w:before": before, "w:after": 40, "w:line": 276, "w:lineRule": "auto"}),
                   E("w:ind", {"w:left": ind}))
            st.append(pp)
            rp = E("w:rPr")
            if bold:
                rp.append(E("w:b"))
                rp.append(E("w:bCs"))
            st.append(rp)
            estilos.append(st)

    # ------------------------------------------------------------ insercion
    def _agregar(self, el):
        self.sect_final.addprevious(el)
        return el

    def registrar(self, tipo, etiqueta):
        if tipo == "fig":
            self.n_fig += 1
            self.refs[f"fig:{etiqueta}"] = self.n_fig
            return self.n_fig
        self.n_tab += 1
        self.refs[f"tab:{etiqueta}"] = self.n_tab
        return self.n_tab

    def predefinir(self, figuras, tablas):
        """Numera de antemano figuras y tablas (en orden de aparicion) para poder
        citarlas en el texto antes de insertarlas."""
        for f in figuras:
            self.registrar("fig", f)
        for t in tablas:
            self.registrar("tab", t)
        self._orden = {"fig": 0, "tab": 0}

    def _verificar_orden(self, tipo, n):
        esperado = self._orden[tipo] + 1
        if n != esperado:
            raise ValueError(f"{tipo} {n} insertada fuera de orden (se esperaba {esperado}).")
        self._orden[tipo] = n

    def resolver(self, texto):
        def rep(m):
            clave = m.group(1)
            if clave not in self.refs:
                raise KeyError(f"Referencia sin resolver: {clave}")
            return str(self.refs[clave])
        return re.sub(r"\{((?:fig|tab):[\w-]+)\}", rep, texto)

    @staticmethod
    def _espacios(tok):
        tok = re.sub(r"(?<=\d) (?=\d{3}\b)", " ", tok)       # 14 964
        tok = re.sub(r"(?<=\d) (?=%)", "\u00a0", tok)               # 6.8 %
        tok = re.sub(r"\b([Tt]ablas?|[Ff]iguras?|[Hh]oja|[Ff]ase|[Rr]egistro) (?=\d)", "\\1\u00a0", tok)
        return tok

    NEGRITA = re.compile(r"\*\*(.+?)\*\*")
    SEGMENTO = re.compile(r"(\$[^$]+\$|\*[^*\s][^*]*\*)")

    def inline(self, texto, size=22, bold=False, italic=False):
        """Convierte marcado ligero (**negrita**, *cursiva*, $latex$) en runs y OMML.
        La negrita puede contener cursivas y formulas."""
        texto = self.resolver(texto.replace(r"\$", PH_DOLAR))
        elementos, pos = [], 0
        for m in self.NEGRITA.finditer(texto):
            elementos += self._segmento(texto[pos:m.start()], size, bold, italic)
            elementos += self._segmento(m.group(1), size, True, italic)
            pos = m.end()
        elementos += self._segmento(texto[pos:], size, bold, italic)
        return elementos

    def _segmento(self, texto, size, bold, italic):
        elementos = []
        for tok in self.SEGMENTO.split(texto):
            if not tok:
                continue
            if tok.startswith("$") and tok.endswith("$") and len(tok) > 1:
                elementos.append(self.mat.omml(tok[1:-1], size=size))
                continue
            tok = self._espacios(tok).replace(PH_DOLAR, "$")
            if tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
                elementos.append(run(tok[1:-1], size=size, bold=bold, italic=True))
            else:
                elementos.append(run(tok, size=size, bold=bold, italic=italic))
        return elementos

    @staticmethod
    def _mantener_con_siguiente(p):
        pp = p.find(qn("w:pPr"))
        if pp is None:
            pp = E("w:pPr")
            p.insert(0, pp)
        if pp.find(qn("w:keepNext")) is None:
            estilo = pp.find(qn("w:pStyle"))
            if estilo is not None:
                estilo.addnext(E("w:keepNext"))
            else:
                pp.insert(0, E("w:keepNext"))

    def parrafo(self, texto, size=22, bold=False, italic=False, **kw_ppr):
        p = E("w:p", None, ppr(**kw_ppr))
        for el in self.inline(texto, size=size, bold=bold, italic=italic):
            p.append(el)
        return self._agregar(p)

    # ------------------------------------------------------------ bloques
    def titulo1(self, texto, numerado=True, salto=True):
        ind = {"left": 432, "hanging": 432} if numerado else {"left": 0, "firstLine": 0}
        p = E("w:p", None, ppr(style="Heading1", keep_next=True, page_break=salto,
                                numpr=(0, NUM_CAPITULOS) if numerado else None,
                                ind=ind, jc="left", line=240, before=2000, after=480))
        p.append(run(texto, size=40, bold=True))
        return self._agregar(p)

    def titulo2(self, texto):
        p = E("w:p", None, ppr(style="Heading2", keep_next=True, keep_lines=True,
                                numpr=(1, NUM_CAPITULOS), ind={"left": 720, "hanging": 720},
                                jc="left", line=240, before=400, after=200))
        p.append(run(texto, size=32, bold=True))
        return self._agregar(p)

    def titulo3(self, texto):
        p = E("w:p", None, ppr(style="Heading3", keep_next=True, keep_lines=True,
                                numpr=(2, NUM_CAPITULOS), ind={"left": 864, "hanging": 864},
                                jc="left", line=240, before=320, after=160))
        p.append(run(texto, size=28, bold=True))
        return self._agregar(p)

    def p(self, texto, **kw):
        kw.setdefault("jc", "both")
        kw.setdefault("line", 360)
        kw.setdefault("after", 240)
        return self.parrafo(texto, **kw)

    def vinetas(self, items, after_final=240):
        for i, item in enumerate(items):
            self.parrafo(item, numpr=(0, NUM_VINETAS), ind={"left": 720, "hanging": 360},
                         jc="both", line=360, after=after_final if i == len(items) - 1 else 80)

    def ecuacion(self, latex, after=200):
        previo = self.sect_final.getprevious()
        if previo is not None and previo.tag == qn("w:p"):
            self._mantener_con_siguiente(previo)  # la frase que introduce la ecuacion no queda sola
        p = E("w:p", None, ppr(jc="center", line=276, before=60, after=after, keep_lines=True))
        om = self.mat.omml(latex, size=22)
        para = OxmlElement("m:oMathPara")
        para.append(E("m:oMathParaPr", None, E("m:jc", {"m:val": "center"})))
        para.append(om)
        p.append(para)
        return self._agregar(p)

    def leyenda(self, tipo, n, texto):
        etiqueta = "Figura" if tipo == "fig" else "Tabla"
        p = E("w:p", None, ppr(jc="both", line=276, before=120, after=120, keep_next=True,
                                ind={"left": 1134, "hanging": 1134},
                                tabs=[("left", 1134, None)]))
        p.append(run(f"{etiqueta} {n}:\t", size=22, bold=True))
        for el in self.inline(texto, size=22):
            p.append(el)
        return self._agregar(p)

    def fuente(self, texto, after=240):
        return self.parrafo(texto, size=18, jc="left", line=240, before=60, after=after)

    def figura(self, etiqueta, ruta, leyenda, fuente, ancho_cm=15.0):
        n = self.refs[f"fig:{etiqueta}"]
        self._verificar_orden("fig", n)
        self.leyenda("fig", n, leyenda)
        par = self.doc.add_paragraph()
        par.add_run().add_picture(str(ruta), width=Cm(ancho_cm))
        pel = par._p
        vieja = pel.find(qn("w:pPr"))
        if vieja is not None:
            pel.remove(vieja)
        pel.insert(0, ppr(jc="center", line=240, before=0, after=0, keep_next=True))
        self.fuente(f"Fuente: {fuente}")

    # ------------------------------------------------------------ tablas
    def tabla(self, etiqueta, leyenda, encabezados, filas, anchos, size=18, alinear=None,
              fuente=None, unida=True, encabezado_repetido=True):
        n = self.refs[f"tab:{etiqueta}"]
        self._verificar_orden("tab", n)
        self.leyenda("tab", n, leyenda)
        total = sum(anchos)
        if abs(total - ANCHO_TEXTO) > 2:
            anchos = [round(a * ANCHO_TEXTO / total) for a in anchos]
            anchos[-1] += ANCHO_TEXTO - sum(anchos)
        ncol = len(anchos)
        alinear = alinear or ["left"] * ncol
        borde = {"w:val": "single", "w:sz": 4, "w:space": 0, "w:color": "000000"}
        tbl = E("w:tbl")
        tbl_pr = E("w:tblPr", None,
                   E("w:tblStyle", {"w:val": "Table2"}),
                   E("w:tblW", {"w:w": ANCHO_TEXTO, "w:type": "dxa"}),
                   E("w:jc", {"w:val": "center"}),
                   E("w:tblBorders", None, *[E(f"w:{b}", borde) for b in
                                             ("top", "left", "bottom", "right", "insideH", "insideV")]),
                   E("w:tblLayout", {"w:type": "fixed"}),
                   E("w:tblCellMar", None,
                     E("w:top", {"w:w": 40, "w:type": "dxa"}),
                     E("w:left", {"w:w": 80, "w:type": "dxa"}),
                     E("w:bottom", {"w:w": 40, "w:type": "dxa"}),
                     E("w:right", {"w:w": 80, "w:type": "dxa"})),
                   E("w:tblLook", {"w:val": "04A0"}))
        tbl.append(tbl_pr)
        grid = E("w:tblGrid")
        for a in anchos:
            grid.append(E("w:gridCol", {"w:w": a}))
        tbl.append(grid)

        todas = [("hdr", encabezados)] + [("row", f) for f in filas]
        for k, (tipo, fila) in enumerate(todas):
            ultima = k == len(todas) - 1
            siguiente_es_grupo = not ultima and isinstance(todas[k + 1][1], dict)
            if tipo == "hdr" or isinstance(fila, dict):
                mantener = not ultima
            elif unida == "grupos":
                mantener = not ultima and not siguiente_es_grupo
            else:
                mantener = bool(unida) and not ultima
            tr = E("w:tr")
            trpr = E("w:trPr", None, E("w:cantSplit"))
            if tipo == "hdr" and encabezado_repetido:
                trpr.append(E("w:tblHeader"))
            tr.append(trpr)
            if isinstance(fila, dict):  # fila combinada (grupo)
                celdas = [(fila["grupo"], ncol, True, "EDEDED")]
            else:
                celdas = [(c, 1, tipo == "hdr", "D9D9D9" if tipo == "hdr" else None) for c in fila]
            col = 0
            for texto, span, negrita, relleno in celdas:
                ancho = sum(anchos[col:col + span])
                tcpr = E("w:tcPr", None, E("w:tcW", {"w:w": ancho, "w:type": "dxa"}))
                if span > 1:
                    tcpr.append(E("w:gridSpan", {"w:val": span}))
                if relleno:
                    tcpr.append(E("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": relleno}))
                tcpr.append(E("w:vAlign", {"w:val": "center"}))
                tc = E("w:tc", None, tcpr)
                jc = "center" if (tipo == "hdr" or span > 1) else alinear[col]
                lineas = str(texto).split("\n")
                for li, linea in enumerate(lineas):
                    par = E("w:p", None, ppr(jc=jc, line=240, before=0,
                                              after=40 if li < len(lineas) - 1 else 0,
                                              keep_next=mantener))
                    for el in self.inline(linea, size=size, bold=negrita):
                        par.append(el)
                    tc.append(par)
                tr.append(tc)
                col += span
            tbl.append(tr)
        self._agregar(tbl)
        if fuente:
            self.fuente(f"Fuente: {fuente}")
        else:
            self._agregar(E("w:p", None, ppr(line=240, after=120)))

    # ------------------------------------------------------------ indice
    def contenido(self):
        p = E("w:p", None, ppr(jc="left", line=360, before=2000, after=200, page_break=True))
        p.append(run("Contenido", size=40, bold=True))
        self._agregar(p)
        p = E("w:p", None, ppr(jc="right", line=240, before=200, after=120))
        p.append(run("Pág.", size=22, bold=True))
        self._agregar(p)
        p = E("w:p", None, ppr(jc="left", line=276, after=0))
        r1 = E("w:r", None, E("w:fldChar", {"w:fldCharType": "begin"}))
        instr = E("w:instrText")
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instr.text = ' TOC \\o "1-2" \\h \\z \\u '
        r2 = E("w:r", None, instr)
        r3 = E("w:r", None, E("w:fldChar", {"w:fldCharType": "separate"}))
        r4 = run("Haga clic derecho y seleccione «Actualizar campo» para generar el índice.")
        r5 = E("w:r", None, E("w:fldChar", {"w:fldCharType": "end"}))
        for r in (r1, r2, r3, r4, r5):
            p.append(r)
        self._agregar(p)

    # ------------------------------------------------------------ secciones
    def cerrar_seccion1(self):
        """Adjunta el salto de seccion (con los encabezados de la seccion principal)
        al ultimo parrafo agregado, para que la Bibliografia inicie en pagina nueva
        sin dejar parrafos vacios."""
        ultimo = self.sect_final.getprevious()
        if ultimo.tag != qn("w:p"):
            ultimo = self._agregar(E("w:p", None, ppr(line=240, after=0)))
        pp = ultimo.find(qn("w:pPr"))
        pp.append(copy.deepcopy(self.sect_seccion1))

    def guardar(self, ruta):
        Path(ruta).parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(ruta))
