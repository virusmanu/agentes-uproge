"""Escritor mínimo de documentos Word (.docx), sin dependencias externas.

Formato de minuta UPROGE: logo del Ministerio en el encabezado, título centrado en mayúsculas con
"MINISTERIO DE AGRICULTURA", línea de fecha, secciones numeradas en negrita, tablas con la columna izquierda
sombreada y pie "Página | X"."""
import os, zipfile, datetime
from xml.sax.saxutils import escape

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FUENTE = 'Arial'
MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']


def _run(t, b=False, i=False, tam=None, color=None):
    pr = f'<w:rFonts w:ascii="{FUENTE}" w:hAnsi="{FUENTE}" w:cs="{FUENTE}"/>' + ('<w:b/>' if b else '') + ('<w:i/>' if i else '') \
         + (f'<w:color w:val="{color}"/>' if color else '') + (f'<w:sz w:val="{int(tam * 2)}"/><w:szCs w:val="{int(tam * 2)}"/>' if tam else '')
    partes = str(t).split('\n')
    x = ''.join((f'<w:br/>' if k else '') + f'<w:t xml:space="preserve">{escape(p)}</w:t>' for k, p in enumerate(partes))
    return f'<w:r><w:rPr>{pr}</w:rPr>{x}</w:r>'


def _par(runs, al='both', antes=0, despues=120, sangria=None):
    ind = f'<w:ind w:left="{sangria}"/>' if sangria else ''
    return f'<w:p><w:pPr><w:spacing w:before="{antes}" w:after="{despues}" w:line="276" w:lineRule="auto"/>{ind}<w:jc w:val="{al}"/></w:pPr>{runs}</w:p>'


class Documento:
    def __init__(self, titulo, fecha=None, logo=True):
        self.cuerpo = []; self.logo = os.path.join(RAIZ, 'recursos', 'logo.png') if logo else None
        if self.logo and not os.path.exists(self.logo): self.logo = None
        f = fecha or datetime.date.today()
        self.cuerpo.append(_par(_run(titulo.upper(), b=True, tam=13), al='center', despues=0))
        self.cuerpo.append(_par(_run('MINISTERIO DE AGRICULTURA', b=True, tam=11), al='center', despues=240))
        self.cuerpo.append(_par(_run('Fecha: ', b=True) + _run(f'{f.day} de {MESES[f.month - 1]} de {f.year}'), al='left', despues=240))

    def seccion(self, texto):
        self.cuerpo.append(_par(_run(texto, b=True, tam=11), al='left', antes=200, despues=100))

    def parrafo(self, texto, b=False, i=False, color=None, tam=10.5, al='both'):
        self.cuerpo.append(_par(_run(texto, b=b, i=i, color=color, tam=tam), al=al))

    def parrafo_mixto(self, partes, al='both'):
        """partes: [(texto, negrita)]"""
        self.cuerpo.append(_par(''.join(_run(t, b=b, tam=10.5) for t, b in partes), al=al))

    def vineta(self, texto):
        self.cuerpo.append(_par(_run('•  ' + texto, tam=10.5), sangria=284, despues=60))

    def tabla(self, encabezado, filas, anchos=None, alinear=None, total=False):
        """filas: listas de textos. La primera columna va sombreada; la última fila en negrita si total=True."""
        n = len(encabezado); anchos = anchos or [int(9360 / n)] * n; alinear = alinear or (['left'] + ['right'] * (n - 1))
        borde = ''.join(f'<w:{b} w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' for b in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
        def celda(t, j, fondo=None, b=False):
            sh = f'<w:shd w:val="clear" w:color="auto" w:fill="{fondo}"/>' if fondo else ''
            return (f'<w:tc><w:tcPr><w:tcW w:w="{anchos[j]}" w:type="dxa"/>{sh}<w:vAlign w:val="center"/></w:tcPr>'
                    + _par(_run(t, b=b, tam=9), al=alinear[j], despues=0) + '</w:tc>')
        x = (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(anchos)}" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>{borde}</w:tblBorders>'
             '<w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'
             + ''.join(f'<w:gridCol w:w="{a}"/>' for a in anchos) + '</w:tblGrid>')
        x += '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + ''.join(celda(t, j, '1F4E5F' if False else 'D9E2F3', True) for j, t in enumerate(encabezado)) + '</w:tr>'
        for k, f in enumerate(filas):
            ult = total and k == len(filas) - 1
            x += '<w:tr>' + ''.join(celda(t, j, 'F2F2F2' if j == 0 or ult else None, ult) for j, t in enumerate(f)) + '</w:tr>'
        self.cuerpo.append(x + '</w:tbl>'); self.cuerpo.append(_par('', despues=60))

    def guardar(self, ruta):
        ns = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
              'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
              'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
        sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'
                '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1700" w:right="1440" w:bottom="1300" w:left="1440" w:header="560" w:footer="560" w:gutter="0"/></w:sectPr>')
        doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {ns}><w:body>{"".join(self.cuerpo)}{sect}</w:body></w:document>'
        if self.logo:
            cx, cy = 762000, 696000   # 0,8 x 0,73 pulgadas aprox. (logo 70x64)
            img = (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="1" name="Logo"/>'
                   f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>'
                   f'<pic:blipFill><a:blip r:embed="rIdL"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                   f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
        else:
            img = ''
        hdr = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {ns}><w:p><w:pPr><w:jc w:val="left"/></w:pPr>{img}</w:p></w:hdr>'
        campo = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                 '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
        ftr = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {ns}><w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
               f'{_run("Página | ", tam=9)}{campo}</w:p></w:ftr>')
        estilos = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {ns}><w:docDefaults><w:rPrDefault><w:rPr>'
                   f'<w:rFonts w:ascii="{FUENTE}" w:hAnsi="{FUENTE}" w:cs="{FUENTE}"/><w:sz w:val="21"/><w:lang w:val="es-CL"/></w:rPr></w:rPrDefault>'
                   '<w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>'
                   '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        rels_doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                    '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
                    '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
        rels_hdr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    + ('<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>' if self.logo else '')
                    + '</Relationships>')
        tipos = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                 '<Default Extension="png" ContentType="image/png"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                 '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
                 '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        tmp = ruta + '.tmp'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', tipos); z.writestr('_rels/.rels', rels)
            z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', estilos)
            z.writestr('word/header1.xml', hdr); z.writestr('word/footer1.xml', ftr)
            z.writestr('word/_rels/document.xml.rels', rels_doc); z.writestr('word/_rels/header1.xml.rels', rels_hdr)
            if self.logo: z.write(self.logo, 'word/media/logo.png')
        try:
            os.replace(tmp, ruta)
        except PermissionError:
            os.remove(tmp)
            raise PermissionError(f'No pude guardar {os.path.basename(ruta)}: está abierto en Word. Ciérrelo e intente de nuevo.')
        return ruta
