#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Construye la plantilla Word del OFICIO MASIVO a partir de los .docx oficiales
(DOM VICHUQUEN / MINVU VARIOS / SERVIU VARIOS) y la imprime en Base64 para
incrustarla en Generador_DL2695_v2.html (OFICIO_TEMPLATES.masivo).

- Reutiliza estilos, numeracion, encabezado (logo) y configuracion de pagina
  del documento original, para que el resultado sea identico en formato.
- Sustituye el logo por la version comprimida que ya usan las otras plantillas
  del generador, para no inflar el HTML.
- El cuerpo (word/document.xml) se reescribe limpio, con tokens {campo}.
- NO incluye fecha: solo "TALCA," (requisito del usuario).
"""
import base64
import io
import glob
import re
import zipfile

ORIGEN = [f for f in glob.glob('*.docx') if f.startswith('MINVU')][0]
HTML = 'Generador_DL2695_v2.html'

# ---------------------------------------------------------------- formatos
NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'mc:Ignorable="w14 w15 w16se wp14"'
)

FONT = '<w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/>'
SZ = '<w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="es-ES"/>'
RPR_N = '<w:rPr>' + FONT + SZ + '</w:rPr>'
RPR_B = '<w:rPr>' + FONT + '<w:b/><w:bCs/>' + SZ + '</w:rPr>'
RPR_I = '<w:rPr>' + FONT + '<w:i/><w:iCs/>' + SZ + '</w:rPr>'
RPR_U = '<w:rPr>' + FONT + '<w:u w:val="single"/>' + SZ + '</w:rPr>'

# formato de las celdas de la tabla (Verdana 9)
CELL_RPR = ('<w:rPr><w:rFonts w:ascii="Verdana" w:eastAsia="Times New Roman" '
            'w:hAnsi="Verdana" w:cs="Calibri"/><w:color w:val="000000"/>'
            '<w:sz w:val="18"/><w:szCs w:val="18"/><w:lang w:eastAsia="es-CL"/></w:rPr>')
CELL_PPR = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>' + CELL_RPR + '</w:pPr>'


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def run(texto, rpr=RPR_N):
    return '<w:r>' + rpr + '<w:t xml:space="preserve">' + esc(texto) + '</w:t></w:r>'


def tab(rpr=RPR_N):
    return '<w:r>' + rpr + '<w:tab/></w:r>'


def par(ppr_inner, runs=''):
    ppr = '<w:pPr>' + ppr_inner + RPR_N + '</w:pPr>' if ppr_inner is not None else ''
    return '<w:p>' + ppr + runs + '</w:p>'


IND_ORD = '<w:ind w:left="1843" w:firstLine="2268"/><w:jc w:val="both"/>'
IND_ANT = '<w:ind w:left="5529"/><w:jc w:val="both"/>'
SP = '<w:spacing w:line="240" w:lineRule="auto"/>'
CUERPO = '<w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="708"/><w:jc w:val="both"/>'
CUERPO2 = '<w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="461"/><w:jc w:val="both"/>'
LISTA = ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
         '<w:spacing w:line="240" w:lineRule="auto"/>'
         '<w:ind w:left="0" w:firstLine="0"/><w:jc w:val="both"/>')
CIERRE = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/>'

# ---------------------------------------------------------------- tabla
COLS = [('EXP.', 932), ('RUT', 1574), ('SOLICITANTE', 1565),
        ('UBICACIÓN', 1644), ('ROL', 1573), ('INSCRIPCIÓN', 1638)]
TOKENS = ['{exp}', '{rut}', '{solicitante}', '{ubicacion}', '{rol}', '{inscripcion}']


def celda(texto, ancho, primera):
    izq = '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' if primera else '<w:left w:val="nil"/>'
    borders = ('<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               + izq +
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>')
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>'
            '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc>'
            % (ancho, borders, CELL_PPR, CELL_RPR, esc(texto)))


def fila(valores):
    celdas = ''.join(celda(v, COLS[i][1], i == 0) for i, v in enumerate(valores))
    return '<w:tr><w:trPr><w:trHeight w:val="300"/></w:trPr>' + celdas + '</w:tr>'


TBL = (
    '<w:tbl><w:tblPr><w:tblW w:w="8926" w:type="dxa"/>'
    '<w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar>'
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" '
    'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
    '<w:tblGrid>' + ''.join('<w:gridCol w:w="%d"/>' % c[1] for c in COLS) + '</w:tblGrid>'
    + fila([c[0] for c in COLS])
    + fila(TOKENS)
    + '</w:tbl>'
)

# ---------------------------------------------------------------- textos fijos
T_ANT = (' Decreto Ley N°2695 de 1979; Dictamen E472530/2024, de la Contraloría General de la '
         'República, sobre aplicación del procedimiento de regularización contemplado en el '
         'Decreto Ley N°2695 de 1979, aclara Dictámenes N°42.084 de 2017 y complementa '
         'Dictámenes N°s 11.662, de 1985, y 26.443 de 1987, en los términos que se indican.')

T_INTRO = ('Junto con saludar cordialmente, por medio del presente, informo que {intro_frase} '
           'presentado solicitud de regularización de la pequeña propiedad raíz, conforme a las '
           'disposiciones contenidas en el Decreto Ley N°2695 de 1979, ante esta Secretaría '
           'Regional Ministerial.')

T_AL_RESPECTO = ('Al respecto, en relación a la aplicación del procedimiento de regularización '
                 'contemplado en el Decreto Ley N’ 2695 de 1979, la Contraloría General de la '
                 'República emitió el Dictamen N° E472530/2024, que aclara Dictamen N°42.084, de '
                 '2017, y complementa Dictámenes N°s11.662, de 1985, y 26.443, de 1987, '
                 'estableciendo el derecho a regularización en los términos que dicho '
                 'pronunciamiento establece.')

T_COORD = ('En virtud del principio de coordinación que debe existir entre los diversos órganos de '
           'la Administración del Estado, este Servicio debe requerir información a entidades con '
           'potestad urbanística, y en consideración a lo señalado, solicitamos a usted, tenga a '
           'bien informar lo siguiente:')

T_P1_A = 'Si el inmueble objeto de regularización es parte de un '
T_P1_B = 'loteo irregular o subdivisión de hecho que pueda generar un nuevo núcleo urbano.'

T_P2_A = 'Esto porque el dictamen N°E472530/2024 señala que “'
T_P2_B = ('la no aplicación del artículo 55 de la LGUC a las regularizaciones de que trata el citado '
          'decreto ley N° 2.695, de 1979, implica que la generación de estos nuevos predios no esté '
          'sujeta a cabida de subdivisión -ya que tampoco esa generación predial se someterá a las '
          'reglas concernientes a dicha norma urbanística para los terrenos de que se trate, en '
          'particular, a las aprobaciones e informes favorables exigidos por aquel precepto para '
          'esas zonas-, ni a exigencias de urbanización, propias de esos procesos de división. En '
          'todo caso, tales excepciones, no aplican en los casos de loteos irregulares de predios '
          'rurales, pues ello conllevaría, en definitiva, una alteración significativa de la '
          'planificación territorial.”.')

T_P3_A = ('En el caso de que el inmueble objeto de regularización, en vuestra opinión técnica '
          'urbanística, sea parte de un ')
T_P3_B = 'loteo irregular o subdivisión de hecho que pueda generar un nuevo núcleo urbano, '
T_P3_C = ('se solicita que usted señale si éste se ha acogido a las normas y procedimiento dispuesto '
          'en la Ley N° 20.234; y de ser efectivo esto último, si ha obtenido la recepción provisoria '
          'o definitiva por parte de la Dirección de Obras Municipales respectiva, en virtud a lo '
          'dispuesto en el artículo 4 inciso 17 de este mismo cuerpo normativo. Lo anterior, con el '
          'fin de dar curso o no a la solicitud de regularización presentada.')

T_P4 = ('Por último, se solicita nos pueda informar cualquier otra situación que requiera especial '
        'consideración, en relación a la zona donde se ubica el inmueble.')

T_FINAL = ('Finalmente, y con el objeto de que pueda contar con la mayor cantidad de antecedentes '
           'para un correcto análisis, se adjuntan los siguientes antecedentes acompañados por '
           '{ref_solicitante}: i) Certificado de Avalúo; ii) Certificado Informaciones Previas y '
           'iii) Inscripción conservatoria.')

SECTPR = ('<w:sectPr><w:headerReference w:type="default" r:id="rId7"/>'
          '<w:pgSz w:w="12240" w:h="18720" w:code="41"/>'
          '<w:pgMar w:top="1417" w:right="1701" w:bottom="2410" w:left="1701" '
          'w:header="454" w:footer="454" w:gutter="0"/>'
          '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>')

# ---------------------------------------------------------------- cuerpo
cuerpo = []
cuerpo.append(par(IND_ORD, run('                     ORD. N°SE07-{ord_n} - {ord_anio}.', RPR_B)))
cuerpo.append(par(IND_ANT, run('ANT.:', RPR_B) + run(T_ANT)))
cuerpo.append(par(IND_ANT, run('MAT.:', RPR_B) + run(' Solicita información que indica.')))
cuerpo.append(par(IND_ANT, run('TALCA,')))
cuerpo.append(par(IND_ANT))
cuerpo.append(par(SP, run('DE:', RPR_B) + tab(RPR_B)
                  + run('SECRETARIO REGIONAL MINISTERIAL DE BIENES NACIONALES REGIÓN DEL MAULE.', RPR_B)))
cuerpo.append(par(SP, run('A:', RPR_B) + tab() + run('{destinatario}', RPR_B)))
cuerpo.append(par(SP))
cuerpo.append(par(CUERPO, run(T_INTRO)))
cuerpo.append(par(CUERPO))
cuerpo.append(TBL)
cuerpo.append(par(CUERPO))
cuerpo.append(par(CUERPO, run(T_AL_RESPECTO)))
cuerpo.append(par(CUERPO, run(T_COORD)))
cuerpo.append(par(LISTA, run(T_P1_A) + run(T_P1_B, RPR_I)))
cuerpo.append(par('<w:spacing w:line="240" w:lineRule="auto"/><w:jc w:val="both"/>',
                  run(T_P2_A) + run(T_P2_B, RPR_I)))
cuerpo.append(par(LISTA, run(T_P3_A) + run(T_P3_B, RPR_I) + run(T_P3_C)))
cuerpo.append(par(LISTA, run(T_P4)))
cuerpo.append(par(CUERPO2, run(T_FINAL)))
cuerpo.append(par(CUERPO2, run('Saluda atenta y cordialmente a usted,')))
cuerpo.append(par('<w:spacing w:line="240" w:lineRule="auto"/><w:jc w:val="both"/>'))
for _ in range(9):
    cuerpo.append(par(CIERRE))
cuerpo.append(par(CIERRE, run('{iniciales}')))
cuerpo.append(par(CIERRE, run('DISTRIBUCIÓN', RPR_U) + run(': ')))
cuerpo.append(par(CIERRE, run('Destinatario.')))
cuerpo.append(par(CIERRE, run('Expediente')))
cuerpo.append(par(CIERRE, run('Unidad de Regularización.')))

DOC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
       '<w:document ' + NS + '><w:body>' + ''.join(cuerpo) + SECTPR + '</w:body></w:document>')

# ---------------------------------------------------------------- logo liviano
html = open(HTML, encoding='utf-8').read()
m = re.search(r'const OFICIO_TEMPLATES = \{(.*?)\n\};', html, re.S)
plantillas = dict(re.findall(r'(\w+):\s*"([^"]+)"', m.group(1)))
zdom = zipfile.ZipFile(io.BytesIO(base64.b64decode(plantillas['dom'])))
logo = zdom.read('word/media/image1.jpeg')
print('logo reutilizado:', len(logo), 'bytes')

# ---------------------------------------------------------------- armar docx
src = zipfile.ZipFile(ORIGEN)
COPIAR = ['[Content_Types].xml', '_rels/.rels', 'word/_rels/document.xml.rels',
          'word/_rels/header1.xml.rels', 'word/header1.xml', 'word/styles.xml',
          'word/numbering.xml', 'word/settings.xml', 'word/webSettings.xml',
          'word/fontTable.xml', 'word/footnotes.xml', 'word/endnotes.xml',
          'word/theme/theme1.xml', 'docProps/app.xml', 'docProps/core.xml']

buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as out:
    out.writestr('word/document.xml', DOC.encode('utf-8'))
    out.writestr('word/media/image1.jpeg', logo)
    for n in COPIAR:
        out.writestr(n, src.read(n))

data = buf.getvalue()
print('docx generado:', len(data), 'bytes')

# verificacion
z = zipfile.ZipFile(io.BytesIO(data))
assert z.testzip() is None, 'CRC invalido'
from xml.etree import ElementTree as ET
ET.fromstring(z.read('word/document.xml'))
faltan = [t for t in TOKENS + ['{ord_n}', '{ord_anio}', '{destinatario}', '{intro_frase}',
                               '{ref_solicitante}', '{iniciales}']
          if t not in DOC]
assert not faltan, 'tokens faltantes: %s' % faltan
print('verificacion OK — entradas:', len(z.namelist()))

b64 = base64.b64encode(data).decode('ascii')
open('masivo.b64', 'w').write(b64)
print('base64:', len(b64), 'chars -> masivo.b64')
open('masivo_preview.docx', 'wb').write(data)
