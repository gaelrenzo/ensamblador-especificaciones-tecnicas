"""
generar_ejemplo.py - Generador de Datos de Prueba para el Sistema de Especificaciones

Crea:
1. 'Maestro_Especificaciones.docx': Documento maestro completo con 6 partidas detalladas,
   tablas con bordes y sombreados, diagramas/imágenes incrustadas y estilos formales.
2. 'Plantilla_Base.docx': Portada institucional y encabezados de proyecto.
3. 'Solicitud_Partidas.xlsx': Archivo Excel con la columna 'Código de Partida',
   solicitando un subconjunto específico en orden particular y celdas vacías de prueba.
"""

import os
import pandas as pd
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn


def aplicar_sombreado_celda(cell, color_hex="1B365D"):
    """Aplica color de fondo a una celda de tabla."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)


def formatear_celda_encabezado(cell, texto):
    """Estiliza una celda como encabezado de tabla profesional."""
    aplicar_sombreado_celda(cell, "1F4E79")
    cell.text = texto
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if p.runs:
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)


def crear_documento_maestro(ruta_salida: str = "Maestro_Especificaciones.docx"):
    """Genera un documento Word maestro realista con especificaciones técnicas completas."""
    doc = docx.Document()
    
    # Configurar márgenes estándar A4
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Título principal del Maestro
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_titulo = p_titulo.add_run("ESPECIFICACIONES TÉCNICAS GENERALES (DOCUMENTO MAESTRO)")
    r_titulo.font.name = "Arial"
    r_titulo.font.size = Pt(16)
    r_titulo.font.bold = True
    r_titulo.font.color.rgb = RGBColor(27, 54, 93)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("PROYECTO: INFRAESTRUCTURA Y SERVICIOS CIVILES")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph("─" * 55)
    
    # -------------------------------------------------------------
    # PARTIDA 1
    # -------------------------------------------------------------
    h1 = doc.add_heading("01.01.01 CARTEL DE IDENTIFICACIÓN DE OBRA DE 3.60m x 2.40m", level=1)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)
    
    doc.add_paragraph(
        "1.1 Descripción:\n"
        "Comprende la confección e instalación del cartel informativo de la obra, el cual "
        "deberá ubicarse en un lugar visible de fácil acceso seleccionado por la Supervisión. "
        "El diseño, logotipo, colores y leyendas serán proporcionados por la Entidad Contratante."
    )
    doc.add_paragraph(
        "1.2 Materiales y Estructura:\n"
        "Estará construido con bastidor de madera tornillo de 2\" x 3\", parantes de 4\" x 4\" "
        "anclados en dados de concreto ciclópeo de 0.40m x 0.40m x 0.60m. La gigantografía será "
        "en lona banner de 13 onzas de alta resolución resistente a intemperie."
    )
    
    # Tabla de especificaciones técnicas
    tbl1 = doc.add_table(rows=4, cols=3)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.style = "Table Grid"
    
    headers1 = ["Componente", "Especificación Técnica", "Unidad"]
    for j, h in enumerate(headers1):
        formatear_celda_encabezado(tbl1.cell(0, j), h)
        
    datos1 = [
        ["Parantes principales", "Madera Tornillo 4\" x 4\" x 4.50m", "Pieza"],
        ["Bastidor de soporte", "Listonería 2\" x 3\" seca", "Pie cuadrado"],
        ["Banner publicitario", "Lona mate 13 oz con protección UV", "m2"]
    ]
    for row_idx, fila in enumerate(datos1, 1):
        for col_idx, valor in enumerate(fila):
            c = tbl1.cell(row_idx, col_idx)
            c.text = valor
            p = c.paragraphs[0]
            if p.runs:
                p.runs[0].font.name = "Arial"
                p.runs[0].font.size = Pt(9)
                
    doc.add_paragraph(
        "1.3 Método de Medición y Base de Pago:\n"
        "La unidad de medida para el pago será la Unidad (UND), pagándose al precio unitario "
        "del contrato tras la verificación física de su correcta instalación por el Supervisor."
    )
    doc.add_paragraph("")

    # -------------------------------------------------------------
    # PARTIDA 2 (Con Imagen/Diagrama)
    # -------------------------------------------------------------
    h2 = doc.add_heading("01.01.02 CASETA DE ALMACÉN Y GUARDIANÍA", level=1)
    doc.add_paragraph(
        "2.1 Descripción:\n"
        "Comprende la construcción de ambientes provisionales adecuados para el depósito de "
        "materiales sensibles a la humedad (cemento, acero, aditivos) y custodia de herramientas, "
        "así como oficinas técnicas y ambiente de guardianía diurna y nocturna."
    )
    
    # Incrustar imagen técnica si está disponible
    img_path = "diag1_exact.png"
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(img_path, width=Inches(4.5))
        p_cap = doc.add_paragraph("Figura 1: Esquema de distribución de caseta y zona de acopio temporal.")
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(8)
        if p_cap.runs:
            p_cap.runs[0].font.italic = True
            p_cap.runs[0].font.size = Pt(8.5)
            
    doc.add_paragraph(
        "2.2 Requisitos Funcionales:\n"
        "La caseta dispondrá de ventilación cruzada, piso elevado mínimo 10 cm sobre terreno "
        "natural para evitar anegamientos, cerraduras de seguridad y extintor PQS de 6 kg vigente."
    )
    doc.add_paragraph("")

    # -------------------------------------------------------------
    # PARTIDA 3
    # -------------------------------------------------------------
    h3 = doc.add_heading("01.02.01 LIMPIEZA Y DESBROCE MANUAL DE TERRENO", level=1)
    doc.add_paragraph(
        "3.1 Descripción:\n"
        "Trabajos previos consistentes en la eliminación de malezas, raíces, desmontes y elementos "
        "que obstaculicen la visibilidad o impidan el correcto replanteo del terreno natural."
    )
    
    tbl2 = doc.add_table(rows=3, cols=4)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.style = "Table Grid"
    headers2 = ["Zona", "Área (m2)", "Espesor Desbroce", "Tipo de Vegetación"]
    for j, h in enumerate(headers2):
        formatear_celda_encabezado(tbl2.cell(0, j), h)
    datos2 = [
        ["Sector A (Plataforma)", "850.00", "0.15 m", "Maleza media"],
        ["Sector B (Accesos)", "320.00", "0.10 m", "Césped y arbustos"]
    ]
    for row_idx, fila in enumerate(datos2, 1):
        for col_idx, valor in enumerate(fila):
            c = tbl2.cell(row_idx, col_idx)
            c.text = valor
            if c.paragraphs[0].runs:
                c.paragraphs[0].runs[0].font.size = Pt(9)
                
    doc.add_paragraph(
        "3.2 Disposición de Desechos:\n"
        "Todo material vegetal y desmonte será transportado al botadero municipal autorizado "
        "por la municipalidad distrital en camión volquete."
    )
    doc.add_paragraph("")

    # -------------------------------------------------------------
    # PARTIDA 4 (Con Diagrama 2)
    # -------------------------------------------------------------
    h4 = doc.add_heading("01.02.02 TRAZO, NIVELES Y REPLANTEO PRELIMINAR", level=1)
    doc.add_paragraph(
        "4.1 Descripción:\n"
        "Consiste en llevar al terreno los ejes, rasantes y alineamientos señalados en los "
        "planos de arquitectura y estructuras, empleando estación total y nivel de ingeniero."
    )
    
    img_path2 = "diag2_exact.png"
    if os.path.exists(img_path2):
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(img_path2, width=Inches(4.5))
        p_cap2 = doc.add_paragraph("Figura 2: Diagrama de control topográfico y tolerancias de cota.")
        p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p_cap2.runs:
            p_cap2.runs[0].font.italic = True
            p_cap2.runs[0].font.size = Pt(8.5)
            
    doc.add_paragraph(
        "4.2 Tolerancias:\n"
        "Las discrepancias máximas admisibles serán de ± 5 mm en nivelación y ± 10 segundos "
        "sexagesimales en ángulos azimutales entre vértices de poligonal."
    )
    doc.add_paragraph("")

    # -------------------------------------------------------------
    # PARTIDA 5
    # -------------------------------------------------------------
    h5 = doc.add_heading("02.01.01 EXCAVACIÓN MANUAL PARA ZANJAS Y CIMIENTOS", level=1)
    doc.add_paragraph(
        "5.1 Descripción y Alcance:\n"
        "Comprende la excavación manual para alojar cimientos corridos y zapatas aisladas "
        "hasta los niveles de fondo de cimentación indicados en planos estructurales."
    )
    tbl3 = doc.add_table(rows=3, cols=3)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.style = "Table Grid"
    headers3 = ["Elemento Estructural", "Profundidad (m)", "Capacidad Portante (kg/cm2)"]
    for j, h in enumerate(headers3):
        formatear_celda_encabezado(tbl3.cell(0, j), h)
    datos3 = [
        ["Cimiento Corrido CC-1", "1.20 m", "1.80"],
        ["Zapatas Z-1 y Z-2", "1.60 m", "2.10"]
    ]
    for row_idx, fila in enumerate(datos3, 1):
        for col_idx, valor in enumerate(fila):
            c = tbl3.cell(row_idx, col_idx)
            c.text = valor
            if c.paragraphs[0].runs:
                c.paragraphs[0].runs[0].font.size = Pt(9)
                
    doc.add_paragraph(
        "5.2 Medición:\n"
        "La unidad de medida para el pago es el Metro Cúbico (m3) de volumen excavado neto."
    )
    doc.add_paragraph("")

    # -------------------------------------------------------------
    # PARTIDA 6
    # -------------------------------------------------------------
    h6 = doc.add_heading("02.01.02 RELLENO Y COMPACTACIÓN CON MATERIAL PROPIO", level=1)
    doc.add_paragraph(
        "6.1 Descripción:\n"
        "Comprende la colocación de material selecto proveniente de las excavaciones, libre "
        "de materia orgánica y piedras mayores a 2.5 pulgadas, en capas sucesivas de 20 cm."
    )
    doc.add_paragraph(
        "6.2 Control de Calidad:\n"
        "Se realizarán ensayos de densidad de campo (cono de arena ASTM D1556) exigiendo un "
        "grado de compactación mínimo del 95% del Proctor Modificado (AASHTO T-180)."
    )
    doc.add_paragraph("")
    
    doc.save(ruta_salida)
    print(f"✔ Documento Maestro generado: '{ruta_salida}'")


def crear_plantilla_base(ruta_salida: str = "Plantilla_Base.docx"):
    """Crea una plantilla base con carátula institucional y membretes."""
    doc = docx.Document()
    
    # Espaciado superior
    for _ in range(4):
        doc.add_paragraph("")
        
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run("REPÚBLICA DEL PERÚ\nGOBIERNO REGIONAL DE INFRAESTRUCTURA")
    r1.font.name = "Arial"
    r1.font.size = Pt(13)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(80, 80, 80)
    
    for _ in range(2):
        doc.add_paragraph("")
        
    p_exp = doc.add_paragraph()
    p_exp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_exp = p_exp.add_run("EXPEDIENTE TÉCNICO CONSOLIDADO\nDE ESPECIFICACIONES TÉCNICAS")
    r_exp.font.name = "Arial"
    r_exp.font.size = Pt(20)
    r_exp.font.bold = True
    r_exp.font.color.rgb = RGBColor(27, 54, 93)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("DOCUMENTO TÉCNICO OFICIAL")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(12)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(180, 100, 20)
    
    for _ in range(6):
        doc.add_paragraph("")
        
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run(
        "PROYECTO: MEJORAMIENTO INTEGRAL DEL SISTEMA DE DISTRIBUCIÓN\n"
        "SOLICITANTE: DIRECCIÓN DE OBRAS PÚBLICAS\n"
        "FECHA: OCTUBRE 2026\n"
    )
    r_meta.font.name = "Arial"
    r_meta.font.size = Pt(10)
    r_meta.font.color.rgb = RGBColor(60, 60, 60)
    
    # Salto de página para que las partidas empiecen en la página siguiente
    doc.add_page_break()
    
    doc.save(ruta_salida)
    print(f"✔ Plantilla Base generada: '{ruta_salida}'")


def crear_solicitud_excel(ruta_salida: str = "Solicitud_Partidas.xlsx"):
    """Crea una plantilla Excel con una lista específica de partidas solicitadas."""
    datos = {
        "Item": [1, 2, 3, 4, None, 5, None],
        "Código de Partida": [
            "01.01.01",       # Cartel de obra
            "01.02.02",       # Trazo y replanteo (salta 01.01.02 y 01.02.01 a propósito)
            "02.01.01",       # Excavación manual
            "  02.01.02  ",   # Relleno con espacios para probar limpieza
            None,             # Fila vacía para probar eliminación de nulos
            "01.01.02",       # Caseta (con diagrama 1) al final
            ""                # Celda vacía
        ],
        "Descripción Referencial": [
            "Cartel de identificación de obra",
            "Trazo y replanteo topográfico con estación total",
            "Excavación manual en zanjas para cimientos",
            "Relleno y compactación con material propio",
            None,
            "Caseta de almacén y guardianía provisional",
            None
        ],
        "Unidad": ["UND", "M2", "M3", "M3", None, "GLB", None],
        "Metrado Solicitado": [1.0, 1170.0, 48.5, 32.0, None, 1.0, None]
    }
    
    df = pd.DataFrame(datos)
    df.to_excel(ruta_salida, index=False, engine="openpyxl")
    print(f"✔ Archivo de Solicitud Excel generado: '{ruta_salida}'")


def main():
    print("==================================================")
    print(" GENERADOR DE ARCHIVOS DE PRUEBA Y DEMOSTRACIÓN")
    print("==================================================")
    crear_documento_maestro("Maestro_Especificaciones.docx")
    crear_plantilla_base("Plantilla_Base.docx")
    crear_solicitud_excel("Solicitud_Partidas.xlsx")
    print("--------------------------------------------------")
    print("Archivos de prueba generados exitosamente.")
    print("==================================================")


if __name__ == "__main__":
    main()
