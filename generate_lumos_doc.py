import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def set_cell_background(cell, hex_color):
    """Establece el color de fondo de una celda."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    """Establece márgenes internos de celda en dxa."""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="D1D5DB", sz="4", val="single"):
    """Establece bordes sutiles para una tabla."""
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:left w:val="none"/>
                <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:right w:val="none"/>
                <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideV w:val="none"/>
            </w:tblBorders>
        ''')
        tblPr[0].append(borders)

def add_header_styled(doc, text, level=1):
    """Añade encabezados con estilo y color institucional."""
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    
    run = p.add_run(text)
    run.bold = True
    if level == 1:
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(15, 23, 42) # Midnight blue
        p.paragraph_format.space_before = Pt(20)
    elif level == 2:
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(180, 130, 30) # Antique Gold
    elif level == 3:
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(51, 65, 85) # Slate
    return p

def add_callout(doc, title, text, bg_color="F8FAFC", border_color="E6B758"):
    """Añade una caja de llamada destacada."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=160, bottom=160, left=220, right=200)
    
    tcPr = cell._element.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
            <w:top w:val="none"/>
            <w:right w:val="none"/>
            <w:bottom w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    r_title = p.add_run(title + "\n")
    r_title.bold = True
    r_title.font.size = Pt(10)
    r_title.font.color.rgb = RGBColor(15, 23, 42)
    
    r_text = p.add_run(text)
    r_text.font.size = Pt(9.5)
    r_text.font.color.rgb = RGBColor(71, 85, 105)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_lumus_doc():
    doc = Document()
    
    # Configuración de márgenes de página
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.85)
        s.right_margin = Inches(0.85)
        
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(5)

    print("Generando portada...")
    p_pre = doc.add_paragraph()
    p_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pre.paragraph_format.space_before = Pt(15)
    p_pre.paragraph_format.space_after = Pt(4)
    r_pre = p_pre.add_run("PROYECTO TECNOLÓGICO INDUSTRIAL & PLAN DE ESCALABILIDAD")
    r_pre.font.size = Pt(9.5)
    r_pre.font.bold = True
    r_pre.font.color.rgb = RGBColor(180, 130, 30)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("L U M U S")
    r_title.font.size = Pt(36)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(15, 23, 42)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Auriculares Inteligentes con Iluminación LED & Sincronización Escénica")
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)
    
    if os.path.exists("lumos_img/hero.jpg"):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(14)
        run_img = p_img.add_run()
        run_img.add_picture("lumos_img/hero.jpg", width=Inches(5.2))
        
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_before = Pt(8)
    r_meta = p_meta.add_run("Documentación de Ingeniería de Producto, Viabilidad Económica y Hoja de Ruta Industrial\nVersión Oficial 1.0 — Octubre 2026 | Argentina")
    r_meta.font.size = Pt(9)
    r_meta.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_page_break()

    # ÍNDICE GENERAL DEL DOCUMENTO
    add_header_styled(doc, "Índice General de Contenidos", level=1)
    
    p_idx_desc = doc.add_paragraph(
        "A continuación se detalla la estructura y navegación temática del presente documento técnico:"
    )
    p_idx_desc.paragraph_format.space_after = Pt(8)
    
    index_items = [
        ("1. Identificación del Proyecto, Integrantes y Plataforma Web", "Pág. 3", [
            ("1.1 Integrantes del Proyecto", ""),
            ("1.2 Plataforma Web Oficial y Recursos Digitales", "")
        ]),
        ("2. Resumen Ejecutivo y Tesis de Negocio", "Pág. 4", [
            ("2.1 Tesis de Producto: De Dispositivo Pasivo a Elemento Escénico", "")
        ]),
        ("3. Especificaciones Técnicas y Arquitectura Electroacústica", "Pág. 4", [
            ("3.1 Transductores de Audio y Desempeño Acústico", ""),
            ("3.2 Subsistema de Iluminación Óptica y Microcontrolador", ""),
            ("3.3 Batería, Circuito de Protección y Gestión Térmica", "")
        ]),
        ("4. Arquitectura de Conectividad Dual y Versatilidad de Uso", "Pág. 5", [
            ("4.1 Modo LUMUS Wireless (Bluetooth 5.3 Low Latency)", ""),
            ("4.2 Modo LUMUS Dual Hi-Res (Jack 3.5 mm Pasivo)", "")
        ]),
        ("5. Catálogo Estético: Las Cinco Órdenes de Color", "Pág. 6", [
            ("5.1 Orden del Fuego Solar (Rubí & Oro Real)", ""),
            ("5.2 Orden del Vacío Astral (Azul Medianoche & Zafiro)", ""),
            ("5.3 Orden del Bosque Esmeralda (Jade & Oro Antiguo)", ""),
            ("5.4 Orden de la Nebulosa Rosa (Perla & Oro Rosa)", ""),
            ("5.5 Orden Signature / Eclipse (Negro Carbón & Oro Bruñido)", "")
        ]),
        ("6. Extensiones de Línea: Glide Fit y Nova Mini", "Pág. 7", [
            ("6.1 LUMUS Glide Fit (Open-Ear Deportivo para Patinaje)", ""),
            ("6.2 LUMUS Nova Mini (TWS Intrauditivo Ultracompacto)", "")
        ]),
        ("7. Plan Integral de Escalabilidad Técnica e Industrial (Roadmap 3 Años)", "Pág. 8", [
            ("7.1 Fase 1: Validación de Mercado & Adaptación ODM/OEM (Meses 0 a 6)", ""),
            ("7.2 Fase 2: Firmware Propietario & Sincronización Escénica Mesh (Meses 6 a 18)", ""),
            ("7.3 Fase 3: Fabricación Propia, Moldes Propios y Expansión Latam (Meses 18 a 36)", "")
        ]),
        ("8. Modelo Financiero y Economías de Escala", "Pág. 9", [
            ("8.1 Comparativa Presupuestaria: Lote Piloto (200 u.) vs. Escala (1.000 u.)", ""),
            ("8.2 Desglose Unitario Puesto en Almacén y Margen de Contribución", "")
        ]),
        ("9. Matriz de Riesgos y Protocolo de Mitigación", "Pág. 10", [
            ("9.1 Riesgos Técnicos, Regulatorios, de Calidad Fabril y Comerciales", "")
        ]),
        ("10. Experiencia de Empaque (Unboxing & Accesorios Oficiales)", "Pág. 11", [
            ("10.1 Contenido del Cofre Rígido Magnético y Presentación", "")
        ])
    ]
    
    t_idx = doc.add_table(rows=len(index_items) + 1, cols=2)
    t_idx.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_idx.rows[0].cells[0].text = "Sección / Capítulo"
    t_idx.rows[0].cells[1].text = "Referencia"
    t_idx.rows[0].cells[0].width = Inches(5.6)
    t_idx.rows[0].cells[1].width = Inches(1.0)
    
    set_cell_background(t_idx.rows[0].cells[0], "0F172A")
    set_cell_background(t_idx.rows[0].cells[1], "0F172A")
    set_cell_margins(t_idx.rows[0].cells[0], 100, 100, 120, 120)
    set_cell_margins(t_idx.rows[0].cells[1], 100, 100, 120, 120)
    p_ih0 = t_idx.rows[0].cells[0].paragraphs[0]
    p_ih0.runs[0].font.bold = True
    p_ih0.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p_ih0.runs[0].font.size = Pt(8.5)
    p_ih1 = t_idx.rows[0].cells[1].paragraphs[0]
    p_ih1.runs[0].font.bold = True
    p_ih1.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    p_ih1.runs[0].font.size = Pt(8.5)
    p_ih1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    for r_idx, (sec_title, page_ref, subitems) in enumerate(index_items, start=1):
        cell_title = t_idx.rows[r_idx].cells[0]
        cell_page = t_idx.rows[r_idx].cells[1]
        cell_title.width = Inches(5.6)
        cell_page.width = Inches(1.0)
        
        bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        set_cell_background(cell_title, bg)
        set_cell_background(cell_page, bg)
        set_cell_margins(cell_title, 70, 70, 100, 100)
        set_cell_margins(cell_page, 70, 70, 100, 100)
        
        p_t = cell_title.paragraphs[0]
        r_t = p_t.add_run(sec_title)
        r_t.font.bold = True
        r_t.font.size = Pt(8.5)
        r_t.font.color.rgb = RGBColor(15, 23, 42)
        
        if subitems:
            for sub_title, _ in subitems:
                p_sub = cell_title.add_paragraph()
                p_sub.paragraph_format.space_before = Pt(1)
                p_sub.paragraph_format.space_after = Pt(1)
                r_sub = p_sub.add_run("   • " + sub_title)
                r_sub.font.size = Pt(7.5)
                r_sub.font.color.rgb = RGBColor(100, 116, 139)
                
        p_p = cell_page.paragraphs[0]
        p_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_p = p_p.add_run(page_ref)
        r_p.font.size = Pt(8)
        r_p.font.color.rgb = RGBColor(100, 116, 139)
        
    set_table_borders(t_idx, color="E2E8F0")

    doc.add_page_break()

    # SECCIÓN 1: IDENTIFICACIÓN DEL PROYECTO, INTEGRANTES Y WEB
    add_header_styled(doc, "1. Identificación del Proyecto, Integrantes y Plataforma Web", level=1)
    
    add_header_styled(doc, "1.1 Integrantes del Proyecto", level=2)
    doc.add_paragraph(
        "A continuación se detallan los integrantes del proyecto LUMOS a cargo de la dirección, "
        "gestión de producto, diseño y desarrollo:"
    )
    
    integrantes = [
        ("1", "Victor Adrian Orue"),
        ("2", "Myriam Beatriz Almiron"),
        ("3", "Karina Valles")
    ]
    
    table_team = doc.add_table(rows=len(integrantes) + 1, cols=2)
    table_team.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_team.rows[0].cells[0].text = "N°"
    table_team.rows[0].cells[1].text = "Nombre y Apellido"
    table_team.rows[0].cells[0].width = Inches(0.8)
    table_team.rows[0].cells[1].width = Inches(5.8)
    
    set_cell_background(table_team.rows[0].cells[0], "0F172A")
    set_cell_background(table_team.rows[0].cells[1], "0F172A")
    set_cell_margins(table_team.rows[0].cells[0], 110, 110, 130, 130)
    set_cell_margins(table_team.rows[0].cells[1], 110, 110, 130, 130)
    
    for c_i in range(2):
        p = table_team.rows[0].cells[c_i].paragraphs[0]
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        p.runs[0].font.size = Pt(9)
        
    for r_i, (num, nombre) in enumerate(integrantes, start=1):
        c_num = table_team.rows[r_i].cells[0]
        c_nom = table_team.rows[r_i].cells[1]
        c_num.width = Inches(0.8)
        c_nom.width = Inches(5.8)
        c_num.text = num
        c_nom.text = nombre
        
        bg = "F8FAFC" if r_i % 2 == 1 else "FFFFFF"
        set_cell_background(c_num, bg)
        set_cell_background(c_nom, bg)
        set_cell_margins(c_num, 90, 90, 110, 110)
        set_cell_margins(c_nom, 90, 90, 110, 110)
        
        p0 = c_num.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.runs[0].font.size = Pt(9)
        p0.runs[0].font.bold = True
        p0.runs[0].font.color.rgb = RGBColor(180, 130, 30)
        
        p1 = c_nom.paragraphs[0]
        p1.runs[0].font.size = Pt(9.5)
        p1.runs[0].font.bold = True
        p1.runs[0].font.color.rgb = RGBColor(15, 23, 42)
        
    set_table_borders(table_team)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # SUBSECCIÓN 1.2: PLATAFORMA WEB OFICIAL
    add_header_styled(doc, "1.2 Plataforma Web Oficial y Recursos Digitales", level=2)
    doc.add_paragraph(
        "El proyecto LUMUS cuenta con una plataforma web interactiva donde se alojan el probador virtual con tecnología de "
        "Realidad Aumentada (LUMUS AR Mirror), el catálogo interactivo de Las Cinco Órdenes y las cotas dimensionales de ergonomía:"
    )
    
    web_fields = [
        ("Sitio Web Oficial / Demo Online", "https://vao-source-code.github.io/lumus/", "Sitio interactivo en producción con catálogo visual, presentación y probador virtual AR."),
        ("Entorno Local de Desarrollo", "index.html / lumos_presentation.html", "Archivos de presentación interactiva locales disponibles en el repositorio."),
        ("Repositorio Oficial de Código", "https://github.com/vao-source-code/lumus", "Repositorio de código abierto con recursos gráficos, estilos y despliegue GitHub Pages.")
    ]
    
    tbl_web = doc.add_table(rows=len(web_fields) + 1, cols=3)
    tbl_web.alignment = WD_TABLE_ALIGNMENT.CENTER
    w_hdrs = ["Canal Digital / Recurso", "Dirección Web / Enlace (URL)", "Descripción del Entorno"]
    w_widths = [Inches(2.0), Inches(2.3), Inches(2.3)]
    
    w_h_cells = tbl_web.rows[0].cells
    for i, h_txt in enumerate(w_hdrs):
        w_h_cells[i].text = h_txt
        w_h_cells[i].width = w_widths[i]
        set_cell_background(w_h_cells[i], "0F172A")
        set_cell_margins(w_h_cells[i], 110, 110, 120, 120)
        p = w_h_cells[i].paragraphs[0]
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        p.runs[0].font.size = Pt(8.5)
        
    for r_i, (c_name, c_url, c_desc) in enumerate(web_fields, start=1):
        r_cells = tbl_web.rows[r_i].cells
        r_cells[0].text = c_name
        r_cells[1].text = c_url
        r_cells[2].text = c_desc
        for c_i in range(3):
            r_cells[c_i].width = w_widths[c_i]
            bg = "F8FAFC" if r_i % 2 == 1 else "FFFFFF"
            set_cell_background(r_cells[c_i], bg)
            set_cell_margins(r_cells[c_i], 80, 80, 100, 100)
            p = r_cells[c_i].paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].font.size = Pt(8)
                if c_i == 0:
                    p.runs[0].font.bold = True
                elif c_i == 1:
                    p.runs[0].font.color.rgb = RGBColor(2, 132, 199)
                    p.runs[0].font.underline = True
                    
    set_table_borders(tbl_web)
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    
    # SECCIÓN 2: RESUMEN EJECUTIVO
    add_header_styled(doc, "2. Resumen Ejecutivo y Tesis de Negocio", level=1)
    
    doc.add_paragraph(
        "LUMUS nace como una respuesta innovadora en el sector de la electrónica de consumo y el equipamiento para artes escénicas. "
        "Frente a un mercado saturado de auriculares Bluetooth sobrios y genéricos, LUMUS fusiona ingeniería acústica Hi-Res con un "
        "sistema de iluminación LED perimetral de estética mística ('música, magia y movimiento'). Su mercado primario abarca academias "
        "de patinaje artístico, danza contemporánea, gimnasia y amantes de shows y festivales inmersivos."
    )
    
    add_callout(
        doc,
        "✦ TESIS DE PRODUCTO: DE DISPOSITIVO PASIVO A ELEMENTO ESCÉNICO",
        "El auricular tradicional aísla al usuario del entorno. LUMUS transforma la audición en una experiencia colectiva y visual. "
        "En su primera fase comercial, integra patrones reactivos autónomos al compás musical. En su etapa de escalabilidad, implementa "
        "sincronización inalámbrica multipunto (BLE Audio / Broadcast) permitiendo a elencos de competición o audiencias en espectáculos "
        "latir con sincronía cromática perfecta comandada desde una consola de luces o dispositivo maestro."
    )
    
    # SECCIÓN 3: ESPECIFICACIONES TÉCNICAS
    add_header_styled(doc, "3. Especificaciones Técnicas y Arquitectura Electroacústica", level=1)
    
    specs_data = [
        ("Subsistema Acústico", "Transductores Acústicos", "Drivers dinámicos de 40 mm con recubrimiento de titanio y cámara acústica sellada."),
        ("Subsistema Acústico", "Respuesta en Frecuencia", "20 Hz – 20.000 Hz, graves contundentes y medios/agudos calibrados para precisión rítmica."),
        ("Subsistema Acústico", "Sensibilidad / Impedancia", "102 ± 3 dB SPL a 1 kHz / 32 Ω nominales."),
        ("Conectividad RF", "SoC Bluetooth", "Bluetooth 5.3 Low Energy (SoC JL / Realtek con firmware optimizado)."),
        ("Conectividad RF", "Códecs de Audio", "AAC, SBC, aptX HD. Latencia ultra-baja (< 45 ms en modo Gaming/Escena)."),
        ("Conectividad Cableada", "Entrada Analógica Hi-Res", "Jack estéreo de 3.5 mm bañado en oro con modo pasivo continuo (sin consumo de batería)."),
        ("Sistema Óptico", "Matriz Lumínica", "Anillos microprismáticos perimetrales con LEDs RGB SMD 0603 de alta eficiencia lumínica."),
        ("Sistema Óptico", "Difusión & Uniformidad", "Difusor de policarbonato opal mate (uniformidad > 92% sin halos ni 'puntos calientes')."),
        ("Energía & Batería", "Capacidad de Batería", "Celda de Polímero de Litio (Li-Po) 3.7V / 800 mAh con circuito de protección PCM."),
        ("Energía & Batería", "Autonomía Nominal", "Hasta 30 horas continuas (sólo audio) / 12 a 15 horas (audio con iluminación activa)."),
        ("Energía & Batería", "Carga & Alimentación", "Puerto USB-C reversible (5V / 1A). Tiempo de recarga completa: ~2.2 horas."),
        ("Ergonomía & Peso", "Estructura Mecánica", "Diadema de aleación ligera con almohadillas de espuma viscoelástica y cuero PU transpirable."),
        ("Ergonomía & Peso", "Peso Total Equilibrado", "265 gramos; centro de gravedad bajo para evitar deslizamientos en saltos y giros de danza.")
    ]
    
    table_specs = doc.add_table(rows=len(specs_data) + 1, cols=3)
    table_specs.alignment = WD_TABLE_ALIGNMENT.CENTER
    sp_headers = ["Subsistema", "Parámetro de Ingeniería", "Especificación Técnica"]
    sp_widths = [Inches(1.7), Inches(1.8), Inches(3.2)]
    
    sp_hdr_cells = table_specs.rows[0].cells
    for i, title in enumerate(sp_headers):
        sp_hdr_cells[i].text = title
        sp_hdr_cells[i].width = sp_widths[i]
        set_cell_background(sp_hdr_cells[i], "0F172A")
        set_cell_margins(sp_hdr_cells[i], 110, 110, 130, 130)
        p = sp_hdr_cells[i].paragraphs[0]
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        p.runs[0].font.size = Pt(8.5)
        
    for r_i, (cat, param, val) in enumerate(specs_data, start=1):
        row_c = table_specs.rows[r_i].cells
        row_c[0].text = cat
        row_c[1].text = param
        row_c[2].text = val
        for c_i in range(3):
            row_c[c_i].width = sp_widths[c_i]
            bg = "F8FAFC" if r_i % 2 == 1 else "FFFFFF"
            set_cell_background(row_c[c_i], bg)
            set_cell_margins(row_c[c_i], 80, 80, 100, 100)
            p = row_c[c_i].paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].font.size = Pt(8)
                if c_i == 0:
                    p.runs[0].font.bold = True
                    
    set_table_borders(table_specs)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # SECCIÓN 4: CONECTIVIDAD DUAL (CON IMÁGENES)
    add_header_styled(doc, "4. Arquitectura de Conectividad Dual y Versatilidad de Uso", level=1)
    
    doc.add_paragraph(
        "Para responder tanto a las exigencias de alta movilidad en pista como a los requerimientos de latencia cero en cabinas de audio y "
        "entrenamientos de alta fidelidad, LUMUS cuenta con una arquitectura de interfaz dual:"
    )
    
    tbl_conn = doc.add_table(rows=1, cols=2)
    tbl_conn.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_w, c_wr = tbl_conn.rows[0].cells
    c_w.width = Inches(3.3)
    c_wr.width = Inches(3.3)
    
    if os.path.exists("lumos_img/wireless.jpg"):
        c_w.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_w.paragraphs[0].add_run().add_picture("lumos_img/wireless.jpg", width=Inches(2.7))
    p_w_t = c_w.add_paragraph()
    r_w1 = p_w_t.add_run("Modo LUMUS Wireless (BT 5.3)\n")
    r_w1.bold = True
    r_w1.font.size = Pt(9.5)
    r_w1.font.color.rgb = RGBColor(15, 23, 42)
    r_w2 = p_w_t.add_run("Libertad de movimiento absoluta sin cables para rutinas deportivas, entrenamientos en pista y escucha cotidiana.")
    r_w2.font.size = Pt(8)
    r_w2.font.color.rgb = RGBColor(71, 85, 105)
    
    if os.path.exists("lumos_img/wired.jpg"):
        c_wr.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_wr.paragraphs[0].add_run().add_picture("lumos_img/wired.jpg", width=Inches(2.7))
    p_wr_t = c_wr.add_paragraph()
    r_wr1 = p_wr_t.add_run("Modo LUMUS Dual Hi-Res (Jack 3.5 mm)\n")
    r_wr1.bold = True
    r_wr1.font.size = Pt(9.5)
    r_wr1.font.color.rgb = RGBColor(15, 23, 42)
    r_wr2 = p_wr_t.add_run("Conexión pasiva cableada directa para consolas de producción, mesas de mezclas y edición sin consumo de batería.")
    r_wr2.font.size = Pt(8)
    r_wr2.font.color.rgb = RGBColor(71, 85, 105)
    
    set_cell_background(c_w, "F8FAFC")
    set_cell_background(c_wr, "F8FAFC")
    set_cell_margins(c_w, 80, 80, 100, 100)
    set_cell_margins(c_wr, 80, 80, 100, 100)
    set_table_borders(tbl_conn, color="E2E8F0")

    doc.add_page_break()

    # SECCIÓN 5: CATÁLOGO DE COLOR (LAS CINCO ÓRDENES)
    add_header_styled(doc, "5. Catálogo Estético: Las Cinco Órdenes de Color", level=1)
    
    doc.add_paragraph(
        "El diseño de LUMUS se articula en torno a cinco identidades cromáticas inspiradas en la magia y los elementos, "
        "cada una con acabados de deposición física de vapor (PVD), electrochapado y difusión LED calibrada a tono:"
    )
    
    orders = [
        ("Orden del Fuego Solar", "Rubí Profundo & Oro Real", "Diseñada para coreografías de alta potencia, ritmos latinos y performances de intensa presencia escénica.", "lumos_img/var_solaris.jpg"),
        ("Orden del Vacío Astral", "Azul Medianoche & Zafiro", "Evoca la inmensidad cósmica y la precisión matemática; predilecta en patinaje clásico y música electrónica.", "lumos_img/var_void.jpg"),
        ("Orden del Bosque Esmeralda", "Jade & Oro Antiguo", "Equilibrio entre mística orgánica y elegancia aristocrática, inspirada en relatos de fantasía clásica.", "lumos_img/var_veridian.jpg"),
        ("Orden de la Nebulosa Rosa", "Perla Rosa & Oro Rosa", "Matiz sutil con destellos irisados; gran afinidad en el segmento juvenil, danza contemporánea y pop visual.", "lumos_img/var_nebula.jpg"),
        ("Orden Signature / Eclipse", "Negro Carbón & Oro Bruñido", "La expresión más sobria y exclusiva de LUMUS; acabado mate táctil con detalles dorados pulidos a espejo.", "lumos_img/var_aurora.jpg")
    ]
    
    for title, subtitle, desc, img_path in orders:
        t_ord = doc.add_table(rows=1, cols=2)
        t_ord.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_i = t_ord.rows[0].cells[0]
        c_t = t_ord.rows[0].cells[1]
        c_i.width = Inches(2.1)
        c_t.width = Inches(4.5)
        
        if os.path.exists(img_path):
            c_i.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            c_i.paragraphs[0].add_run().add_picture(img_path, width=Inches(1.8))
            
        p_ot = c_t.paragraphs[0]
        p_ot.paragraph_format.space_before = Pt(2)
        p_ot.paragraph_format.space_after = Pt(2)
        r1 = p_ot.add_run(f"✦ {title}\n")
        r1.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = RGBColor(15, 23, 42)
        
        r2 = p_ot.add_run(f"Paleta CMF: {subtitle}\n\n")
        r2.bold = True
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(180, 130, 30)
        
        r3 = p_ot.add_run(desc)
        r3.font.size = Pt(9)
        r3.font.color.rgb = RGBColor(71, 85, 105)
        
        set_cell_background(c_i, "F8FAFC")
        set_cell_background(c_t, "F8FAFC")
        set_cell_margins(c_i, 80, 80, 100, 100)
        set_cell_margins(c_t, 80, 80, 100, 100)
        set_table_borders(t_ord, color="E2E8F0")
        
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(4)

    # SECCIÓN 6: EXTENSIONES DE LÍNEA
    add_header_styled(doc, "6. Extensiones de Línea: Glide Fit y Nova Mini", level=1)
    
    doc.add_paragraph(
        "Como parte de la estrategia de expansión de catálogo, se proyectan dos modelos complementarios desarrollados "
        "sobre la misma identidad de marca y lenguaje lumínico:"
    )
    
    tbl_ext = doc.add_table(rows=1, cols=2)
    tbl_ext.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_gl, c_no = tbl_ext.rows[0].cells
    c_gl.width = Inches(3.3)
    c_no.width = Inches(3.3)
    
    if os.path.exists("lumos_img/glide.jpg"):
        c_gl.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_gl.paragraphs[0].add_run().add_picture("lumos_img/glide.jpg", width=Inches(2.6))
    p_gl_t = c_gl.add_paragraph()
    r_g1 = p_gl_t.add_run("LUMUS Glide Fit (Open-Ear Deportivo)\n")
    r_g1.bold = True
    r_g1.font.size = Pt(9.5)
    r_g1.font.color.rgb = RGBColor(15, 23, 42)
    r_g2 = p_gl_t.add_run("Transmisión de conducción de aire / oreja libre. Permite a los deportistas escuchar las instrucciones de sus entrenadores mientras ejecutan sus rutinas con plena conciencia espacial.")
    r_g2.font.size = Pt(8)
    r_g2.font.color.rgb = RGBColor(71, 85, 105)
    
    if os.path.exists("lumos_img/nova.jpg"):
        c_no.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_no.paragraphs[0].add_run().add_picture("lumos_img/nova.jpg", width=Inches(2.6))
    p_no_t = c_no.add_paragraph()
    r_n1 = p_no_t.add_run("LUMUS Nova Mini (Ultra-Portable TWS)\n")
    r_n1.bold = True
    r_n1.font.size = Pt(9.5)
    r_n1.font.color.rgb = RGBColor(15, 23, 42)
    r_n2 = p_no_t.add_run("Formato intrauditivo ultracompacto con estuche joyero iluminado; portabilidad extrema para el día a día y calentamientos previos a la competencia.")
    r_n2.font.size = Pt(8)
    r_n2.font.color.rgb = RGBColor(71, 85, 105)
    
    set_cell_background(c_gl, "F8FAFC")
    set_cell_background(c_no, "F8FAFC")
    set_cell_margins(c_gl, 80, 80, 100, 100)
    set_cell_margins(c_no, 80, 80, 100, 100)
    set_table_borders(tbl_ext, color="E2E8F0")

    doc.add_page_break()

    # SECCIÓN 7: PLAN DE ESCALABILIDAD
    add_header_styled(doc, "7. Plan Integral de Escalabilidad Técnica e Industrial", level=1)
    
    doc.add_paragraph(
        "Para transformar LUMUS de un producto importado a una plataforma tecnológica consolidada con ventajas competitivas "
        "sostenibles, se implementa una hoja de ruta estructurada en tres fases acumulativas:"
    )
    
    phases = [
        ("FASE 1: Validación de Mercado & Adaptación ODM/OEM (Meses 0 a 6)",
         "Objetivo Estratégico", "Lanzamiento ágil con inversión de capital acotada, testeo en academias y cumplimiento de normativas locales.",
         "Plan de Acción de Ingeniería & Operaciones",
         "• Búsqueda y auditoría de 2 fabricantes ODM líderes con plataformas Bluetooth 5.3 existentes.\n"
         "• Personalización visual: serigrafía de logotipo LUMUS, acabados electrochapados PVD y empaque exclusivo de marca.\n"
         "• Producción de lote piloto de 200 a 500 unidades para pruebas con academias de patinaje y danza.\n"
         "• Ensayos regulatorios en Argentina: Obtención de certificación de seguridad eléctrica de baterías de litio (Res. 169/2018) y homologación ENACOM (módulo RF 2.4 GHz)."),
        
        ("FASE 2: Firmware Propietario & Sincronización Escénica Mesh (Meses 6 a 18)",
         "Objetivo Estratégico", "Desarrollo del valor tecnológico diferencial: sincronización lumínica grupal para espectáculos y elencos.",
         "Plan de Acción de Ingeniería & Operaciones",
         "• Desarrollo de firmware propietario para SoC compatible con Bluetooth LE Audio y Auracast Broadcast.\n"
         "• Creación del módulo 'LUMUS Stage Master' (emisor RF / DMX): un solo transmisor comanda el patrón y color de más de 150 auriculares en pista con latencia inferior a 20 ms.\n"
         "• Lanzamiento de la aplicación móvil LUMUS App (iOS/Android): ecualizador paramétrico, perfiles lumínicos por canción y modo de sincronización por compás musical.\n"
         "• Escalabilidad del lote a 1.500 – 3.000 unidades con reducción de costo unitario FOB de más del 20%."),
        
        ("FASE 3: Fabricación Propia, Moldes Propios y Expansión Latam (Meses 18 a 36)",
         "Objetivo Estratégico", "Consolidación de marca de referencia en artes sobre ruedas/hielo y expansión regional en América Latina.",
         "Plan de Acción de Ingeniería & Operaciones",
         "• Fabricación de moldes y matrices de inyección plástica propios para carcasa y diadema con patente de diseño industrial registrada.\n"
         "• Ecosistema de accesorios: estaciones de carga múltiple para academias (hasta 20 auriculares en simultáneo), diademas intercambiables y estuches rígidos reforzados para giras.\n"
         "• Expansión comercial a Brasil, Chile, Colombia y México mediante distribuidores especializados en deportes artísticos y eventos inmersivos.\n"
         "• Producción anual escalada: 10.000+ unidades anuales con control de calidad continuo (AQL 1.0) y soporte técnico oficial local.")
    ]
    
    for f_title, f_obj_label, f_obj_txt, f_tech_label, f_tech_txt in phases:
        add_header_styled(doc, f_title, level=2)
        
        t_phase = doc.add_table(rows=2, cols=2)
        t_phase.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_phase.rows[0].cells[0].text = f_obj_label
        t_phase.rows[0].cells[1].text = f_obj_txt
        t_phase.rows[1].cells[0].text = f_tech_label
        t_phase.rows[1].cells[1].text = f_tech_txt
        
        for r_i in range(2):
            c0 = t_phase.rows[r_i].cells[0]
            c1 = t_phase.rows[r_i].cells[1]
            c0.width = Inches(1.8)
            c1.width = Inches(4.8)
            set_cell_background(c0, "0F172A" if r_i == 0 else "1E293B")
            set_cell_background(c1, "F8FAFC" if r_i == 0 else "FFFFFF")
            set_cell_margins(c0, 90, 90, 110, 110)
            set_cell_margins(c1, 90, 90, 110, 110)
            
            p0 = c0.paragraphs[0]
            p0.runs[0].font.bold = True
            p0.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            p0.runs[0].font.size = Pt(8.5)
            
            p1 = c1.paragraphs[0]
            p1.runs[0].font.size = Pt(8.5)
            p1.runs[0].font.color.rgb = RGBColor(30, 41, 59)
            
        set_table_borders(t_phase, color="CBD5E1")
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # SECCIÓN 8: MODELO FINANCIERO Y PRESUPUESTO ESCALABLE
    add_header_styled(doc, "8. Modelo Financiero y Economías de Escala", level=1)
    
    doc.add_paragraph(
        "A continuación se presenta el análisis comparativo del costo unitario y márgenes comerciales entre el lote piloto (200 unidades) "
        "y el lote industrial escalado (1.000 unidades), mostrando el impacto de la amortización de gastos fijos y la reducción FOB:"
    )
    
    fin_data = [
        ("Costo FOB Unitario Auricular", "USD 32,00", "USD 24,50", "Economía de escala y volumen con el fabricante ODM/OEM."),
        ("Flete Internacional & Seguro (Unitario)", "USD 9,00", "USD 4,80", "Consolidación de carga marítima vs aérea express."),
        ("Aranceles, IVA, Tasas y Despacho Aduanero", "USD 16,00", "USD 12,25", "Estimación arancelaria e impuestos de nacionalización."),
        ("Homologación ENACOM & Seg. Eléctrica (Amortizada)", "USD 5,00", "USD 1,00", "Amortización de gastos de ensayos técnicos fijos."),
        ("Costo Total Unitario Puesto en Almacén", "USD 62,00", "USD 42,55", "Reducción neta del 31.4% en costo por unidad."),
        ("Precio de Venta Sugerido al Público (PVP)", "USD 110,00", "USD 110,00", "Posicionamiento competitivo dentro de la gama media-alta."),
        ("Margen Bruto de Contribución por Unidad", "USD 48,00 (43.6%)", "USD 67,45 (61.3%)", "Ampliación de rentabilidad neta para soporte e I+D."),
        ("Facturación Bruta Proyectada por Lote", "USD 22.000", "USD 110.000", "Escalabilidad sostenida y flujo de fondos positivo.")
    ]
    
    t_fin = doc.add_table(rows=len(fin_data) + 1, cols=4)
    t_fin.alignment = WD_TABLE_ALIGNMENT.CENTER
    f_hdrs = ["Rubro / Concepto Financiero", "Lote Piloto (200 u.)", "Lote Escala (1.000 u.)", "Análisis de Optimización"]
    f_widths = [Inches(2.1), Inches(1.3), Inches(1.3), Inches(1.9)]
    
    f_h_cells = t_fin.rows[0].cells
    for i, h_txt in enumerate(f_hdrs):
        f_h_cells[i].text = h_txt
        f_h_cells[i].width = f_widths[i]
        set_cell_background(f_h_cells[i], "0F172A")
        set_cell_margins(f_h_cells[i], 110, 110, 120, 120)
        p = f_h_cells[i].paragraphs[0]
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        p.runs[0].font.size = Pt(8.5)
        
    for r_i, (c_name, p_200, p_1000, c_note) in enumerate(fin_data, start=1):
        r_cells = t_fin.rows[r_i].cells
        r_cells[0].text = c_name
        r_cells[1].text = p_200
        r_cells[2].text = p_1000
        r_cells[3].text = c_note
        for c_i in range(4):
            r_cells[c_i].width = f_widths[c_i]
            bg = "F8FAFC" if r_i % 2 == 1 else "FFFFFF"
            set_cell_background(r_cells[c_i], bg)
            set_cell_margins(r_cells[c_i], 80, 80, 100, 100)
            p = r_cells[c_i].paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].font.size = Pt(8)
                if c_i in (1, 2) and ("Total" in c_name or "Margen" in c_name or "Facturación" in c_name):
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = RGBColor(180, 130, 30)
                    
    set_table_borders(t_fin)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # SECCIÓN 9: MATRIZ DE RIESGOS Y MITIGACIÓN
    add_header_styled(doc, "9. Matriz de Riesgos y Protocolo de Mitigación", level=1)
    
    risks_data = [
        ("Técnico", "Latencia o interferencia en la sincronización de luces grupal.", "Alta", "Separar el lanzamiento: la Fase 1 ofrece modos autónomos integrados y la sincronización grupal se lanza en Fase 2 tras pruebas de estrés con hardware de prueba dedicado."),
        ("Regulatorio", "Demoras en ensayos aduaneros, ENACOM o certificación eléctrica.", "Media", "Contratar despachante aduanero especialista en telecomunicaciones e iniciar gestiones documentales anticipadas con las muestras de preembarque."),
        ("Calidad / QC", "Desviaciones cromáticas o tasa de fallas en batería/LEDs.", "Alta", "Protocolo de inspección en fábrica origen bajo estándar AQL 1.0 para defectos críticos y exigencia de 3% de unidades adicionales en concepto de spare parts."),
        ("Comercial", "Curva de adopción más lenta en el público general.", "Media", "Enfoque inicial B2B en nichos de alta tracción (patinaje artístico, escuelas de danza y acrobacia) que actúan como embajadores orgánicos de la marca.")
    ]
    
    t_risk = doc.add_table(rows=len(risks_data) + 1, cols=4)
    t_risk.alignment = WD_TABLE_ALIGNMENT.CENTER
    r_hdrs = ["Dimensión", "Riesgo Evaluado", "Nivel", "Estrategia Preventiva & Mitigación"]
    r_widths = [Inches(1.1), Inches(1.8), Inches(0.9), Inches(2.8)]
    
    r_h_cells = t_risk.rows[0].cells
    for i, h_txt in enumerate(r_hdrs):
        r_h_cells[i].text = h_txt
        r_h_cells[i].width = r_widths[i]
        set_cell_background(r_h_cells[i], "0F172A")
        set_cell_margins(r_h_cells[i], 110, 110, 120, 120)
        p = r_h_cells[i].paragraphs[0]
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        p.runs[0].font.size = Pt(8.5)
        
    for r_i, (nat, r_name, imp, mit) in enumerate(risks_data, start=1):
        r_cells = t_risk.rows[r_i].cells
        r_cells[0].text = nat
        r_cells[1].text = r_name
        r_cells[2].text = imp
        r_cells[3].text = mit
        for c_i in range(4):
            r_cells[c_i].width = r_widths[c_i]
            bg = "F8FAFC" if r_i % 2 == 1 else "FFFFFF"
            set_cell_background(r_cells[c_i], bg)
            set_cell_margins(r_cells[c_i], 80, 80, 100, 100)
            p = r_cells[c_i].paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].font.size = Pt(8)
                if c_i == 2:
                    p.runs[0].font.bold = True
                    if imp == "Alta":
                        p.runs[0].font.color.rgb = RGBColor(190, 24, 93)
                        
    set_table_borders(t_risk)
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # SECCIÓN 10: EXPERIENCIA DE PACKAGING (CON IMAGEN)
    add_header_styled(doc, "10. Experiencia de Empaque (Unboxing & Accesorios)", level=1)
    
    doc.add_paragraph(
        "El packaging de LUMUS no es sólo un contenedor de transporte, sino un elemento central del ritual de experiencia. "
        "Se trata de un cofre rígido de apertura magnética con acabados táctiles soft-touch y estampados en oro reflectivo:"
    )
    
    tbl_pack = doc.add_table(rows=1, cols=2)
    tbl_pack.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_p_i = tbl_pack.rows[0].cells[0]
    c_p_t = tbl_pack.rows[0].cells[1]
    c_p_i.width = Inches(2.7)
    c_p_t.width = Inches(3.9)
    
    if os.path.exists("lumos_img/packaging.jpg"):
        c_p_i.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_p_i.paragraphs[0].add_run().add_picture("lumos_img/packaging.jpg", width=Inches(2.5))
        
    p_pack_t = c_p_t.paragraphs[0]
    p_pack_t.paragraph_format.space_before = Pt(2)
    r_pt = p_pack_t.add_run("Contenido Oficial del Cofre LUMUS:\n\n")
    r_pt.bold = True
    r_pt.font.size = Pt(9.5)
    r_pt.font.color.rgb = RGBColor(15, 23, 42)
    
    r_pb = p_pack_t.add_run(
        "1. Auriculares Inteligentes LUMUS (edición de color elegida).\n"
        "2. Estuche rígido semiesférico de viaje contra impactos.\n"
        "3. Cable de carga USB-C trenzado de alta durabilidad con terminales dorados.\n"
        "4. Cable de audio estéreo Jack 3.5 mm bañado en oro para modo cableado Hi-Res.\n"
        "5. Tarjeta de autenticidad y activación con acceso a contenido digital exclusivo y manual oficial."
    )
    r_pb.font.size = Pt(8.5)
    r_pb.font.color.rgb = RGBColor(71, 85, 105)
    
    set_cell_background(c_p_i, "F8FAFC")
    set_cell_background(c_p_t, "F8FAFC")
    set_cell_margins(c_p_i, 80, 80, 100, 100)
    set_cell_margins(c_p_t, 80, 80, 100, 100)
    set_table_borders(tbl_pack, color="E2E8F0")

    output_filename = "LUMOS_Documento_Tecnico_y_Escalabilidad.docx"
    doc.save(output_filename)
    print(f"Documento regenerado con éxito: {output_filename}")

if __name__ == "__main__":
    build_lumus_doc()
