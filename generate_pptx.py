#!/usr/bin/env python3
"""
generate_pptx.py - Genera PPTX vertical (portrait) con texto editable
"""
import json
import sys
import base64
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def hex_to_rgb(hex_color):
    """Convierte hex a RGB"""
    if not hex_color:
        return RGBColor(0, 0, 0)
    hex_color = str(hex_color).lstrip('#')
    try:
        return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
    except:
        return RGBColor(0, 0, 0)

def create_vertical_pptx(slides_data):
    """Crea presentación PPTX vertical (portrait) A4"""
    
    # Crear presentación
    prs = Presentation()
    
    # Configurar slide master para PORTRAIT (vertical) A4
    # A4 vertical: 7.5 x 10.8 pulgadas (190 x 274 mm)
    prs.slide_width = Inches(7.5)
    prs.slide_height = Inches(10.8)
    
    print(f"Creando {len(slides_data)} slides en orientación vertical A4...")
    
    for idx, slide_data in enumerate(slides_data):
        print(f"Procesando slide {idx + 1}: {slide_data.get('name', 'Sin nombre')}")
        
        # Agregar slide en blanco
        blank_slide_layout = prs.slide_layouts[6]  # Layout en blanco
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Fondo blanco
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        # Procesar elementos del slide
        elements = slide_data.get('elements', [])
        for element in elements:
            add_element_to_slide(slide, element, slide_data)
    
    return prs

def add_element_to_slide(slide, element, slide_data):
    """Agrega un elemento (texto, shape, imagen) al slide"""
    
    elem_type = element.get('type', '').lower()
    
    # Convertir coordenadas de píxeles Figma a pulgadas PowerPoint
    # Slide A4 vertical: 7.5 x 10.8 pulgadas
    frame_width = slide_data.get('width', 960)
    frame_height = slide_data.get('height', 1280)
    
    scale_x = 7.5 / frame_width
    scale_y = 10.8 / frame_height
    
    x = Inches(element.get('x', 0) * scale_x)
    y = Inches(element.get('y', 0) * scale_y)
    width = Inches(element.get('width', 100) * scale_x)
    height = Inches(element.get('height', 100) * scale_y)
    
    try:
        if elem_type == 'text':
            add_text(slide, element, x, y, width, height, scale_x, scale_y)
        elif elem_type == 'rectangle':
            add_rectangle(slide, element, x, y, width, height)
        elif elem_type == 'circle':
            add_circle(slide, element, x, y, width, height)
        elif elem_type == 'image':
            add_image(slide, element, x, y, width, height)
        elif elem_type == 'group':
            for child in element.get('children', []):
                add_element_to_slide(slide, child, slide_data)
    except Exception as e:
        print(f"  Advertencia - {elem_type}: {str(e)}")

def add_text(slide, element, x, y, w, h, scale_x, scale_y):
    """Agrega texto editable nativo"""
    
    text_box = slide.shapes.add_textbox(x, y, w, h)
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_bottom = Inches(0.05)
    text_frame.margin_top = Inches(0.05)
    text_frame.margin_left = Inches(0.05)
    text_frame.margin_right = Inches(0.05)
    
    # Agregar texto
    p = text_frame.paragraphs[0]
    p.text = element.get('content', '')
    
    # Formato de texto
    font_size = element.get('fontSize', 12)
    font_family = element.get('fontFamily', 'Calibri')
    
    if p.runs:
        run = p.runs[0]
    else:
        run = p.add_run()
    
    run.font.size = Pt(max(8, font_size * scale_x))
    run.font.name = font_family
    run.font.bold = element.get('fontWeight', 400) >= 700
    
    # Color
    color_hex = element.get('color')
    if color_hex:
        try:
            run.font.color.rgb = hex_to_rgb(color_hex)
        except:
            pass
    
    # Alineación
    align_map = {
        'LEFT': PP_ALIGN.LEFT,
        'CENTER': PP_ALIGN.CENTER,
        'RIGHT': PP_ALIGN.RIGHT
    }
    p.alignment = align_map.get(element.get('textAlign', 'LEFT'), PP_ALIGN.LEFT)

def add_rectangle(slide, element, x, y, w, h):
    """Agrega rectángulo"""
    shape = slide.shapes.add_shape(1, x, y, w, h)  # 1 = Rectangle
    
    # Fill
    fill = shape.fill
    fill_color = element.get('fill')
    if fill_color:
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(fill_color)
    
    # Line
    line = shape.line
    stroke_color = element.get('stroke')
    if stroke_color:
        line.color.rgb = hex_to_rgb(stroke_color)
        line.width = Pt(1)

def add_circle(slide, element, x, y, w, h):
    """Agrega círculo/elipse"""
    shape = slide.shapes.add_shape(3, x, y, w, h)  # 3 = Oval
    
    # Fill
    fill = shape.fill
    fill_color = element.get('fill')
    if fill_color:
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(fill_color)

def add_image(slide, element, x, y, w, h):
    """Agrega imagen desde base64"""
    image_base64 = element.get('imageBase64')
    if image_base64:
        try:
            image_bytes = base64.b64decode(image_base64)
            image_stream = io.BytesIO(image_bytes)
            slide.shapes.add_picture(image_stream, x, y, width=w, height=h)
        except Exception as e:
            print(f"  Advertencia - Imagen: {str(e)}")

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Uso: python3 generate_pptx.py <input_json> <output_pptx>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    try:
        with open(input_file, 'r') as f:
            data = json.load(f)
        
        slides = data.get('slides', [])
        prs = create_vertical_pptx(slides)
        prs.save(output_file)
        print(f"✓ PPTX guardado: {output_file}")
        
    except Exception as e:
        print(f"ERROR: {str(e)}", file=sys.stderr)
        sys.exit(1)
