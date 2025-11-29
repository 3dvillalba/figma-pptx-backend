#!/usr/bin/env python3
"""
generate_pptx.py - Genera PPTX vertical (portrait) A4 con texto editable
"""
import json
import sys
import base64
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

def hex_to_rgb(hex_color):
    """Convierte hex a RGB"""
    if not hex_color:
        return RGBColor(0, 0, 0)
    hex_color = str(hex_color).lstrip('#')
    try:
        return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
    except:
        return RGBColor(0, 0, 0)

def set_slide_layout_portrait(prs):
    """Configura presentación en orientación PORTRAIT (vertical)"""
    # A4 vertical: 7.5 x 10.8 pulgadas
    prs.slide_width = Inches(7.5)
    prs.slide_height = Inches(10.8)
    
    # Forzar orientación en XML
    try:
        sld_size = prs.slide_master.part.rels.get_or_add_rel_to_part(prs.slide_master.part)
        # Acceder al elemento de tamaño de slide
        prs_elem = prs.core_properties.element
    except:
        pass

def create_vertical_pptx(slides_data):
    """Crea presentación PPTX vertical (portrait) A4"""
    
    prs = Presentation()
    set_slide_layout_portrait(prs)
    
    print(f"Creando {len(slides_data)} slides en orientación vertical A4...")
    
    for idx, slide_data in enumerate(slides_data):
        print(f"Procesando slide {idx + 1}: {slide_data.get('name', 'Sin nombre')}")
        
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        elements = slide_data.get('elements', [])
        for element in elements:
            add_element_to_slide(slide, element, slide_data)
    
    return prs

def add_element_to_slide(slide, element, slide_data):
    """Agrega un elemento al slide (imagen completa del frame)"""
    
    elem_type = element.get('type', '').lower()
    
    # Las imágenes ocupan el slide completo
    if elem_type == 'image':
        add_image(slide, element)

def add_image(slide, element):
    """Agrega imagen que ocupa todo el slide"""
    image_base64 = element.get('imageBase64')
    if image_base64:
        try:
            # Decodificar base64
            image_bytes = base64.b64decode(image_base64)
            image_stream = io.BytesIO(image_bytes)
            
            # Agregar imagen a todo el ancho del slide
            # Slide A4 vertical: 7.5 x 10.8 pulgadas
            slide.shapes.add_picture(
                image_stream, 
                left=Inches(0), 
                top=Inches(0), 
                width=Inches(7.5),
                height=Inches(10.8)
            )
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
        print(f"✓ PPTX guardado (VERTICAL A4): {output_file}")
        
    except Exception as e:
        print(f"ERROR: {str(e)}", file=sys.stderr)
        sys.exit(1)
