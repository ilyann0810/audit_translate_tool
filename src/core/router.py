import os
from processors.docx_processor import translate_docx_universal
from processors.pptx_processor import translate_pptx_universal
from processors.xlsx_processor import translate_xlsx_universal
from processors.pdf_processor import translate_pdf_universal # NOUVEAU

def process_translation(input_path, output_path, glossary_path=None, progress_callback=None, direction="Allemand ➔ Français"):
    ext = os.path.splitext(input_path)[1].lower()
    
    if ext == '.docx':
        return translate_docx_universal(input_path, output_path, glossary_path, progress_callback, direction)
    elif ext == '.pptx':
        return translate_pptx_universal(input_path, output_path, glossary_path, progress_callback, direction)
    elif ext in ['.xlsx', '.xlsm']:
        return translate_xlsx_universal(input_path, output_path, glossary_path, progress_callback, direction)
    elif ext == '.pdf':
        return translate_pdf_universal(input_path, output_path, glossary_path, progress_callback, direction) 
    else:
        raise ValueError(f"Le format de fichier '{ext}' n'est pas pris en charge.")