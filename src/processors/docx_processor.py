import os
import shutil
import json
import re
import time
import docx
from copy import deepcopy
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from concurrent.futures import ThreadPoolExecutor, as_completed

# Imports de notre nouvelle architecture centrale
from core.translation_api import call_gemini_batch, call_gemini_fallback, build_prompt
from core.text_utils import load_glossary_df, get_matching_glossary_rules

def get_run_highlight_val(r):
    vals = r._r.xpath('.//w:highlight/@w:val')
    if vals and vals[0] != 'none':
        return f"hl:{vals[0]}"
    
    shd_fills = r._r.xpath('.//w:shd/@w:fill')
    if shd_fills and shd_fills[0] not in ['none', 'auto']:
        return f"shd:{shd_fills[0]}"
        
    if r.font and r.font.highlight_color is not None:
        try:
            val = str(r.font.highlight_color)
            if val != '0' and val != 'None':
                return f"hl:{val}"
        except Exception:
            pass
    return None

def set_run_highlight(run, color_val):
    rPr = run._r.get_or_add_rPr()
    for old in rPr.xpath('.//w:highlight | .//w:shd'):
        rPr.remove(old)
        
    if not color_val:
        return
        
    if color_val.startswith("hl:"):
        val = color_val.split("hl:")[1]
        if val.isdigit():
            try:
                run.font.highlight_color = WD_COLOR_INDEX(int(val))
                return
            except Exception:
                pass
        hl_elem = OxmlElement('w:highlight')
        hl_elem.set(qn('w:val'), val)
        rPr.append(hl_elem)
    elif color_val.startswith("shd:"):
        fill = color_val.split("shd:")[1]
        shd_elem = OxmlElement('w:shd')
        shd_elem.set(qn('w:val'), 'clear')
        shd_elem.set(qn('w:color'), 'auto')
        shd_elem.set(qn('w:fill'), fill)
        rPr.append(shd_elem)

def format_run_text(r):
    txt = r.text
    if not txt: return ""

    match = re.match(r'^(\s*)(.*?)(\s*)$', txt, flags=re.DOTALL)
    if not match: return txt

    lead_ws, core_txt, trail_ws = match.groups()
    if not core_txt: return txt

    is_bold = bool(r.bold or (r.font and r.font.bold))
    is_italic = bool(r.italic or (r.font and r.font.italic))
    is_strike = bool(r.font and r.font.strike)
    is_underline = bool(r.underline or (r.font and r.font.underline))
    hl_color = get_run_highlight_val(r)

    if hl_color: core_txt = f'<hl c="{hl_color}">{core_txt}</hl>'
    if is_strike: core_txt = f"<s>{core_txt}</s>"
    if is_underline: core_txt = f"<u>{core_txt}</u>"
    if is_bold: core_txt = f"<b>{core_txt}</b>"
    if is_italic: core_txt = f"<i>{core_txt}</i>"
        
    return f"{lead_ws}{core_txt}{trail_ws}"

def get_paragraph_tagged_text(p):
    p_elem = p._element
    hyperlinks = p_elem.xpath('.//w:hyperlink')
    
    if hyperlinks:
        tagged_str = ""
        for child in p_elem:
            tag = child.tag.split('}')[-1]
            if tag == 'r':
                r = docx.text.run.Run(child, p)
                tagged_str += format_run_text(r)
            elif tag == 'hyperlink':
                h_idx = hyperlinks.index(child)
                h_runs = child.xpath('.//w:r')
                h_text = "".join([format_run_text(docx.text.run.Run(hr, p)) for hr in h_runs])
                if h_text.strip():
                    tagged_str += f'<link id="{h_idx}">{h_text}</link>'
            else:
                sub_runs = child.xpath('.//w:r')
                for sr in sub_runs:
                    tagged_str += format_run_text(docx.text.run.Run(sr, p))
        res = tagged_str if tagged_str.strip() else p.text
    else:
        r_elems = p_elem.xpath('.//w:r')
        tagged_parts = [format_run_text(docx.text.run.Run(r, p)) for r in r_elems]
        res = "".join(tagged_parts) if tagged_parts else p.text
    
    return res

def sanitize_translation_highlights(orig_tagged_text, translated_text):
    orig_has_hl = bool(re.search(r'<hl\s+c="[^"]+">', orig_tagged_text))
    if not orig_has_hl:
        return re.sub(r'</?hl(?:\s+c="[^"]+")?>', '', translated_text)
    return translated_text

def apply_tagged_text_to_paragraph(p, tagged_text):
    p_elem = p._element

    rPr_template = None
    for r in p.runs:
        if r._r.rPr is not None:
            rPr_template = deepcopy(r._r.rPr)
            for tag in ['w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:strike', 'w:u', 'w:highlight', 'w:shd']:
                for el in rPr_template.xpath(f'.//{tag}'):
                    el.getparent().remove(el)
            break

    hyperlink_elems = [deepcopy(h) for h in p_elem.xpath('.//w:hyperlink')]

    for child in list(p_elem):
        tag = child.tag.split('}')[-1]
        if tag in ['r', 'hyperlink']:
            p_elem.remove(child)

    link_pattern = r'(<link id="\d+">.*?</link>)'
    parts = re.split(link_pattern, tagged_text, flags=re.DOTALL)

    for part in parts:
        if not part: continue

        link_match = re.match(r'<link id="(\d+)">(.*?)</link>', part, flags=re.DOTALL)
        if link_match:
            h_idx = int(link_match.group(1))
            h_content = link_match.group(2)

            if h_idx < len(hyperlink_elems):
                h_elem = deepcopy(hyperlink_elems[h_idx])
                clean_link_text = re.sub(r'</?(?:b|i|hl|s|u)(?:\s+c="[^"]+")?>', '', h_content)
                
                h_runs = h_elem.xpath('.//w:r')
                if h_runs:
                    docx.text.run.Run(h_runs[0], p).text = clean_link_text
                    for hr in h_runs[1:]:
                        docx.text.run.Run(hr, p).text = ""
                p_elem.append(h_elem)
        else:
            tokens = re.split(r'(</?(?:b|i|s|u)>|<hl c="[^"]+">|</hl>)', part)
            curr_bold, curr_italic, curr_strike, curr_underline, curr_hl_color = False, False, False, False, None

            for token in tokens:
                if not token: continue
                if token == '<b>': curr_bold = True
                elif token == '</b>': curr_bold = False
                elif token == '<i>': curr_italic = True
                elif token == '</i>': curr_italic = False
                elif token == '<s>': curr_strike = True
                elif token == '</s>': curr_strike = False
                elif token == '<u>': curr_underline = True
                elif token == '</u>': curr_underline = False
                elif token.startswith('<hl c="'):
                    m_hl = re.match(r'<hl c="([^"]+)">', token)
                    if m_hl: curr_hl_color = m_hl.group(1)
                elif token == '</hl>':
                    curr_hl_color = None
                else:
                    run = p.add_run(token)
                    if rPr_template is not None:
                        if run._r.rPr is not None:
                            run._r.remove(run._r.rPr)
                        run._r.insert(0, deepcopy(rPr_template))

                    run.bold = curr_bold
                    run.italic = curr_italic
                    run.font.strike = curr_strike
                    run.font.underline = curr_underline
                    set_run_highlight(run, curr_hl_color)

def extract_all_paragraphs_universal(doc_obj):
    seen_ids = set()
    all_paragraphs = []
    
    def add_p(p_elem, parent):
        p_id = id(p_elem)
        if p_id not in seen_ids:
            seen_ids.add(p_id)
            p = Paragraph(p_elem, parent)
            if p.text.strip():
                all_paragraphs.append(p)

    for p_elem in doc_obj.element.xpath('.//w:p'):
        add_p(p_elem, doc_obj)

    for section in doc_obj.sections:
        for container in [section.header, section.first_page_header, section.even_page_header,
                          section.footer, section.first_page_footer, section.even_page_footer]:
            if container:
                for p_elem in container._element.xpath('.//w:p'):
                    add_p(p_elem, container)

    for part in doc_obj.part.related_parts.values():
        if any(k in part.content_type for k in ["footnote", "endnote", "comment"]):
            if hasattr(part, 'element'):
                for p_elem in part.element.xpath('.//w:p'):
                    add_p(p_elem, doc_obj)

    return all_paragraphs

def _process_single_batch(args):
    """Traite un lot et renvoie les traductions ET les tokens consommés."""
    batch_idx, batch_paragraphs, glossary_snippet, direction = args
    orig_tagged_texts = [get_paragraph_tagged_text(p) for p in batch_paragraphs]
    
    
    # Ajout des règles de style spécifiques à Word par-dessus le glossaire
    extra_rules = (
        "\n\nCONSIGNES STRICTES DE SURLIGNAGE ET STYLE :\n"
        "1. PRIORITÉ ABSOLUE AU GLOSSAIRE : Applique strictement les traductions du glossaire, y compris pour les sigles. Si un sigle n'est pas dans le glossaire, laisse-le tel quel.\n"
        "2. BALISES <hl c=\"...\"> : Conserve STRICTEMENT l'attribut c=\"...\" intact s'il est présent.\n"
        "3. MOTS ISOLÉS SURLIGNÉS : Si un mot est surligné (<hl c=\"...\">Wort</hl>), entoure son équivalent traduit avec la même balise.\n"
        "4. CONSERVE TOUTES LES BALISES STRUCTURALES (<b>, <i>, <u>, <s>, <link>)."
    )
    combined_context = (glossary_snippet or "") + extra_rules
    combined_context = (glossary_snippet or "") + extra_rules

    # Construction du prompt avec la nouvelle fonction centrale
    prompt = build_prompt(orig_tagged_texts, combined_context, direction=direction, is_batch=True)
    
    batch_in_tokens = 0
    batch_out_tokens = 0

    translations, in_tok, out_tok = call_gemini_batch(prompt, expected_count=len(batch_paragraphs))
    batch_in_tokens += in_tok
    batch_out_tokens += out_tok
    
    if translations:
        sanitized_texts = [sanitize_translation_highlights(orig, trans) for orig, trans in zip(orig_tagged_texts, translations)]
        return batch_idx, batch_paragraphs, sanitized_texts, batch_in_tokens, batch_out_tokens

    # Mode secours unitaire
    print(f"⚠️ Lot {batch_idx + 1} : Bascule en mode de secours unitaire...")
    fallback_translated = []
    
    for p, orig_txt in zip(batch_paragraphs, orig_tagged_texts):
        fb_prompt = build_prompt(orig_txt, combined_context, direction=direction, is_batch=False)
        trans_txt, fb_in, fb_out = call_gemini_fallback(fb_prompt)
        
        batch_in_tokens += fb_in
        batch_out_tokens += fb_out
        
        if trans_txt:
            fallback_translated.append(sanitize_translation_highlights(orig_txt, trans_txt))
        else:
            fallback_translated.append(orig_txt)
            
    return batch_idx, batch_paragraphs, fallback_translated, batch_in_tokens, batch_out_tokens

def translate_docx_universal(input_path: str, output_path: str, glossary_path: str = None, progress_callback=None, direction="Allemand ➔ Français"):
    """Fonction principale appelée par le routeur Streamlit."""
    start_time = time.time()
    total_input_tokens = 0
    total_output_tokens = 0

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    tmp_in = f"/tmp/in_{os.path.basename(input_path)}"
    tmp_out = f"/tmp/out_{os.path.basename(input_path)}"
    
    shutil.copy(input_path, tmp_in)
    doc = docx.Document(tmp_in)
    glossary_df = load_glossary_df(glossary_path)

    paragraphs_to_translate = extract_all_paragraphs_universal(doc)
    total_p = len(paragraphs_to_translate)

    batch_size, max_workers = (20, 6) if total_p > 150 else ((20, 3) if total_p > 20 else (10, 1))

    batch_args_list = [] 
    for i in range(0, total_p, batch_size):
        batch = paragraphs_to_translate[i:i + batch_size]
        batch_text = " ".join([p.text for p in batch])
        glossary_snippet = get_matching_glossary_rules(batch_text, glossary_df,direction=direction)
        
        # On passe la direction et le glossaire aux workers
        batch_args_list.append((i//batch_size, batch, glossary_snippet, direction))

    # Sécurité anti-crash si le document est vide ou sans texte
    if not batch_args_list:
        doc.save(tmp_out)
        shutil.copy(tmp_out, output_path)
        os.remove(tmp_in)
        os.remove(tmp_out)
        return {
            "status": "success",
            "output_path": output_path,
            "time_seconds": round(time.time() - start_time, 2),
            "input_tokens": 0,
            "output_tokens": 0,
            "total_cost_usd": 0.0
        }

    print(f"⚡ Traitement global DOCX de {total_p} éléments...")
    total_batches = len(batch_args_list)
    completed_batches = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for future in as_completed([executor.submit(_process_single_batch, arg) for arg in batch_args_list]):
            idx, batch, translated_texts, batch_in, batch_out = future.result()
            
            total_input_tokens += batch_in
            total_output_tokens += batch_out
            
            for p, t_text in zip(batch, translated_texts):
                apply_tagged_text_to_paragraph(p, t_text)
            
            completed_batches += 1
            if progress_callback:
                progress_callback(completed_batches, total_batches)

    doc.save(tmp_out)
    shutil.copy(tmp_out, output_path)
    os.remove(tmp_in)
    os.remove(tmp_out)

    execution_time = time.time() - start_time
    cost_input = (total_input_tokens / 1_000_000) * 0.075
    cost_output = (total_output_tokens / 1_000_000) * 0.30
    total_cost = cost_input + cost_output

    return {
        "status": "success",
        "output_path": output_path,
        "time_seconds": round(execution_time, 2),
        "input_tokens": total_input_tokens,
        "output_tokens": total_output_tokens,
        "total_cost_usd": total_cost
    }