import os
import shutil
import json
import re
import time
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Pt
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from concurrent.futures import ThreadPoolExecutor, as_completed

# Imports de l'architecture centrale
from core.translation_api import call_gemini_batch, call_gemini_fallback, build_prompt
from core.text_utils import load_glossary_df, get_matching_glossary_rules

def check_manual_override(orig_txt):
    """Détecte souplement la phrase allemande pour imposer le texte court du Comex."""
    clean = re.sub(r'<[^>]+>', '', orig_txt).lower()
    clean = re.sub(r'\s+', ' ', clean).strip()
    if "verteilung" in clean and "prüfungsergebnisse" in clean:
        return "Le rapport d'audit sera également envoyé aux membres du Comex"
    return None

def is_untranslatable_text(clean_text: str) -> bool:
    """Détecte si un texte est composé uniquement de chiffres, dates ou symboles."""
    cleaned = clean_text.strip()
    if not cleaned:
        return True
    return not bool(re.search(r'[a-zA-ZÀ-ÿ]', cleaned))

def is_already_in_target_language(clean_text: str, direction: str) -> bool:
    """
    Évite d'inverser des titres 100% déjà dans la langue cible,
    tout en envoyant à la traduction les titres hybrides contenant du français.
    """
    t_lower = clean_text.lower()
    if direction == "Français ➔ Allemand":
        # Marqueurs français signalant qu'il reste du texte à traduire
        french_indicators = [
            r'\b(avec|sans|pour|dans|sur|par|marge|négative|pertes|magasin|semaine|hebdo|hebdomadaire|'
            r'chiffres|clés|non\s+conforme|du|des|de|le|la|les|un|une|et|en|selon|nouveau|ventilation|'
            r'poisson|sec|fruits|légumes|frais|surgelés|viandes|boissons|attention|prochaines)\b'
        ]
        if any(re.search(pat, t_lower) for pat in french_indicators):
            return False  # Contient du français : doit obligatoirement être traduit

        german_markers = [
            "wochenkennzahlen", "wochenperformance", "abschriften",
            "kassenausgang", "eingangskoffer", "sammelstelle"
        ]
        if any(m in t_lower for m in german_markers):
            return True  # 100% allemand : conservé tel quel

    elif direction == "Allemand ➔ Français":
        german_indicators = [
            r'\b(und|oder|mit|ohne|für|in|an|auf|aus|bei|nach|von|zu|vor|über|unter|nicht|konform|'
            r'filiale|filialen|geschäft|kunden|einkauf|supermärkte|überblick|wochenperformance|'
            r'abschriften|verwertung|gesamt|trockensortiment|frische|warengruppe)\b'
        ]
        if any(re.search(pat, t_lower) for pat in german_indicators):
            return False

        french_markers = [
            "indicateurs hebdomadaires", "chiffres clés",
            "point de collecte", "sas d'entrée"
        ]
        if any(m in t_lower for m in french_markers):
            return True

    return False

def fix_french_grammar(text: str) -> str:
    if not text:
        return text
    text = re.sub(r'\bauprès\s+de\s*(<[^>]+>)?\s*les\b', r'auprès des \1', text, flags=re.IGNORECASE)
    text = re.sub(r'\bauprès\s+de\s*(<[^>]+>)?\s*le\b', r'auprès du \1', text, flags=re.IGNORECASE)
    text = re.sub(r'\bà\s*(<[^>]+>)?\s*les\b', r'aux \1', text, flags=re.IGNORECASE)
    text = re.sub(r'\bà\s*(<[^>]+>)?\s*le\b', r'au \1', text, flags=re.IGNORECASE)
    text = re.sub(r'\bde\s*(<[^>]+>)?\s*les\b', r'des \1', text, flags=re.IGNORECASE)
    text = re.sub(r'\bde\s*(<[^>]+>)?\s*le\b', r'du \1', text, flags=re.IGNORECASE)
    return re.sub(r'  +', ' ', text)

def sanitize_pptx_translation(orig_text, trans_text, direction="Allemand ➔ Français"):
    if not trans_text or not trans_text.strip():
        return orig_text
    bad_phrases = [
        "pourriez-vous", "veuillez fournir", "je suis prêt",
        "texte à traduire", "n'hésitez pas", "voici la traduction",
        "tableau json", "données à traduire"
    ]
    trans_lower = trans_text.lower()
    for bp in bad_phrases:
        if bp in trans_lower and bp not in orig_text.lower():
            return re.sub(r'</?r[^>]*>', '', orig_text)

    if re.search(r'\bdv\.\b', orig_text):
        trans_text = re.sub(r'\bdavon\b', 'dv.', trans_text, flags=re.IGNORECASE)

    # Protection contre le retour intempestif d'Indicateurs hebdomadaires
    if direction == "Français ➔ Allemand":
        if re.search(r'\bWochenkennzahlen\b', orig_text, re.IGNORECASE):
            trans_text = re.sub(r'\bIndicateurs hebdomadaires\b', 'Wochenkennzahlen', trans_text, flags=re.IGNORECASE)

    if direction == "Allemand ➔ Français":
        return fix_french_grammar(trans_text)
    return trans_text

# --- GESTION DU XML ET DES PROPRIÉTÉS FINES ---

def add_br(p):
    """Ajoute un saut forcé XML <a:br/> conforme au schéma DrawingML."""
    try:
        br = parse_xml(f'<a:br {nsdecls("a")}/>')
        endParaRPr = p._p.find(qn('a:endParaRPr'))
        if endParaRPr is not None:
            endParaRPr.addprevious(br)
        else:
            p._p.append(br)
    except Exception:
        pass

def lock_shape_geometry(tf):
    """Active le word wrap et verrouille la géométrie de la boîte."""
    try:
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_top = Pt(1.0)
        tf.margin_bottom = Pt(1.0)
        body_pr = tf._txBody.bodyPr
        for child in list(body_pr):
            if child.tag.endswith('spAutoFit') or child.tag.endswith('normAutofit'):
                body_pr.remove(child)
        if not any(child.tag.endswith('noAutofit') for child in body_pr):
            body_pr.append(parse_xml(f'<a:noAutofit {nsdecls("a")}/>'))
    except Exception:
        pass

def _copy_font_style(src_font, dst_font):
    """Copie fidèlement les attributs d'une police."""
    try:
        if src_font.name: dst_font.name = src_font.name
        if src_font.size: dst_font.size = src_font.size
        if src_font.bold is not None: dst_font.bold = src_font.bold
        if src_font.italic is not None: dst_font.italic = src_font.italic
        if src_font.underline is not None: dst_font.underline = src_font.underline
        if src_font.color and src_font.color.type is not None:
            if src_font.color.type == 1:
                dst_font.color.rgb = src_font.color.rgb
            elif src_font.color.type == 2:
                dst_font.color.theme_color = src_font.color.theme_color
                if hasattr(src_font.color, 'brightness') and src_font.color.brightness is not None:
                    dst_font.color.brightness = src_font.color.brightness
    except Exception:
        pass

def get_paragraph_base_size(p):
    """Extrait la taille de référence du paragraphe s'il en possède une."""
    for r in p.runs:
        if r.font and r.font.size:
            return r.font.size.pt
    if p.font and p.font.size:
        return p.font.size.pt
    return None

# --- SIMULATEUR DE DÉBORDEMENT CALIBRÉ ---

def calculate_optimal_container_scale(paragraphs_data, box_width_pt, box_height_pt):
    """Simule la mise en page exacte dans la boîte."""
    if not box_width_pt or not box_height_pt:
        return 1.0

    scalable = [item for item in paragraphs_data if item['base_size'] and item['base_size'] <= 18.0 and not item['is_title']]
    if not scalable:
        return 1.0

    usable_w = max(20.0, box_width_pt - 10.0)
    usable_h = box_height_pt if box_height_pt <= 25.0 else max(10.0, box_height_pt - 2.0)

    for scale in [1.0, 0.95, 0.90, 0.85, 0.80]:
        total_h = 0.0
        for item in paragraphs_data:
            text = item['text']
            base_sz = item['base_size']
            is_title = item['is_title']

            if not text.strip():
                total_h += (base_sz or 10.0) * 0.20
                continue

            curr_sz = base_sz if (is_title or base_sz is None or base_sz > 18.0) else base_sz * scale
            char_w = curr_sz * 0.46
            chars_per_line = max(10, int(usable_w / char_w))

            sub_lines = text.split('\n')
            lines = 0
            for sub in sub_lines:
                sub_clean = sub.strip()
                if not sub_clean:
                    lines += 1
                else:
                    lines += max(1, (len(sub_clean) + chars_per_line - 1) // chars_per_line)

            line_h = curr_sz * 1.08
            total_h += (lines * line_h) + (0.5 if lines > 1 else 0.0)

        if total_h <= usable_h:
            return scale

    return 0.80

def merge_continuation_paragraphs_in_textbox(tf):
    """Rassemble les faux paragraphes coupés par 'Entrée' au milieu d'une phrase."""
    paras = list(tf.paragraphs)
    if len(paras) <= 1:
        return

    i = 0
    while i < len(paras) - 1:
        curr_p = paras[i]
        next_p = paras[i + 1]

        # Protège impérativement les numéros de page dynamiques
        if curr_p._element.xpath('.//a:fld') or next_p._element.xpath('.//a:fld'):
            i += 1
            continue

        if curr_p.level != next_p.level:
            i += 1
            continue

        curr_text = curr_p.text.strip()
        next_text = next_p.text.strip()

        if not curr_text or not next_text:
            i += 1
            continue

        if curr_text.endswith(':'):
            i += 1
            continue

        if next_text.startswith(('-', '•', '–', '*', '·', '1.', '2.', '3.', '>')):
            i += 1
            continue

        starts_with_lowercase = next_text[0].islower()
        ends_with_connector = bool(re.search(
            r'\b(des|de|du|d\'|par|sur|dans|en|pour|et|ou|avec|sans)\s*$',
            curr_text, re.IGNORECASE
        ))

        if starts_with_lowercase or ends_with_connector:
            space_run = curr_p.add_run()
            space_run.text = " " if not curr_text.endswith(" ") else ""
            if curr_p.runs:
                _copy_font_style(curr_p.runs[0].font, space_run.font)

            for r in next_p.runs:
                new_r = curr_p.add_run()
                new_r.text = r.text
                _copy_font_style(r.font, new_r.font)

            next_p._p.getparent().remove(next_p._p)
            paras.pop(i + 1)
        else:
            i += 1

def serialize_paragraph_runs(p):
    """Capture chaque run avec sa taille, sa graisse et sa couleur réelles."""
    runs_data = []
    run_idx = 0
    raw_text = "".join(r.text for r in p.runs if r.text).strip()
    is_short_title = len(raw_text) < 40 and not raw_text.startswith(('-', '•', '–', '*', '>')) and p.level == 0
    
    for child in p._p:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'br':
            if is_short_title:
                runs_data.append('<br/>')
            else:
                runs_data.append(' ')
        elif tag == 'r':
            from pptx.text.text import _Run
            r = _Run(child, p)
            txt = r.text
            if not txt:
                continue
            
            styles = []
            if r.font.bold is True:
                styles.append("b=1")
            elif r.font.bold is False:
                styles.append("b=0")

            if r.font.italic is True:
                styles.append("i=1")
            elif r.font.italic is False:
                styles.append("i=0")
            
            if r.font.size:
                styles.append(f"s={round(r.font.size.pt, 1)}")
            elif p.font.size:
                styles.append(f"s={round(p.font.size.pt, 1)}")
            
            if r.font.color and r.font.color.type is not None:
                try:
                    if r.font.color.type == 1:
                        styles.append(f"c={r.font.color.rgb}")
                    elif r.font.color.type == 2:
                        styles.append(f"t={r.font.color.theme_color}")
                except Exception:
                    pass
                
            style_attr = ";".join(styles)
            if is_short_title:
                txt_clean = txt.replace('\x0b', '<br/>').replace('\n', '<br/>')
            else:
                txt_clean = txt.replace('\x0b', ' ').replace('\n', ' ').replace('\r', ' ')
            
            if style_attr:
                runs_data.append(f'<r id="{run_idx}" st="{style_attr}">{txt_clean}</r>')
            else:
                runs_data.append(f'<r id="{run_idx}">{txt_clean}</r>')
            run_idx += 1
            
    return "".join(runs_data) if runs_data else p.text

def deserialize_paragraph_runs(p, tagged_text, container_scale=1.0):
    """Reconstruit le paragraphe en restaurant strictement la graisse et les styles d'origine."""
    if p._element.xpath('.//a:fld'):
        return

    orig_runs_info = [(r.font, r.text) for r in p.runs]
    orig_raw_text = "".join(t for _, t in orig_runs_info).strip()
    
    lead_ws = ""
    if orig_runs_info and orig_runs_info[0][1]:
        match_ws = re.match(r'^(\s+)', orig_runs_info[0][1])
        if match_ws:
            lead_ws = match_ws.group(1)

    clean_trans_text = re.sub(r'<[^>]+>', '', tagged_text).strip()
    is_short_title = len(orig_raw_text) < 40 and not orig_raw_text.startswith(('-', '•', '–', '*', '>')) and p.level == 0
    is_dense = container_scale < 1.0 or len(clean_trans_text) > 40

    if is_short_title:
        current_run_idx = -1
        orig_breaks_after = set()
        for child in p._p:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'r':
                current_run_idx += 1
            elif tag == 'br':
                if current_run_idx >= 0:
                    orig_breaks_after.add(current_run_idx)
        if orig_breaks_after and '<br' not in tagged_text:
            for r_idx in sorted(orig_breaks_after, reverse=True):
                tagged_text = re.sub(rf'(<r\s+id="{r_idx}"[^>]*>.*?</r>)', r'\1<br/>', tagged_text)
    else:
        tagged_text = re.sub(r'<br\s*/?>', ' ', tagged_text)

    p.clear()

    if not tagged_text.strip():
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = Pt(3)
        return

    if is_dense and not is_short_title:
        p.space_before = Pt(0)
        p.space_after = Pt(0.5)
        p.line_spacing = 0.86

    token_pattern = re.compile(r'(<br\s*/?>|<r(?:\s+id="(\d+)")?(?:\s+st="([^"]*)")?>(.*?)</r>)', re.DOTALL)
    tokens = list(token_pattern.finditer(tagged_text))

    if not tokens:
        run = p.add_run()
        clean_txt = re.sub(r'<[^>]+>', '', tagged_text).replace('\n', ' ').replace('\r', ' ')
        run.text = lead_ws + clean_txt.lstrip()
        if orig_runs_info:
            _copy_font_style(orig_runs_info[0][0], run.font)
            if run.font.size and run.font.size.pt <= 18.0 and not is_short_title:
                clean_sz = round((run.font.size.pt * container_scale) * 2) / 2
                run.font.size = Pt(clean_sz)
        return

    run_counter = 0
    for m in tokens:
        full_match = m.group(1)
        if full_match.startswith('<br'):
            add_br(p)
            continue
        
        orig_id = int(m.group(2)) if m.group(2) else None
        st_str = m.group(3) or ""
        content = m.group(4)
        
        if not is_short_title:
            content = content.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        
        if run_counter == 0 and lead_ws and not content.startswith(lead_ws):
            content = lead_ws + content.lstrip()
        run_counter += 1
        
        content_subparts = re.split(r'<br\s*/?>', content)
        for sub_idx, sub_text in enumerate(content_subparts):
            if sub_idx > 0:
                add_br(p)
            
            run = p.add_run()
            run.text = sub_text
            
            if orig_id is not None and orig_id < len(orig_runs_info):
                src_font = orig_runs_info[orig_id][0]
                _copy_font_style(src_font, run.font)
                run.font.bold = src_font.bold
            else:
                if "b=1" in st_str: run.font.bold = True
                elif "b=0" in st_str: run.font.bold = False
                
            if "i=1" in st_str: run.font.italic = True
            elif "i=0" in st_str: run.font.italic = False
            
            base_size = None
            for part in st_str.split(";"):
                if part.startswith("s="):
                    try:
                        base_size = float(part.split("=")[1])
                    except Exception:
                        pass
            
            if base_size is not None:
                if base_size <= 18.0 and not is_short_title:
                    clean_sz = round((base_size * container_scale) * 2) / 2
                    run.font.size = Pt(clean_sz)
                else:
                    run.font.size = Pt(base_size)
            elif run.font.size and run.font.size.pt <= 18.0 and not is_short_title:
                clean_sz = round((run.font.size.pt * container_scale) * 2) / 2
                run.font.size = Pt(clean_sz)

# --- EXTRACTION GROUPÉE PAR CONTENEURS SANS ÉCRASER LES CHAMPS DYNAMIQUES ---

def extract_all_pptx_containers(prs, direction="Allemand ➔ Français"):
    """Extrait tous les conteneurs avec calcul de la largeur cumulée en cas de gridSpan."""
    containers_list = []
    
    def _extract_from_shape(shape):
        if shape.has_text_frame:
            lock_shape_geometry(shape.text_frame)
            merge_continuation_paragraphs_in_textbox(shape.text_frame)
            valid_paras = []
            for p in shape.text_frame.paragraphs:
                if p._element.xpath('.//a:fld'):
                    for r in p.runs:
                        if direction == "Allemand ➔ Français" and "Seite" in r.text:
                            r.text = r.text.replace("Seite", "Page")
                        elif direction == "Français ➔ Allemand" and "Page" in r.text:
                            r.text = r.text.replace("Page", "Seite")
                elif p.text.strip():
                    valid_paras.append(p)

            if valid_paras:
                containers_list.append({
                    'tf': shape.text_frame,
                    'box_width': shape.width.pt,
                    'box_height': shape.height.pt,
                    'paragraphs': valid_paras
                })
        elif shape.has_table:
            for r_idx, row in enumerate(shape.table.rows):
                for c_idx, cell in enumerate(row.cells):
                    if cell._tc.attrib.get('hMerge') == '1':
                        continue
                    
                    try: cell.text_frame.word_wrap = True
                    except Exception: pass
                    lock_shape_geometry(cell.text_frame)
                    merge_continuation_paragraphs_in_textbox(cell.text_frame)
                    valid_paras = []
                    for p in cell.text_frame.paragraphs:
                        if p._element.xpath('.//a:fld'):
                            for r in p.runs:
                                if direction == "Allemand ➔ Français" and "Seite" in r.text:
                                    r.text = r.text.replace("Seite", "Page")
                                elif direction == "Français ➔ Allemand" and "Page" in r.text:
                                    r.text = r.text.replace("Page", "Seite")
                        elif p.text.strip():
                            valid_paras.append(p)
                    
                    if valid_paras:
                        grid_span = int(cell._tc.attrib.get('gridSpan', '1'))
                        total_col_width = sum(
                            shape.table.columns[c].width.pt 
                            for c in range(c_idx, min(len(shape.table.columns), c_idx + grid_span))
                        )
                        row_span = int(cell._tc.attrib.get('rowSpan', '1'))
                        total_row_height = sum(
                            shape.table.rows[r].height.pt 
                            for r in range(r_idx, min(len(shape.table.rows), r_idx + row_span))
                        )
                        
                        containers_list.append({
                            'tf': cell.text_frame,
                            'box_width': total_col_width,
                            'box_height': total_row_height,
                            'paragraphs': valid_paras
                        })
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub in shape.shapes:
                _extract_from_shape(sub)

    for slide in prs.slides:
        for s in slide.shapes: _extract_from_shape(s)
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            merge_continuation_paragraphs_in_textbox(slide.notes_slide.notes_text_frame)
            valid_paras = [p for p in slide.notes_slide.notes_text_frame.paragraphs if not p._element.xpath('.//a:fld') and p.text.strip()]
            if valid_paras:
                containers_list.append({
                    'tf': slide.notes_slide.notes_text_frame,
                    'box_width': 500.0,
                    'box_height': 200.0,
                    'paragraphs': valid_paras
                })

    for master in prs.slide_masters:
        for s in master.shapes: _extract_from_shape(s)
        for l in master.slide_layouts:
            for s in l.shapes: _extract_from_shape(s)

    return containers_list

# --- TRAITEMENT DES BATCHS PARALLÈLES ---

def _process_unique_batch(args):
    """Traite un lot de textes uniques avec verrouillage absolu de la langue cible."""
    batch_idx, batch_texts, glossary_snippet, direction = args
    
    batch_in_tokens, batch_out_tokens = 0, 0
    formatted_glossary = f"\n\n--- GLOSSAIRE À RESPECTER IMPÉRATIVEMENT ---\n{glossary_snippet}\n" if glossary_snippet else ""
    
    if direction == "Français ➔ Allemand":
        ex_source, ex_cible = "SPM 4440", "Filiale 4440"
        target_lang_instruction = (
            "1. LANGUE CIBLE EXCLUSIVE : ALLEMAND. Le document final doit être INTÉGRALEMENT en ALLEMAND.\n"
            "   - TITRES HYBRIDES OU PARTIELLEMENT EN ALLEMAND : Si un titre contient un préfixe allemand (ex: 'Wochenkennzahlen | ') "
            "mais se poursuit avec du français (ex: 'EM avec marge négative', 'Poisson', 'Sec'), conserve impérativement le préfixe allemand intact "
            "('Wochenkennzahlen | ') et traduis la partie française vers l'allemand ('EM mit negativer Marge', 'Fisch', 'Troso'). Ne traduis JAMAIS vers le français !\n"
            "   - Si une phrase ou un titre d'origine est DÉJÀ 100% en allemand (ex: 'Wochenkennzahlen | Online Wochenperformance'), "
            "CONSERVE-LE STRICTEMENT EN ALLEMAND TEL QUEL ! Ne le traduis JAMAIS vers le français !"
        )
        grammar_instruction = "9. GRAMMAIRE : Respecte la grammaire allemande, les déclinaisons et les majuscules aux substantifs."
    else:
        ex_source, ex_cible = "Filiale 4440", "SPM 4440"
        target_lang_instruction = (
            "1. LANGUE CIBLE EXCLUSIVE : FRANÇAIS. Le document final doit être INTÉGRALEMENT en FRANÇAIS.\n"
            "   - RÈGLE CRITIQUE D'IDEMPOTENCE : Si une phrase d'origine est DÉJÀ en français, conserve-la strictement en français !"
        )
        grammar_instruction = "9. GRAMMAIRE : Réalise toutes les contractions nécessaires en français ('auprès des', 'du')."

    extra_rules = (
        "\n\n--- CONSIGNES DE STYLE ET DE STRUCTURE ---\n"
        f"{target_lang_instruction}\n"
        "2. BALISES STRICTES : Conserve impérativement toutes les balises <r id=\"...\" st=\"...\">...</r> et <br/>.\n"
        "3. LOCALISATION DU TEXTE : Traduis uniquement le contenu situé À L'INTÉRIEUR des balises sans déplacer les balises ni les fusionner.\n"
        "4. RETOURS À LA LIGNE (<br/>) : Conserve la balise <br/> uniquement si elle est déjà présente dans le texte d'origine.\n"
        "5. STYLE FLUIDE ET CONTINU : Dans les phrases et puces, NE METS AUCUN retour forcé (\\n). Laisse la phrase s'enchaîner naturellement.\n"
        "6. RESPECT DE LA GRAISSE (GRAS) : Ne mets JAMAIS en gras un texte qui n'était pas en gras à l'origine.\n"
        "7. CONCISION POUR LES ENCADRÉS : Dans les encadrés d'analyse, reste direct. Conserve 'dv.' si la source utilise 'dv.'.\n"
        f"8. PRIORITÉ AU GLOSSAIRE : Applique systématiquement les équivalences (ex: '{ex_source}' ➔ '{ex_cible}').\n"
        f"{grammar_instruction}"
    )
    combined_context = formatted_glossary + extra_rules

    prompt = build_prompt(batch_texts, combined_context, direction=direction, is_batch=True)
    ai_translations, in_tok, out_tok = call_gemini_batch(prompt, expected_count=len(batch_texts))
    batch_in_tokens += in_tok
    batch_out_tokens += out_tok

    results = []
    if ai_translations:
        for orig, trans in zip(batch_texts, ai_translations):
            results.append((orig, sanitize_pptx_translation(orig, trans, direction=direction)))
        return batch_idx, results, batch_in_tokens, batch_out_tokens

    for orig in batch_texts:
        fb_prompt = build_prompt(orig, combined_context, direction=direction, is_batch=False)
        fb_trans, fb_in, fb_out = call_gemini_fallback(fb_prompt)
        batch_in_tokens += fb_in
        batch_out_tokens += fb_out
        results.append((orig, sanitize_pptx_translation(orig, fb_trans, direction=direction) if fb_trans else orig))
        
    return batch_idx, results, batch_in_tokens, batch_out_tokens

# --- FONCTION PRINCIPALE ---

def translate_pptx_universal(input_path: str, output_path: str, glossary_path: str = None, progress_callback=None, direction="Allemand ➔ Français"):
    start_time = time.time()
    total_input_tokens, total_output_tokens = 0, 0

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    tmp_in = f"/tmp/in_{os.path.basename(input_path)}"
    tmp_out = f"/tmp/out_{os.path.basename(input_path)}"
    
    shutil.copy(input_path, tmp_in)
    prs = Presentation(tmp_in)
    glossary_df = load_glossary_df(glossary_path)

    containers_data = extract_all_pptx_containers(prs, direction=direction)
    all_paragraphs = [p for c in containers_data for p in c['paragraphs']]
    
    if not all_paragraphs:
        prs.save(tmp_out)
        shutil.copy(tmp_out, output_path)
        os.remove(tmp_in); os.remove(tmp_out)
        return {"status": "success", "output_path": output_path, "time_seconds": 0, "input_tokens": 0, "output_tokens": 0, "total_cost_usd": 0.0}

    serialized_texts = [serialize_paragraph_runs(p) for p in all_paragraphs]

    translation_cache = {}
    unique_texts_to_call = []

    for orig_txt in serialized_texts:
        if orig_txt in translation_cache:
            continue
        
        clean_txt = re.sub(r'<[^>]+>', '', orig_txt).strip()
        if not clean_txt:
            translation_cache[orig_txt] = orig_txt
            continue

        override = check_manual_override(orig_txt)
        if override and direction == "Allemand ➔ Français":
            translation_cache[orig_txt] = override
            continue

        if is_untranslatable_text(clean_txt):
            translation_cache[orig_txt] = orig_txt
            continue

        if is_already_in_target_language(clean_txt, direction):
            translation_cache[orig_txt] = orig_txt
            continue

        unique_texts_to_call.append(orig_txt)

    batch_size = 30
    total_unique = len(unique_texts_to_call)
    batch_args_list = []

    for i in range(0, total_unique, batch_size):
        batch = unique_texts_to_call[i:i + batch_size]
        batch_sample = " ".join([re.sub(r'<[^>]+>', '', t) for t in batch])
        glossary_snippet = get_matching_glossary_rules(batch_sample, glossary_df, direction=direction)
        batch_args_list.append((i // batch_size, batch, glossary_snippet, direction))

    total_batches = len(batch_args_list)
    completed_batches = 0
    max_workers = min(10, max(1, total_batches))

    if total_batches > 0:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            for future in as_completed([executor.submit(_process_unique_batch, arg) for arg in batch_args_list]):
                idx, results, batch_in, batch_out = future.result()
                total_input_tokens += batch_in
                total_output_tokens += batch_out
                
                for orig, trans in results:
                    translation_cache[orig] = trans
                    
                completed_batches += 1
                if progress_callback:
                    progress_callback(completed_batches, total_batches)

    # Réinjection précise conteneur par conteneur
    p_index = 0
    for container in containers_data:
        box_w = container['box_width']
        box_h = container['box_height']
        paras = container['paragraphs']

        paras_info = []
        container_trans_texts = []
        for p in paras:
            orig_txt = serialized_texts[p_index]
            p_index += 1
            t_text = translation_cache.get(orig_txt, orig_txt)
            container_trans_texts.append(t_text)

            clean_t = re.sub(r'<[^>]+>', '', t_text).strip()
            base_sz = get_paragraph_base_size(p)
            orig_raw = "".join(r.text for r in p.runs if r.text).strip()
            is_title = (base_sz is None or base_sz > 18.0 or (len(orig_raw) < 40 and p.level == 0 and not orig_raw.startswith(('-', '•', '–', '*', '>'))))

            paras_info.append({
                'text': clean_t,
                'base_size': base_sz,
                'is_title': is_title
            })

        optimal_scale = calculate_optimal_container_scale(paras_info, box_w, box_h)

        for p, t_text in zip(paras, container_trans_texts):
            deserialize_paragraph_runs(p, t_text, container_scale=optimal_scale)

    prs.save(tmp_out)
    shutil.copy(tmp_out, output_path)
    os.remove(tmp_in); os.remove(tmp_out)
    
    return {
        "status": "success",
        "output_path": output_path,
        "time_seconds": round(time.time() - start_time, 2),
        "input_tokens": total_input_tokens,
        "output_tokens": total_output_tokens,
        "total_cost_usd": ((total_input_tokens / 1_000_000) * 0.075) + ((total_output_tokens / 1_000_000) * 0.30)
    }