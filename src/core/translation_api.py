import json
from vertexai.generative_models import GenerationConfig

# Importation de ton instance Gemini et des paramètres de sécurité depuis settings.py
from config.settings import gemini_model, DEFAULT_SAFETY_SETTINGS

def build_prompt(content, context, direction="Allemand ➔ Français", is_batch=False):
    """
    Construit le prompt dynamique pour Vertex AI en fonction du mode (batch ou unitaire).
    """
    if is_batch:
        # Transformation de la liste de textes en dictionnaire indexé pour garantir l'alignement
        payload = {str(i): txt for i, txt in enumerate(content)}
        
        prompt = f"""Tu es un traducteur expert professionnel.
Sens de traduction : {direction}

CONTEXTE ET RÈGLES :
{context}

TÂCHE :
Traduis le dictionnaire JSON suivant. Garde exactement les mêmes clés.
Ne renvoie QUE le JSON valide sans aucun formatage markdown autour.

TEXTES À TRADUIRE :
{json.dumps(payload, ensure_ascii=False)}
"""
    else:
        prompt = f"""Tu es un traducteur expert professionnel.
Sens de traduction : {direction}

CONTEXTE ET RÈGLES :
{context}

TÂCHE :
Traduis le texte suivant en conservant son formatage et ses balises (<b>, <i>, <u>).

TEXTE À TRADUIRE :
{content}
"""
    return prompt


def call_gemini_batch(prompt, expected_count):
    """
    Appelle Gemini avec un prompt demandant un retour en JSON (pour les lots).
    Retourne (liste_des_traductions, input_tokens, output_tokens).
    """
    generation_config = GenerationConfig(
        temperature=0.1,
        response_mime_type="application/json"
    )
    
    try:
        response = gemini_model.generate_content(
            prompt,
            generation_config=generation_config,
            safety_settings=DEFAULT_SAFETY_SETTINGS  # <-- Désactivation des filtres de sécurité
        )
        
        # Récupération sécurisée des compteurs de tokens
        in_tok = getattr(response.usage_metadata, 'prompt_token_count', 0) if hasattr(response, 'usage_metadata') else 0
        out_tok = getattr(response.usage_metadata, 'candidates_token_count', 0) if hasattr(response, 'usage_metadata') else 0
        
        # Vérification des réponses vides (bloquées ou plantées)
        if not response.candidates or not response.candidates[0].content.parts:
            print("⚠️ Gemini a renvoyé une réponse vide (Lot).")
            return None, in_tok, out_tok
            
        data = json.loads(response.text)
        
        # Vérification de l'alignement (est-ce que l'IA a bien traduit tous les segments ?)
        translations = [data.get(str(i)) for i in range(expected_count)]
        
        if any(t is None for t in translations):
            print("⚠️ Lot incomplet détecté. Bascule automatique en mode unitaire.")
            return None, in_tok, out_tok
            
        return translations, in_tok, out_tok
        
    except json.JSONDecodeError:
        print("⚠️ Erreur de formatage JSON renvoyé par Gemini. Bascule unitaire.")
        return None, 0, 0
    except Exception as e:
        print(f"⚠️ Erreur lors de l'appel BATCH à Vertex AI : {e}")
        return None, 0, 0  # <-- Évite le plantage 'cannot unpack non-iterable NoneType'


def call_gemini_fallback(prompt):
    """
    Appelle Gemini en mode texte simple (pour traiter les éléments un par un si le lot échoue).
    Retourne (texte_traduit, input_tokens, output_tokens).
    """
    generation_config = GenerationConfig(
        temperature=0.1
    )
    
    try:
        response = gemini_model.generate_content(
            prompt,
            generation_config=generation_config,
            safety_settings=DEFAULT_SAFETY_SETTINGS  # <-- Désactivation des filtres de sécurité
        )
        
        in_tok = getattr(response.usage_metadata, 'prompt_token_count', 0) if hasattr(response, 'usage_metadata') else 0
        out_tok = getattr(response.usage_metadata, 'candidates_token_count', 0) if hasattr(response, 'usage_metadata') else 0
        
        if not response.candidates or not response.candidates[0].content.parts:
            print("⚠️ Gemini a renvoyé une réponse vide (Fallback).")
            return None, in_tok, out_tok
            
        return response.text.strip(), in_tok, out_tok
        
    except Exception as e:
        print(f"⚠️ Erreur lors de l'appel FALLBACK à Vertex AI : {e}")
        return None, 0, 0  # <-- Évite le plantage 'cannot unpack non-iterable NoneType'