import os
import json
import vertexai
from vertexai.generative_models import GenerativeModel, HarmCategory, HarmBlockThreshold # <-- Ajout des imports ici
from google.oauth2 import service_account

DEFAULT_MODEL = "gemini-2.5-flash"
DEFAULT_LOCATION = "europe-west1"

# --- Ajout des paramètres pour désactiver les filtres de sécurité ---
DEFAULT_SAFETY_SETTINGS = {
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
}

def initialize_gemini():
    """Récupère les secrets depuis l'environnement et initialise Vertex AI."""
    
    # Lecture directe de la variable d'environnement injectée par app.yaml
    service_account_key = os.environ.get("GCP_SERVICE_KEY")
    
    if not service_account_key:
        raise ValueError("La clé GCP_SERVICE_KEY est introuvable dans l'environnement.")
        
    service_account_info = json.loads(service_account_key)
    
    # Définition des autorisations
    creds = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=[
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/documents.readonly",
            "https://www.googleapis.com/auth/presentations.readonly",
            "https://www.googleapis.com/auth/cloud-platform"
        ],
    )

    # Initialisation de Vertex AI (Gemini)
    vertexai.init(
        project=service_account_info["project_id"],
        location=DEFAULT_LOCATION,
        credentials=creds
    )
    
    return GenerativeModel(DEFAULT_MODEL)

gemini_model = initialize_gemini()