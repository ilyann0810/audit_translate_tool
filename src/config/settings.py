import os
import json
import vertexai
from vertexai.generative_models import GenerativeModel, HarmCategory, HarmBlockThreshold
from google.oauth2 import service_account

# Modele et region surchargeables par variables d'environnement, sans toucher
# au code le jour ou le modele change.
DEFAULT_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
DEFAULT_LOCATION = os.environ.get("GEMINI_LOCATION", "europe-west1")

# --- Ajout des paramètres pour désactiver les filtres de sécurité ---
DEFAULT_SAFETY_SETTINGS = {
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
}


def get_project_id():
    """Projet GCP : celui de la cle de service (Databricks), sinon GCP_PROJECT_ID."""
    key = os.environ.get("GCP_SERVICE_KEY")
    if key:
        return json.loads(key).get("project_id")
    return os.environ.get("GCP_PROJECT_ID")


def initialize_gemini():
    """Initialise Vertex AI.

    En production (Databricks), la cle du compte de service arrive par la
    variable GCP_SERVICE_KEY injectee par app.yaml. En local, sans cette
    variable, on utilise les identifiants par defaut de gcloud
    (gcloud auth application-default login) et GCP_PROJECT_ID.
    """
    service_account_key = os.environ.get("GCP_SERVICE_KEY")

    if service_account_key:
        service_account_info = json.loads(service_account_key)
        creds = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=[
                "https://www.googleapis.com/auth/drive.readonly",
                "https://www.googleapis.com/auth/documents.readonly",
                "https://www.googleapis.com/auth/presentations.readonly",
                "https://www.googleapis.com/auth/cloud-platform"
            ],
        )
        vertexai.init(
            project=service_account_info["project_id"],
            location=DEFAULT_LOCATION,
            credentials=creds
        )
    else:
        project = os.environ.get("GCP_PROJECT_ID")
        if not project:
            raise ValueError(
                "Ni GCP_SERVICE_KEY ni GCP_PROJECT_ID dans l'environnement : "
                "impossible d'initialiser Vertex AI.")
        vertexai.init(project=project, location=DEFAULT_LOCATION)

    return GenerativeModel(DEFAULT_MODEL)

gemini_model = initialize_gemini()