from firebase_functions import https_fn
from firebase_admin import initialize_app
import os
import sys

# Afegeix el directori actual al path per poder importar l'app
sys.path.insert(0, os.path.dirname(__file__))

from app import app as flask_app

# Inicialització de Firebase si no s'ha fet
try:
    initialize_app()
except Exception:
    pass

@https_fn.on_request()
def infofoto_backend(req: https_fn.Request) -> https_fn.Response:
    with flask_app.request_context(req.environ):
        return flask_app.full_dispatch_request()
