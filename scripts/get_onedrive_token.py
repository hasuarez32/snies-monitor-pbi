"""
get_onedrive_token.py
---------------------
Script de setup ONE-TIME. Obtiene el refresh token inicial de OneDrive
personal (MSA / outlook.com) via flujo de código de autorización OAuth2.

Cómo usarlo:
    1. Configura las variables al inicio del script con tus credenciales de Azure.
    2. Ejecuta: python scripts/get_onedrive_token.py
    3. Se abrirá tu navegador para hacer login con la cuenta outlook.com.
    4. Tras autorizar, el script imprimirá el REFRESH TOKEN.
    5. Copia ese valor y agrégalo como GitHub Secret: MS_REFRESH_TOKEN

Requisitos extra (solo para este script):
    pip install requests

IMPORTANTE: No commites este archivo con credenciales hardcodeadas.
"""

import json
import os
import secrets
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlparse

import requests

# ── Configura aquí con tus credenciales de Azure ──────────────────────────────
CLIENT_ID     = os.environ.get("MS_CLIENT_ID", "TU_CLIENT_ID_AQUI")
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "TU_CLIENT_SECRET_AQUI")
TENANT_ID     = os.environ.get("MS_TENANT_ID", "consumers")

REDIRECT_URI  = "http://localhost:8765/callback"
SCOPE         = "Files.ReadWrite.All offline_access"
# ─────────────────────────────────────────────────────────────────────────────

auth_code_holder = {}
csrf_state = secrets.token_urlsafe(16)


class CallbackHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path != "/callback":
            self.send_response(404)
            self.end_headers()
            return

        params = parse_qs(parsed.query)

        if params.get("state", [None])[0] != csrf_state:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"Estado CSRF invalido.")
            return

        if "error" in params:
            self.send_response(400)
            self.end_headers()
            msg = params.get("error_description", ["Error desconocido"])[0]
            self.wfile.write(msg.encode())
            auth_code_holder["error"] = msg
            return

        auth_code_holder["code"] = params["code"][0]
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Autorizacion recibida. Puedes cerrar esta pestana.")

    def log_message(self, *args):
        pass  # silencia logs del servidor


def main():
    if CLIENT_ID == "TU_CLIENT_ID_AQUI":
        print("ERROR: Configura MS_CLIENT_ID como variable de entorno o edita el script.")
        sys.exit(1)

    auth_url = (
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={REDIRECT_URI}"
        f"&scope={SCOPE.replace(' ', '%20')}"
        f"&state={csrf_state}"
        f"&prompt=select_account"
    )

    print("Abriendo navegador para autenticación con Microsoft...")
    print(f"URL: {auth_url}\n")
    webbrowser.open(auth_url)

    server = HTTPServer(("localhost", 8765), CallbackHandler)
    print("Esperando callback en http://localhost:8765/callback ...")
    server.handle_request()

    if "error" in auth_code_holder:
        print(f"Error de autenticación: {auth_code_holder['error']}")
        sys.exit(1)

    code = auth_code_holder.get("code")
    if not code:
        print("No se recibió código de autorización.")
        sys.exit(1)

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(
        token_url,
        data={
            "grant_type":    "authorization_code",
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "code":          code,
            "redirect_uri":  REDIRECT_URI,
            "scope":         SCOPE,
        },
        timeout=30,
    )

    try:
        resp.raise_for_status()
    except requests.HTTPError:
        print(f"Error al obtener tokens: {resp.text}")
        sys.exit(1)

    data = resp.json()
    refresh_token = data.get("refresh_token")

    if not refresh_token:
        print(f"No se obtuvo refresh_token. Respuesta: {json.dumps(data, indent=2)}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("REFRESH TOKEN OBTENIDO EXITOSAMENTE")
    print("=" * 60)
    print(refresh_token)
    print("=" * 60)
    print("\nAgrégalo como GitHub Secret con el nombre: MS_REFRESH_TOKEN")
    print("Ve a: Settings → Secrets and variables → Actions → New repository secret")


if __name__ == "__main__":
    main()
