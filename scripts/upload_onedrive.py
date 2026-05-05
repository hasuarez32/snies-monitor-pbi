"""
upload_onedrive.py
------------------
Sube los Excel acumulados de data/novedades/ a OneDrive personal
usando Microsoft Graph API con flujo de refresh token (OAuth2 delegado).

Variables de entorno requeridas (GitHub Secrets):
    MS_CLIENT_ID      — App ID registrada en Azure Portal
    MS_CLIENT_SECRET  — Client secret de la app
    MS_TENANT_ID      — Usar 'consumers' para cuentas outlook.com / MSA
    MS_REFRESH_TOKEN  — Refresh token obtenido con get_onedrive_token.py

Carpeta destino en OneDrive: SNIES-Monitor/
"""

import os
import sys
import logging
import requests
from pathlib import Path

logging.basicConfig(
    stream=sys.stdout,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
NOVEDADES_DIR = ROOT / "data" / "novedades"

ARCHIVOS = [
    "Nuevos_pregrado.xlsx",
    "Inactivos_pregrado.xlsx",
    "Modificados_pregrado.xlsx",
]

ONEDRIVE_FOLDER = "SNIES-Monitor"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
CHUNK_SIZE = 4 * 1024 * 1024  # 4 MB


# ── Auth ─────────────────────────────────────────────────────────────────────

def get_access_token() -> str:
    """Intercambia el refresh token por un access token fresco."""
    tenant = os.environ.get("MS_TENANT_ID", "consumers")
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"

    resp = requests.post(
        url,
        data={
            "grant_type":    "refresh_token",
            "client_id":     os.environ["MS_CLIENT_ID"],
            "client_secret": os.environ["MS_CLIENT_SECRET"],
            "refresh_token": os.environ["MS_REFRESH_TOKEN"],
            "scope":         "Files.ReadWrite.All offline_access",
        },
        timeout=30,
    )

    try:
        resp.raise_for_status()
    except requests.HTTPError:
        log.error("Error al obtener token: %s", resp.text)
        raise

    data = resp.json()
    if "access_token" not in data:
        raise RuntimeError(f"Respuesta inesperada del token endpoint: {data}")

    log.info("Access token obtenido correctamente.")
    return data["access_token"]


# ── OneDrive helpers ──────────────────────────────────────────────────────────

def _auth_header(token: str) -> dict:
    return {"Authorization": f"Bearer {token}"}


def ensure_folder(token: str, folder: str) -> None:
    """Crea la carpeta raíz en OneDrive si no existe."""
    resp = requests.get(
        f"{GRAPH_BASE}/me/drive/root:/{folder}",
        headers=_auth_header(token),
        timeout=30,
    )
    if resp.status_code == 200:
        log.info("Carpeta '%s' ya existe en OneDrive.", folder)
        return
    if resp.status_code != 404:
        resp.raise_for_status()

    resp = requests.post(
        f"{GRAPH_BASE}/me/drive/root/children",
        headers={**_auth_header(token), "Content-Type": "application/json"},
        json={
            "name":   folder,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace",
        },
        timeout=30,
    )
    resp.raise_for_status()
    log.info("Carpeta '%s' creada en OneDrive.", folder)


def upload_simple(token: str, local: Path, folder: str) -> None:
    """PUT directo para archivos <= 4 MB."""
    url = f"{GRAPH_BASE}/me/drive/root:/{folder}/{local.name}:/content"
    with open(local, "rb") as f:
        data = f.read()

    resp = requests.put(
        url,
        headers={
            **_auth_header(token),
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        data=data,
        timeout=120,
    )
    resp.raise_for_status()


def upload_resumable(token: str, local: Path, folder: str) -> None:
    """Upload session para archivos > 4 MB."""
    session_url = f"{GRAPH_BASE}/me/drive/root:/{folder}/{local.name}:/createUploadSession"
    resp = requests.post(
        session_url,
        headers={**_auth_header(token), "Content-Type": "application/json"},
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30,
    )
    resp.raise_for_status()
    upload_url = resp.json()["uploadUrl"]

    file_size = local.stat().st_size
    with open(local, "rb") as f:
        offset = 0
        while offset < file_size:
            chunk = f.read(CHUNK_SIZE)
            end = offset + len(chunk) - 1
            put = requests.put(
                upload_url,
                headers={
                    "Content-Range":  f"bytes {offset}-{end}/{file_size}",
                    "Content-Length": str(len(chunk)),
                },
                data=chunk,
                timeout=120,
            )
            put.raise_for_status()
            offset += len(chunk)


def upload_file(token: str, local: Path, folder: str) -> None:
    size_mb = local.stat().st_size / (1024 * 1024)
    if size_mb > 4:
        upload_resumable(token, local, folder)
    else:
        upload_simple(token, local, folder)
    log.info("  ✓ %s (%.1f MB)", local.name, size_mb)


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    missing = [v for v in ("MS_CLIENT_ID", "MS_CLIENT_SECRET", "MS_REFRESH_TOKEN")
               if not os.environ.get(v)]
    if missing:
        log.error("Variables de entorno faltantes: %s", ", ".join(missing))
        sys.exit(1)

    token = get_access_token()
    ensure_folder(token, ONEDRIVE_FOLDER)

    uploaded, failed = 0, 0
    for name in ARCHIVOS:
        path = NOVEDADES_DIR / name
        if not path.exists():
            log.warning("No encontrado, se omite: %s", path)
            continue
        try:
            upload_file(token, path, ONEDRIVE_FOLDER)
            uploaded += 1
        except Exception as exc:
            log.error("Error subiendo %s: %s", name, exc)
            failed += 1

    log.info("Resultado final: %d subidos, %d fallidos.", uploaded, failed)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
