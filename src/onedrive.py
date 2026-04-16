import msal
import requests

SCOPES = ["https://graph.microsoft.com/Files.ReadWrite"]
GRAPH_URL = "https://graph.microsoft.com/v1.0"
APP_FOLDER = "hr-monthly-reports"


def build_app(client_id: str, tenant_id: str, token_cache_str: str = None):
    cache = msal.SerializableTokenCache()
    if token_cache_str:
        cache.deserialize(token_cache_str)
    # Χρησιμοποιούμε "common" για personal Microsoft accounts
    authority = f"https://login.microsoftonline.com/common"
    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
    )
    return app, cache


def get_token_silent(app, cache):
    accounts = app.get_accounts()
    if not accounts:
        return None, cache.serialize()
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        return result["access_token"], cache.serialize()
    return None, cache.serialize()


def start_device_flow(app):
    return app.initiate_device_flow(scopes=SCOPES)


def complete_device_flow(app, flow):
    return app.acquire_token_by_device_flow(flow)


def upload_file(token: str, filename: str, content: bytes, subfolder: str = "output"):
    path = f"{APP_FOLDER}/{subfolder}/{filename}"
    url = f"{GRAPH_URL}/me/drive/root:/{path}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    r = requests.put(url, headers=headers, data=content)
    r.raise_for_status()
    return r.json()


def list_files(token: str, subfolder: str = "output"):
    path = f"{APP_FOLDER}/{subfolder}"
    url = f"{GRAPH_URL}/me/drive/root:/{path}:/children"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    if r.status_code == 404:
        return []
    r.raise_for_status()
    return r.json().get("value", [])


def download_file(token: str, filename: str, subfolder: str = "output") -> bytes:
    path = f"{APP_FOLDER}/{subfolder}/{filename}"
    url = f"{GRAPH_URL}/me/drive/root:/{path}:/content"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content
