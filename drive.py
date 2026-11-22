"""Read-only access to the historical XLSX exports for the fixed spreadsheet."""

import re
import sys
from time import monotonic, sleep
from urllib.parse import quote, urljoin, urlsplit

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1M2OPMmj3aTnaXGicUWGkOnFhY8lqVNv-qxDWNKA20QY/edit?usp=sharing"
FILE_ID = SPREADSHEET_URL.split("/")[5]
API = "https://www.googleapis.com/drive/v3/"
FILE = f"{API}files/{FILE_ID}"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SCOPE = "https://www.googleapis.com/auth/drive.readonly"


def connect(credentials_path):
    if not credentials_path.is_file():
        raise ValueError(f"Chybí OAuth soubor {credentials_path}. Postup vytvoření je v README.md.")

    from google.auth.transport.requests import AuthorizedSession
    from google_auth_oauthlib.flow import InstalledAppFlow, WSGITimeoutError
    from oauthlib.oauth2 import OAuth2Error
    from requests import RequestException

    flow = InstalledAppFlow.from_client_secrets_file(
        str(credentials_path), scopes=[SCOPE], autogenerate_code_verifier=True,
    )
    try:
        credentials = flow.run_local_server(
            host="127.0.0.1", port=0, timeout_seconds=300,
            authorization_prompt_message="Přihlášení ke Googlu: {url}",
            success_message="Přihlášení dokončeno. Vraťte se do terminálu.",
        )
    except (OAuth2Error, WSGITimeoutError, RequestException) as error:
        raise ValueError(f"Přihlášení ke Googlu se nezdařilo ({type(error).__name__}).") from None
    return Drive(AuthorizedSession(credentials))


class Drive:
    def __init__(self, session):
        self.session = session
        self.resource_key = None

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.session.close()

    def request(self, method, url, *, params=None, resource_key=None):
        from google.auth.exceptions import GoogleAuthError
        from requests import RequestException

        headers = {}
        if resource_key:
            headers["X-Goog-Drive-Resource-Keys"] = f"{FILE_ID}/{resource_key}"
        for attempt in range(4):
            try:
                response = self.session.request(
                    method, url, params=params, headers=headers, timeout=(10, 60),
                    allow_redirects=False,
                )
            except GoogleAuthError:
                raise RuntimeError("Google odmítl přihlašovací údaje.") from None
            except RequestException:
                if attempt == 3:
                    raise RuntimeError("Nepodařilo se navázat spojení s Google Drive.") from None
            else:
                if response.status_code not in (429, 500, 502, 503, 504):
                    break
            if attempt < 3:
                sleep(2 ** attempt)
        if not 200 <= response.status_code < 400:
            status = response.status_code
            detail = ""
            try:
                reasons = response.json()["error"].get("errors", [])
                if reasons:
                    reason = reasons[0].get("reason", "")
                    if re.fullmatch(r"[A-Za-z0-9_]+", reason):
                        detail = f" ({reason})"
            except (ValueError, KeyError, TypeError, AttributeError):
                pass
            hint = " Ověřte účet, oprávnění k úpravám a možnost stažení." if status in (401, 403, 404) else ""
            raise RuntimeError(f"Google Drive vrátil HTTP {status}{detail}.{hint}")
        return response

    def json(self, method, url, *, params=None):
        response = self.request(method, url, params=params, resource_key=self.resource_key)
        if response.status_code >= 300:
            raise RuntimeError("Drive API vrátilo neočekávané přesměrování.")
        try:
            data = response.json()
        except ValueError:
            raise RuntimeError("Drive API vrátilo neplatnou odpověď JSON.") from None
        if not isinstance(data, dict):
            raise RuntimeError("Drive API vrátilo neplatnou odpověď JSON.")
        return data

    def revisions(self):
        file = self.json("GET", FILE, params={
            "fields": "resourceKey,capabilities(canReadRevisions,canDownload)",
            "supportsAllDrives": "true",
        })
        caps = file.get("capabilities", {})
        if caps.get("canReadRevisions") is False or caps.get("canDownload") is False:
            raise RuntimeError("Účet nemá přístup k historii revizí nebo nemůže tabulku stáhnout.")
        self.resource_key = file.get("resourceKey")

        revisions = {}
        params = {"fields": "nextPageToken,revisions(id,modifiedTime)", "pageSize": 100}
        while True:
            page = self.json("GET", f"{FILE}/revisions", params=params)
            for entry in page.get("revisions", []):
                if not entry.get("id") or not entry.get("modifiedTime"):
                    raise RuntimeError("Drive API vrátilo revizi bez identifikátoru nebo času.")
                revisions[entry["id"]] = entry["modifiedTime"]
            token = page.get("nextPageToken")
            if not token:
                return revisions
            params = {**params, "pageToken": token}

    def download(self, revision_id):
        print(f"Stahuji revizi {revision_id}…", file=sys.stderr, flush=True)
        operation = self.json("POST", f"{FILE}/download", params={
            "mimeType": XLSX, "revisionId": revision_id,
        })
        deadline = monotonic() + 180
        delay = 1
        while not operation.get("done"):
            name = operation.get("name", "").removeprefix("operations/")
            if not name:
                raise RuntimeError("Google nevrátil identifikátor operace stahování.")
            if monotonic() + delay > deadline:
                raise RuntimeError("Vypršel čas na přípravu exportu XLSX.")
            sleep(delay)
            delay = min(delay * 2, 10)
            operation = self.json("GET", API + "operations/" + quote(name, safe=""))
        if operation.get("error"):
            code = operation["error"].get("code", "?")
            raise RuntimeError(f"Google nemohl exportovat revizi XLSX (kód {code}).")

        uri = operation.get("response", {}).get("downloadUri")
        if not uri:
            raise RuntimeError("Google nevrátil odkaz na export XLSX.")
        resource_key = operation.get("metadata", {}).get("resourceKey") or self.resource_key
        for _ in range(5):
            host = urlsplit(uri).hostname or ""
            trusted = ("googleapis.com", "googleusercontent.com", "google.com")
            if urlsplit(uri).scheme != "https" or not any(
                host == domain or host.endswith("." + domain) for domain in trusted
            ):
                raise RuntimeError("Google vrátil neočekávanou adresu ke stažení.")
            response = self.request("GET", uri, resource_key=resource_key)
            if 300 <= response.status_code < 400 and response.headers.get("Location"):
                uri = urljoin(uri, response.headers["Location"])
                continue
            if response.status_code != 200 or not response.content.startswith(b"PK\x03\x04"):
                raise RuntimeError("Google místo souboru XLSX vrátil jiný obsah.")
            return response.content
        raise RuntimeError("Při stahování nastalo příliš mnoho přesměrování.")
