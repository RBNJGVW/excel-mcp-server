import contextlib
import logging
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from typing import Iterator, List, Optional

try:
    from azure.identity import DefaultAzureCredential, InteractiveBrowserCredential
    from azure.storage.blob import BlobServiceClient, ContentSettings
except ImportError:
    BlobServiceClient = None
    ContentSettings = None
    DefaultAzureCredential = None
    InteractiveBrowserCredential = None

log = logging.getLogger("excel-mcp.storage")

AZBLOB_SCHEME = "azblob://"


def _norm_prefix(prefix: str) -> str:
    p = prefix.strip().lstrip("/")
    return "" if p in ("", "/") else p


def _join_blob(prefix: str, name: str) -> str:
    prefix = _norm_prefix(prefix)
    name = name.lstrip("/")
    return f"{prefix}/{name}" if prefix else name


class StorageBackend:
    """
    Backend abstracto: maneja lecturas/escrituras locales y 'azblob://'.
    Usar a través de get_storage(base_path).
    """

    def __init__(self, base_path: str):
        self.base_path = base_path
        self._tmp_root = os.environ.get("EXCEL_TMP_DIR", tempfile.gettempdir())
        os.makedirs(self._tmp_root, exist_ok=True)

        self._is_blob = base_path.startswith(AZBLOB_SCHEME)
        self._local_base = base_path  # válido en modo local

        # Campos AZ Blob
        self._container: Optional[str] = None
        self._prefix: str = ""
        self._bsc = None  # BlobServiceClient
        self._container_client = None

        if self._is_blob:
            self._parse_blob_url()
            self._bsc = self._build_blob_service_client()
            self._container_client = self._bsc.get_container_client(self._container)
            # Forzar conexión para abrir el navegador
            self._warm_up_connection()

    # ----------------------------
    # Inicialización AZ Blob
    # ----------------------------

    def _warm_up_connection(self):
        """Intenta contactar con Azure para forzar el popup de login al inicio."""
        try:
            print("🔄 Iniciando conexión con Azure Blob Storage...")
            # .exists() hace una llamada de red ligera (HEAD)
            if self._container_client.exists():
                print(f"✅ Conectado exitosamente al contenedor: {self._container}")
            else:
                print(f"⚠️ El contenedor '{self._container}' no parece existir.")
        except Exception as e:
            # Aquí saltará el error si cancelas el login o falla la red
            print(f"❌ Error conectando a Azure: {e}")
            raise e

    def _parse_blob_url(self) -> None:
        """
        EXCEL_FILES_PATH = azblob://<container>/<optional/prefix>
        """
        raw = self.base_path[len(AZBLOB_SCHEME) :]
        parts = raw.split("/", 1)
        if not parts or not parts[0]:
            raise ValueError(
                "Formato inválido para EXCEL_FILES_PATH. Esperado: azblob://<container>/<prefix-opcional>"
            )
        self._container = parts[0]
        self._prefix = _norm_prefix(parts[1] if len(parts) > 1 else "")

    def _build_blob_service_client(self):
        """
        Crea un BlobServiceClient con la siguiente prioridad:
        1. InteractiveBrowserCredential (SOLO LOCAL si se activa explícitamente)
        2. DefaultAzureCredential (Producción/Managed Identity)
        """
        if BlobServiceClient is None or DefaultAzureCredential is None:
            raise RuntimeError(
                "Librerías de Azure faltantes. Instala 'azure-storage-blob' y 'azure-identity'."
            )

        # Configuración de URL
        account_name = os.environ.get("AZURE_STORAGE_ACCOUNT")
        account_url = os.environ.get("AZURE_STORAGE_ACCOUNT_URL")
        if not account_url and account_name:
            account_url = f"https://{account_name}.blob.core.windows.net"

        if not account_url:
            raise ValueError("Faltan credenciales. Define AZURE_STORAGE_ACCOUNT.")

        # 3. Interactive Browser (SOLO LOCAL)
        # IMPORTANTE: Activamos esto solo si una variable de entorno lo pide.
        # De lo contrario, romperá el servidor en producción al intentar abrir un navegador.
        if os.environ.get("AZURE_USE_INTERACTIVE_AUTH", "").lower() == "true":
            tenant_id = os.environ.get("AZURE_TENANT_ID")
            log.info(
                f"[ALERTA] MODO INTERACTIVO ACTIVADO. Usando InteractiveBrowserCredential (Tenant: {tenant_id}). "
                "Esto requiere interacción humana y navegador."
            )
            # Si tenant_id es None, Azure usará el tenant por defecto del usuario
            credential = InteractiveBrowserCredential(tenant_id=tenant_id)
            return BlobServiceClient(account_url=account_url, credential=credential)

        # 4. Managed Identity / DefaultAzureCredential (Producción)
        log.info("Usando DefaultAzureCredential (Managed Identity) para Azure Blob.")
        managed_identity_client_id = os.environ.get("AZURE_CLIENT_ID")
        credential = DefaultAzureCredential(
            managed_identity_client_id=managed_identity_client_id
        )

        return BlobServiceClient(account_url=account_url, credential=credential)

    # ----------------------------
    # API pública
    # ----------------------------

    @property
    def is_blob(self) -> bool:
        return self._is_blob

    def list_names(self, pattern: Optional[str] = None) -> List[str]:
        if not self._is_blob:
            names = []
            base = self._local_base
            if not os.path.isdir(base):
                return []
            for root, _, files in os.walk(base):
                for f in files:
                    rel = os.path.relpath(os.path.join(root, f), base)
                    names.append(rel.replace("\\", "/"))
        else:
            names = []

            for blob in self._container_client.list_blobs(
                name_starts_with=self._prefix or None
            ):
                name = blob.name
                if self._prefix and name.startswith(self._prefix + "/"):
                    name = name[len(self._prefix) + 1 :]
                elif self._prefix and name == self._prefix:
                    name = os.path.basename(name)
                names.append(name)

        if pattern:
            # Escapamos puntos y convertimos glob simple a regex
            rx = re.compile("^" + pattern.replace(".", r"\.").replace("*", ".*") + "$")
            names = [n for n in names if rx.match(n)]
        return names

    def exists(self, name: str) -> bool:
        if not self._is_blob:
            return os.path.exists(os.path.join(self._local_base, name))

        blob_name = _join_blob(self._prefix, name)
        try:
            return self._container_client.get_blob_client(blob_name).exists()
        except Exception:
            return False

    @contextlib.contextmanager
    def local_read(self, name: str) -> Iterator[str]:
        if not self._is_blob:
            path = os.path.join(self._local_base, name)
            if not os.path.exists(path):
                raise FileNotFoundError(path)
            yield path
            return

        tmpdir = tempfile.mkdtemp(prefix="excel-mcp-r-", dir=self._tmp_root)
        local_path = os.path.join(tmpdir, os.path.basename(name))
        blob_name = _join_blob(self._prefix, name)
        bc = self._container_client.get_blob_client(blob_name)

        try:
            with open(local_path, "wb") as f:
                bc.download_blob(max_concurrency=2).readinto(f)
            yield local_path
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    @contextlib.contextmanager
    def local_write(self, name: str) -> Iterator[str]:
        if not self._is_blob:
            dest = os.path.join(self._local_base, name)
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            yield dest
            return

        tmpdir = tempfile.mkdtemp(prefix="excel-mcp-w-", dir=self._tmp_root)
        local_path = os.path.join(tmpdir, os.path.basename(name))

        try:
            blob_name = _join_blob(self._prefix, name)
            bc = self._container_client.get_blob_client(blob_name)

            # Optimización: comprobar existencia antes de descargar
            if bc.exists():
                with open(local_path, "wb") as f:
                    bc.download_blob(max_concurrency=2).readinto(f)
                log.debug(f"[blob] predescargado: {blob_name} -> {local_path}")
            else:
                log.debug(f"[blob] no existe aún: {blob_name} (se creará nuevo)")

            yield local_path

            # SUBIDA
            content_settings = None
            if ContentSettings and (
                name.lower().endswith(".xlsx") or name.lower().endswith(".xlsm")
            ):
                content_settings = ContentSettings(
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with open(local_path, "rb") as f:
                bc.upload_blob(f, overwrite=True, content_settings=content_settings)
            log.debug(f"[blob] subido: {local_path} -> {blob_name}")

        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def delete(self, name: str) -> None:
        if not self._is_blob:
            try:
                os.remove(os.path.join(self._local_base, name))
            except FileNotFoundError:
                pass
            return

        blob_name = _join_blob(self._prefix, name)
        self._container_client.get_blob_client(blob_name).delete_blob()

    def normalize_name(self, name: str) -> str:
        n = (name or "").strip().replace("\\", "/").lstrip("/")
        if self._is_blob and self._prefix:
            pref = self._prefix.rstrip("/")
            if n == pref:
                return os.path.basename(n)
            if n.startswith(pref + "/"):
                n = n[len(pref) + 1 :]
        return n


def get_storage(base_path: Optional[str]) -> StorageBackend:
    base = base_path or "./excel_files"
    return StorageBackend(base)
