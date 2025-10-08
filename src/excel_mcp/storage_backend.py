import contextlib
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from typing import Iterator, List, Optional

try:
    from azure.storage.blob import BlobServiceClient, ContentSettings
except Exception:
    BlobServiceClient = None
    ContentSettings = None


AZBLOB_SCHEME = "azblob://"


def _norm_prefix(prefix: str) -> str:
    p = prefix.strip().lstrip("/")
    return "" if p in ("", "/") else p


def _join_blob(prefix: str, name: str) -> str:
    prefix = _norm_prefix(prefix)
    name = name.lstrip("/")
    return f"{prefix}/{name}" if prefix else name


@dataclass
class _AzCreds:
    account_url: str
    credential: object
    source: str


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

    # ----------------------------
    # Inicialización AZ Blob
    # ----------------------------

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
        Crea un BlobServiceClient a partir de:
        - AZURE_STORAGE_CONNECTION_STRING
        - o (AZURE_STORAGE_ACCOUNT + AZURE_STORAGE_KEY)
        - o (AZURE_STORAGE_ACCOUNT + AZURE_STORAGE_SAS_TOKEN)
        - opcional: AZURE_STORAGE_ACCOUNT_URL (si usas dominio privado)
        """
        if BlobServiceClient is None:
            raise RuntimeError(
                "azure-storage-blob no está instalado. Añade 'azure-storage-blob' a requirements.txt"
            )

        conn = os.environ.get("AZURE_STORAGE_CONNECTION_STRING")
        if conn:
            return BlobServiceClient.from_connection_string(conn)

        account = os.environ.get("AZURE_STORAGE_ACCOUNT")
        account_url = os.environ.get("AZURE_STORAGE_ACCOUNT_URL") or (
            f"https://{account}.blob.core.windows.net" if account else None
        )
        key = os.environ.get("AZURE_STORAGE_KEY")
        sas = os.environ.get("AZURE_STORAGE_SAS_TOKEN")

        if account_url and key:
            return BlobServiceClient(account_url=account_url, credential=key)
        if account_url and sas:
            return BlobServiceClient(account_url=account_url, credential=sas)

        raise RuntimeError(
            "No hay credenciales de Azure Blob. Configura "
            "AZURE_STORAGE_CONNECTION_STRING o "
            "(AZURE_STORAGE_ACCOUNT + AZURE_STORAGE_KEY) o "
            "(AZURE_STORAGE_ACCOUNT + AZURE_STORAGE_SAS_TOKEN)."
        )

    # ----------------------------
    # API pública
    # ----------------------------

    @property
    def is_blob(self) -> bool:
        return self._is_blob

    def list_names(self, pattern: Optional[str] = None) -> List[str]:
        """
        Lista nombres lógicos (no rutas locales). pattern soporta comodín '*.xlsx'.
        """
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
                # Convertimos a nombre lógico relativo al prefix
                name = blob.name
                if self._prefix and name.startswith(self._prefix + "/"):
                    name = name[len(self._prefix) + 1 :]
                elif self._prefix and name == self._prefix:
                    name = os.path.basename(name)
                names.append(name)

        if pattern:
            rx = re.compile("^" + pattern.replace(".", r"\.").replace("*", ".*") + "$")
            names = [n for n in names if rx.match(n)]
        return names

    def exists(self, name: str) -> bool:
        if not self._is_blob:
            return os.path.exists(os.path.join(self._local_base, name))
        blob_name = _join_blob(self._prefix, name)
        try:
            self._container_client.get_blob_client(blob_name).get_blob_properties()
            return True
        except Exception:
            return False

    @contextlib.contextmanager
    def local_read(self, name: str) -> Iterator[str]:
        """
        Yields una ruta local existente para leer 'name'.
        """
        if not self._is_blob:
            path = os.path.join(self._local_base, name)
            if not os.path.exists(path):
                raise FileNotFoundError(path)
            yield path
            return

        # Blob → descargar a tmp
        tmpdir = tempfile.mkdtemp(prefix="excel-mcp-r-", dir=self._tmp_root)
        local_path = os.path.join(tmpdir, os.path.basename(name))
        blob_name = _join_blob(self._prefix, name)
        bc = self._container_client.get_blob_client(blob_name)
        with open(local_path, "wb") as f:
            bc.download_blob(max_concurrency=2).readinto(f)
        try:
            yield local_path
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    @contextlib.contextmanager
    def local_write(self, name: str) -> Iterator[str]:
        """
        Yields una ruta local para ESCRIBIR 'name'.
        - Local FS: ruta definitiva (dentro de EXCEL_FILES_PATH).
        - Blob: ruta temporal; al salir del context se sube automáticamente.
        """
        if not self._is_blob:
            dest = os.path.join(self._local_base, name)
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            yield dest
            # nada que hacer al salir
            return

        tmpdir = tempfile.mkdtemp(prefix="excel-mcp-w-", dir=self._tmp_root)
        local_path = os.path.join(tmpdir, os.path.basename(name))
        try:
            yield local_path
            # Subir a Blob al cerrar
            blob_name = _join_blob(self._prefix, name)
            bc = self._container_client.get_blob_client(blob_name)
            content_settings = None
            if ContentSettings:
                # content-type de XLSX
                content_settings = ContentSettings(
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with open(local_path, "rb") as f:
                bc.upload_blob(f, overwrite=True, content_settings=content_settings)
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
        try:
            self._container_client.delete_blob(blob_name)
        except Exception:
            pass


def get_storage(base_path: Optional[str]) -> StorageBackend:
    base = base_path or "./excel_files"
    return StorageBackend(base)
