import os
import gzip
import base64
import tempfile
import requests
from contextlib import contextmanager
from cryptography.hazmat.primitives.serialization import (
    pkcs12, Encoding, PrivateFormat, NoEncryption
)


class DistribuicaoConnector:

    def __init__(self, pfx_path: str, pfx_password: str, base_url: str, timeout: int = 45):
        self.pfx_path    = pfx_path
        self.pfx_password = pfx_password
        self.base_url    = base_url.rstrip("/")
        self.timeout     = timeout

        if not os.path.exists(self.pfx_path):
            raise FileNotFoundError(f"Certificado não encontrado: {self.pfx_path}")

    @contextmanager
    def _cert(self):
        with open(self.pfx_path, "rb") as f:
            pfx_data = f.read()

        key, cert, extras = pkcs12.load_key_and_certificates(
            pfx_data,
            self.pfx_password.encode("utf-8")
        )

        cert_pem = cert.public_bytes(Encoding.PEM)
        for extra in (extras or []):
            cert_pem += extra.public_bytes(Encoding.PEM)

        key_pem = key.private_bytes(Encoding.PEM, PrivateFormat.PKCS8, NoEncryption())

        ct = tempfile.NamedTemporaryFile(delete=False, suffix=".pem")
        kt = tempfile.NamedTemporaryFile(delete=False, suffix=".pem")
        try:
            ct.write(cert_pem); ct.flush(); ct.close()
            kt.write(key_pem);  kt.flush(); kt.close()
            yield ct.name, kt.name
        finally:
            os.unlink(ct.name)
            os.unlink(kt.name)

    def buscar_lote(self, nsu: int, cnpj_consulta: str = None) -> dict:
        """
        GET /DFe/{NSU}?lote=true&cnpjConsulta={cnpj}
        cnpj_consulta opcional — testa se o certificado da matriz
        consegue buscar notas de uma filial específica.
        """
        url = f"{self.base_url}/DFe/{nsu}?lote=true"
        if cnpj_consulta:
            cnpj_limpo = "".join(filter(str.isdigit, cnpj_consulta))
            url += f"&cnpjConsulta={cnpj_limpo}"

        with self._cert() as cert_tuple:
            try:
                resp = requests.get(url, cert=cert_tuple, timeout=self.timeout)
                if resp.status_code == 404:
                    return {"sucesso": False, "erro": "HTTP 404", "dados": None}
                resp.raise_for_status()
                return {"sucesso": True, "dados": resp.json()}
            except requests.HTTPError as e:
                return {"sucesso": False, "erro": f"HTTP {resp.status_code}", "dados": None}
            except Exception as e:
                return {"sucesso": False, "erro": str(e), "dados": None}

    @staticmethod
    def descompactar(gzip_b64: str) -> str:
        return gzip.decompress(base64.b64decode(gzip_b64)).decode("utf-8")
