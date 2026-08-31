"""verify_runtime.py

Executes every ```python code block in the reference docs against real fixture
files (generated on the fly) and reports any exception. This catches signature
and behavior bugs that name-only checks miss (e.g. count() vs length, wrong
argument kinds). API errors (AttributeError/TypeError about a call) are real;
fixture/business-logic errors are reported but easy to tell apart by message.
"""

import ast
import base64
import io
import os
import pathlib
import re
import traceback

import aspose.cells as gc

REFS = pathlib.Path(__file__).parent.parent / "references"
TMP = pathlib.Path(__file__).parent.parent / ".rtfix"
TMP.mkdir(parents=True, exist_ok=True)

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQAY3Y2wAAAAAElFTkSuQmCC"
)

# Extensions we can generate as a real workbook
XLSX_EXTS = {".xlsx", ".xlsm", ".xlsb", ".xls"}


def make_fixture(path: pathlib.Path):
    path = path.resolve()
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    ext = path.suffix.lower()
    try:
        if ext in XLSX_EXTS:
            fmt = gc.SaveFormat.XLSM if ext == ".xlsm" else gc.SaveFormat.XLSX
            wb = gc.Workbook()
            if ext == ".xlsm":
                # ensure it is macro-enabled by having a vba_project place
                pass
            wb.save(str(path), fmt)
        elif ext == ".json":
            path.write_text('{"a":1,"b":2,"c":[3,4]}', encoding="utf-8")
        elif ext in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}:
            path.write_bytes(PNG)
        elif ext in {".txt", ".docx", ".doc", ".pdf", ".html", ".csv", ".ods", ".tsv"}:
            path.write_bytes(b"dummy")
        elif ext == ".lic":
            path.write_bytes(b"dummy")
        elif ext == ".pfx":
            make_pfx(path)
        else:
            path.write_bytes(b"dummy")
    except Exception as e:
        print(f"  [fixture create failed] {path}: {e}")


def make_pfx(path: pathlib.Path):
    try:
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        import datetime

        key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        subj = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "Test")])
        cert = (
            x509.CertificateBuilder()
            .subject_name(subj)
            .issuer_name(subj)
            .public_key(key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(datetime.datetime.utcnow())
            .not_valid_after(datetime.datetime.utcnow() + datetime.timedelta(days=365))
            .sign(key, hashes.SHA256())
        )
        pfx = cert  # placeholder; real DER PKCS12 below
        from cryptography.hazmat.primitives.serialization import pkcs12, BestAvailableEncryption

        data = pkcs12.serialize_key_and_certificates(
            b"cert.pfx", key, cert, None, BestAvailableEncryption(b"password")
        )
        path.write_bytes(data)
    except Exception as e:
        print(f"  [pfx create failed] {e}")


FILENAME_RE = re.compile(r"['\"]([^'\"]+\.(xlsx|xlsm|xlsb|xls|json|png|jpg|jpeg|gif|bmp|tif|tiff|txt|docx|doc|pdf|html|csv|ods|tsv|lic|pfx))['\"]")


def block_uses_pfx(src: str) -> bool:
    return "cert.pfx" in src or "DigitalSignature" in src


def main():
    errors = []
    ran = 0
    for f in sorted(REFS.glob("*.md")):
        txt = f.read_text(encoding="utf-8")
        blocks = re.findall(r"```python(.*?)```", txt, re.S)
        for bi, b in enumerate(blocks):
            src = b.strip()
            if not src:
                continue
            # gather referenced fixture files (relative paths anchored to TMP)
            for m in FILENAME_RE.finditer(src):
                p = pathlib.Path(m.group(1))
                if not p.is_absolute():
                    p = TMP / p
                make_fixture(p)
            if block_uses_pfx(src):
                # make sure pfx exists (best effort)
                make_fixture(TMP / "cert.pfx")
            # always have a (dummy) license file so set_license does not FileNotFound
            make_fixture(TMP / "Aspose.Cells.lic")
            cwd = os.getcwd()
            os.chdir(TMP)
            ns = {"__name__": "__rt__", "__builtins__": __builtins__}
            try:
                # allow top-level statements
                tree = ast.parse(src)
                compiled = compile(tree, f"<{f.name}:{bi}>", "exec")
                exec(compiled, ns)
                ran += 1
            except Exception as e:
                tb = traceback.format_exc(limit=2)
                errors.append((f.name, bi, src, tb))
                # concise inline summary
                msg = str(e).replace("\n", " ")[:160]
                print(f"ERR {f.name} block {bi}: {type(e).__name__}: {msg}")
            finally:
                os.chdir(cwd)
    print(f"blocks executed: {ran}")
    print(f"blocks with exceptions: {len(errors)}")
    for name, bi, src, tb in errors:
        print("=" * 60)
        print(f"{name} block {bi}")
        print(tb)
        # show first line of the block for context
        print("  code:", src.splitlines()[0][:90])


if __name__ == "__main__":
    main()
