"""
classify.py - Map an attachment to an ingestion `scan_state` using the same
branch evidence oletools/msoffcrypto actually produce.

Usage:
    .venv/bin/python encryption-triage-lab/scripts/classify.py <file> [<file> ...]
    .venv/bin/python encryption-triage-lab/scripts/classify.py --demo

States (see ../README.md section 3):
    ENCRYPTED_VALID       msoffcrypto parses it and reports encrypted
    ENCRYPTED_NO_PASSWORD valid encryption, but no supplied password works
    ENCRYPTION_SPOOF      encryption-like stream names, content not parseable
    CORRUPT_CONTAINER     OLE magic present but structure unparseable/empty
    PLAIN_MACRO           unencrypted, contains a VBA project
    PLAIN_CLEAN           unencrypted, no VBA project
"""
import io
import os
import sys
import zipfile

import olefile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))        # repo root for `oletools`
sys.path.insert(0, HERE)

from oletools import crypto  # noqa: E402
import msoffcrypto           # noqa: E402

# ordered list of return values
STATES = [
    "ENCRYPTED_VALID", "ENCRYPTED_NO_PASSWORD", "ENCRYPTION_SPOOF",
    "CORRUPT_CONTAINER", "PLAIN_MACRO", "PLAIN_CLEAN",
]


def _has_vba_project_ole(ole):
    """MS-OVBA: VBA root must contain PROJECT + VBA/_VBA_PROJECT + VBA/dir."""
    names = {"/".join(path).lower() for path in ole.listdir()}
    if "vba/_vba_project" in names and "vba/dir" in names:
        return any(n.endswith("project") and not n.startswith("vba/")
                   for n in names) or "project" in names
    return "vba/_vba_project" in names and "vba/dir" in names


def _has_vba_project_zip(path):
    try:
        with zipfile.ZipFile(path) as zf:
            return any(n.lower().endswith("vbaproject.bin")
                       for n in zf.namelist())
    except (zipfile.BadZipFile, OSError):
        return False


def classify(path, passwords=None):
    passwords = passwords or list(crypto.DEFAULT_PASSWORDS)
    detail = {"file": os.path.basename(path), "evidence": {}}

    # ---- format identification by magic ---------------------------------
    is_ole_magic = olefile.isOleFile(path)
    is_zip_magic = zipfile.is_zipfile(path)
    detail["evidence"]["magic_ole"] = is_ole_magic
    detail["evidence"]["magic_zip"] = is_zip_magic

    # ---- corruption: OLE magic but cannot actually open ------------------
    ole = None
    if is_ole_magic:
        try:
            ole = olefile.OleFileIO(path)
        except Exception as exc:  # noqa: BLE001
            detail["evidence"]["ole_open_error"] = \
                "{}: {}".format(type(exc).__name__, str(exc)[:60])
            detail["state"] = "CORRUPT_CONTAINER"
            return detail
        # opened, but an empty/structurally-broken directory means corruption
        if not ole.listdir():
            detail["evidence"]["ole_open_error"] = "empty directory tree"
            detail["state"] = "CORRUPT_CONTAINER"
            ole.close()
            return detail

    # ---- msoffcrypto is the authoritative encryption parser -------------
    msoff_exc = None
    office = None
    try:
        with open(path, "rb") as fh:
            office = msoffcrypto.OfficeFile(fh)
            encrypted = office.is_encrypted()
    except Exception as exc:  # noqa: BLE001
        msoff_exc = "{}: {}".format(type(exc).__name__, str(exc)[:60])
        encrypted = False
    detail["evidence"]["msoffcrypto_class"] = \
        type(office).__name__ if office is not None else None
    detail["evidence"]["msoffcrypto_exception"] = msoff_exc

    if is_ole_magic:
        # oletools' name-only fallback (what oleid ends up showing)
        name_only = crypto.is_encrypted(path)
        detail["evidence"]["oletools_name_only_encrypted"] = name_only

        if msoff_exc is None and encrypted:
            # well-formed encryption: try the passwords we have
            state = "ENCRYPTED_VALID"
            for pwd in passwords:
                try:
                    with open(path, "rb") as fh:
                        candidate = msoffcrypto.OfficeFile(fh)
                        candidate.load_key(password=pwd)
                        buf = io.BytesIO()
                        candidate.decrypt(buf)
                    if buf.tell() > 0:
                        detail["evidence"]["decrypted_with"] = pwd
                        # a decrypted package is rescan material, not terminal
                        state = "ENCRYPTED_VALID"
                        break
                except Exception:  # noqa: BLE001
                    continue
            else:
                state = "ENCRYPTED_NO_PASSWORD"
            detail["state"] = state
            if ole is not None:
                ole.close()
            return detail

        if msoff_exc is not None and name_only:
            # encryption-like stream names but invalid content -> spoof
            detail["state"] = "ENCRYPTION_SPOOF"
            if ole is not None:
                ole.close()
            return detail

        # OLE that is neither valid encryption nor a spoof (e.g. unrecognized)
        has_macro = _has_vba_project_ole(ole) if ole is not None else False
        detail["state"] = "PLAIN_MACRO" if has_macro else "PLAIN_CLEAN"
        if ole is not None:
            ole.close()
        return detail

    # ---- non-OLE: plain OpenXML/zip or other -----------------------------
    if is_zip_magic:
        has_macro = _has_vba_project_zip(path)
    else:
        has_macro = False
    detail["evidence"]["msoffcrypto_encrypted"] = encrypted
    detail["state"] = "PLAIN_MACRO" if has_macro else "PLAIN_CLEAN"
    return detail


def _print(result):
    print("{:<28} -> {:<24} {}".format(
        result["file"], result["state"], result["evidence"]))


def main():
    args = sys.argv[1:]
    if not args or args == ["--demo"]:
        base = os.path.join(os.path.dirname(HERE), "samples")
        args = [os.path.join(base, n) for n in sorted(os.listdir(base))]
    for path in args:
        _print(classify(path))


if __name__ == "__main__":
    main()
