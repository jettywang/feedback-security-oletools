"""
triage.py - Collect the exact branch evidence oletools produces for each
fixture. For every sample it records:

  * magic: olefile.isOleFile / zipfile.is_zipfile (format identification)
  * ftguess: container + filetype (the format-recognition branch)
  * msoffcrypto: concrete class, is_encrypted(), or the raised exception
  * crypto.is_encrypted: oletools' final boolean (name-only fallback included)
  * oleid: the 'Encrypted' and 'VBA Macros' indicator values
  * olevba: process exit code (0 ok, 8 unexpected, 9 encrypted w/ wrong pw, ...)

The point is to show *what evidence backs each conclusion* and therefore which
attachments must be decrypted and re-scanned versus treated as malformed.
"""
import json
import os
import subprocess
import sys
import zipfile

import olefile

HERE = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.dirname(HERE)
SAMPLES = os.path.join(LAB, "samples")
RESULTS = os.path.join(LAB, "results")
sys.path.insert(0, os.path.dirname(LAB))  # import oletools from repo

from oletools import crypto, ftguess, oleid  # noqa: E402
import msoffcrypto  # noqa: E402

PY = sys.executable


def msoffcrypto_probe(path):
    try:
        with open(path, "rb") as fh:
            office = msoffcrypto.OfficeFile(fh)
            return {
                "class": type(office).__name__,
                "is_encrypted": office.is_encrypted(),
                "exception": None,
            }
    except Exception as exc:  # noqa: BLE001 - we want every failure type
        return {"class": None, "is_encrypted": None,
                "exception": "{}: {}".format(type(exc).__name__, str(exc)[:70])}


def oleid_probe(path):
    out = {"encrypted": None, "encrypted_desc": None,
           "vba_macros": None, "xlm_macros": None,
           "file_format": None, "container": None}
    try:
        indicators = oleid.OleID(path).check()
        for ind in indicators:
            if ind.id == "encrypted":
                out["encrypted"] = ind.value
                out["encrypted_desc"] = ind.description
            elif ind.id == "vba":
                out["vba_macros"] = ind.value
            elif ind.id == "xlm":
                out["xlm_macros"] = ind.value
            elif ind.id == "ftype":
                out["file_format"] = ind.value
            elif ind.id == "container":
                out["container"] = ind.value
    except Exception as exc:  # noqa: BLE001
        out["oleid_exception"] = "{}: {}".format(type(exc).__name__, str(exc)[:70])
    return out


def olevba_exit_code(path):
    proc = subprocess.run([PY, "-m", "oletools.olevba", path],
                          capture_output=True, text=True)
    macro_modules = [ln.strip() for ln in proc.stdout.splitlines()
                     if ln.startswith("VBA MACRO")]
    signal = None
    for line in proc.stdout.splitlines():
        if "No VBA or XLM macros" in line:
            signal = line.strip()
            break
    # keep only traceback's final exception line, ignore unrelated parser noise
    err_lines = [ln for ln in proc.stderr.splitlines()
                 if ln.strip().endswith("Error") or "Error" in ln
                 or "Exception" in ln]
    return {"returncode": proc.returncode,
            "macro_modules_found": len(macro_modules),
            "signal": signal,
            "stderr_tail": err_lines[-1][:80] if err_lines else None}


def triage_one(path):
    with open(path, "rb") as fh:
        head = fh.read(8)
    row = {"sample": os.path.basename(path), "size": os.path.getsize(path),
           "magic_ole": olefile.isOleFile(path),
           "magic_zip": zipfile.is_zipfile(path),
           "header_hex": head.hex()}
    try:
        guesser = ftguess.FileTypeGuesser(filepath=path)
        row["ft_container"] = guesser.container
        row["ft_filetype"] = guesser.filetype
    except Exception as exc:  # noqa: BLE001
        row["ft_container"] = "ERROR"
        row["ft_filetype"] = "{}: {}".format(type(exc).__name__, str(exc)[:50])
    row["msoffcrypto"] = msoffcrypto_probe(path)
    try:
        row["crypto_is_encrypted"] = crypto.is_encrypted(path)
    except Exception as exc:  # noqa: BLE001
        row["crypto_is_encrypted"] = "RAISED {}: {}".format(
            type(exc).__name__, str(exc)[:50])
    row["oleid"] = oleid_probe(path)
    row["olevba"] = olevba_exit_code(path)
    return row


def main():
    os.makedirs(RESULTS, exist_ok=True)
    rows = []
    for name in sorted(os.listdir(SAMPLES)):
        rows.append(triage_one(os.path.join(SAMPLES, name)))
    out_json = os.path.join(RESULTS, "triage.json")
    with open(out_json, "w") as fh:
        json.dump(rows, fh, indent=2, ensure_ascii=False)
    print(json.dumps(rows, indent=2, ensure_ascii=False))
    print("\nwrote", out_json)


if __name__ == "__main__":
    main()
