"""
build_samples.py - Generate the four small local fixtures for the
encryption-triage lab.

Categories (all synthesized locally, no external downloads):
  1 encrypted_valid   : real, well-formed encrypted OOXML in an OLE container
                        (agile EncryptionInfo + EncryptedPackage). Reuses the
                        repo's own tests/test-data/basic/encrypted.docx so the
                        fixture is guaranteed parseable by msoffcrypto.
  2 fake_named_stream : a structurally VALID OLE/CFB container that contains a
                        stream literally named "EncryptionInfo" (plus a decoy
                        "EncryptedPackage") whose payload is NOT a valid
                        encryption descriptor. This is the "only the stream
                        name looks encrypted" attachment.
  3 corrupted         : a broken container. Two variants:
                          3a magic present but directory/FAT truncated
                          3b valid CFB then header signature zeroed after parse
  4 plain_macro       : a normal, unencrypted macro-enabled OpenXML document
                        with a real VBA project (reuses the repo's own
                        oleform-PR314.docm, which contains VBA forms).
"""
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.dirname(HERE)
SAMPLES = os.path.join(LAB, "samples")
REPO = os.path.dirname(LAB)
sys.path.insert(0, HERE)
from minicfb import CFBWriter  # noqa: E402

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def write(name, data):
    path = os.path.join(SAMPLES, name)
    with open(path, "wb") as fh:
        fh.write(data)
    print("wrote {:32s} {:6d} bytes".format(name, len(data)))
    return path


def main():
    os.makedirs(SAMPLES, exist_ok=True)

    # 1. genuinely encrypted, well-formed OOXML-in-OLE ---------------------
    src_enc = os.path.join(REPO, "tests", "test-data", "basic", "encrypted.docx")
    write("1_encrypted_valid.docx", open(src_enc, "rb").read())

    # 2. valid CFB, fake EncryptionInfo/EncryptedPackage names, junk body --
    #    EncryptionInfo starts with an agile-looking 4.4 version tag but the
    #    rest is plain ASCII, so msoffcrypto parses the version then fails on
    #    the XML. oletools' name-only fallback still flags it encrypted.
    w = CFBWriter()
    w.add_stream("EncryptionInfo", b"\x04\x00\x04\x00" + b"NOT-REAL-ENCRYPTION-DESCRIPTOR" * 4)
    w.add_stream("EncryptedPackage", b"\x00" * 128 + b"FAKE-PAYLOAD" * 8)
    w.add_stream("DecoyDocument", b"harmless decoy stream " * 4)
    write("2_fake_named_stream.bin", w.build())

    # 3a. OLE magic present, but the body is truncated garbage -------------
    #     Keep the 8-byte signature so naive magic checks call it OLE, then
    #     provide far too few bytes for a header/directory/FAT.
    corrupt_a = OLE_MAGIC + b"\x00" * 32 + b"TRUNCATED-CONTAINER" * 6
    write("3a_corrupted_truncated.bin", corrupt_a)

    # 3b. valid CFB signature, but the directory sector / root entry and the
    #     header's first-directory pointer are deliberately corrupted. Magic
    #     stays intact (naive checks call it OLE) yet structural parsing fails.
    w2 = CFBWriter()
    w2.add_stream("SomeStream", b"content " * 16)
    good = bytearray(w2.build())
    # directory occupies the first sector after the 512-byte header.
    # clobber the root-entry name and its child/start fields inside that sector.
    for off in range(512, 512 + 128):
        good[off] = 0xEE
    # also corrupt the "first directory sector" pointer in the header (offset 44)
    good[44:48] = b"\xFF\xFF\xFF\x7F"
    write("3b_corrupted_fat.bin", bytes(good))

    # 4. normal unencrypted macro document --------------------------------
    src_macro = os.path.join(REPO, "tests", "test-data", "oleform",
                             "oleform-PR314.docm")
    write("4_plain_macro.docm", open(src_macro, "rb").read())

    print("\nall fixtures in:", SAMPLES)


if __name__ == "__main__":
    main()
