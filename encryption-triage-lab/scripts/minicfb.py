"""
minicfb.py - Minimal, dependency-free Compound File Binary (CFB/OLE) writer.

Only supports the subset needed to build small test fixtures:
  - 512-byte sectors, 32-bit DIFAT (v3)
  - a flat hierarchy (root storage + leaf streams), no sub-storages
  - little-endian, UTF-16LE directory names

The produced files open correctly in olefile (valid header, FAT, directory,
mini-FAT and mini-stream), which is what oletools/msoffcrypto rely on.

This is intentionally NOT a general-purpose writer: it targets tiny fixtures
where every stream fits in the 64 KB mini-stream and the whole file fits in
one header DIFAT array.
"""
import struct

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_CUTOFF = 4096
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF
STGTY_STORAGE = 1
STGTY_STREAM = 2


class CFBWriter:
    def __init__(self):
        # name -> bytes
        self.streams = {}

    def add_stream(self, name, data):
        self.streams[name] = data

    def _pad(self, blob, size):
        if len(blob) % size:
            blob += b"\x00" * (size - (len(blob) % size))
        return blob

    def build(self):
        names = list(self.streams.keys())

        # ---- 1. Lay out the mini stream (each stream padded to 64B) -------
        mini_blob = bytearray()
        mini_chain = []          # mini-sector indices per stream
        for name in names:
            data = self.streams[name]
            start = len(mini_blob) // MINI_SECTOR_SIZE
            mini_blob += data
            mini_blob = bytearray(self._pad(bytes(mini_blob), MINI_SECTOR_SIZE))
            n_mini = max(1, (len(data) + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE)
            mini_chain.append((start, n_mini, len(data)))
        # mini stream itself padded to full 512B sectors later
        mini_blob = bytes(self._pad(bytes(mini_blob), SECTOR_SIZE))

        # ---- 2. Directory: root + one entry per stream --------------------
        n_dir_entries = 1 + len(names)
        # directory occupies whole 512B sectors (4 entries each), min 1
        dir_sector_count = max(1, (n_dir_entries + 3) // 4)

        # ---- 3. Decide sector layout (sector 0 is first after header) -----
        # Order (directory pinned to sector 0 for a simple, valid v3 file):
        #   directory, mini-FAT, mini-stream, FAT sectors
        mini_stream_sectors = len(mini_blob) // SECTOR_SIZE
        # mini FAT: one entry per mini-sector; each FAT sector holds 128
        n_mini_sectors_total = len(mini_blob) // MINI_SECTOR_SIZE
        minifat_sectors = max(1, (n_mini_sectors_total + 127) // 128)

        # We do not know FAT sector count until we know total sectors.
        # Total regular sectors = dir + minifat + ministream + fat.
        # Iterate to a fixed point (FAT must describe every sector incl. itself).
        fat_count = 1
        while True:
            total = dir_sector_count + minifat_sectors + mini_stream_sectors + fat_count
            needed_fat = max(1, (total + 127) // 128)
            if needed_fat == fat_count:
                break
            fat_count = needed_fat
        total_sectors = dir_sector_count + minifat_sectors + mini_stream_sectors + fat_count

        # sector indices
        dir_secs = list(range(0, dir_sector_count))
        minifat_start = dir_sector_count
        minifat_secs = list(range(minifat_start, minifat_start + minifat_sectors))
        ministream_start = minifat_start + minifat_sectors
        ministream_secs = list(range(ministream_start,
                                     ministream_start + mini_stream_sectors))
        fat_start = ministream_start + mini_stream_sectors
        fat_secs = list(range(fat_start, fat_start + fat_count))

        # ---- 4. Build FAT --------------------------------------------------
        fat = [FREESECT] * (fat_count * 128)
        for i, s in enumerate(dir_secs):
            fat[s] = ENDOFCHAIN if i == len(dir_secs) - 1 else dir_secs[i + 1]
        for i, s in enumerate(minifat_secs):
            fat[s] = ENDOFCHAIN if i == len(minifat_secs) - 1 else minifat_secs[i + 1]
        for i, s in enumerate(ministream_secs):
            fat[s] = ENDOFCHAIN if i == len(ministream_secs) - 1 else ministream_secs[i + 1]
        for s in fat_secs:
            fat[s] = FATSECT

        # ---- 5. Build mini-FAT (chains of 64B mini-sectors) ---------------
        minifat = [ENDOFCHAIN] * (minifat_sectors * 128)
        for (start, n_mini, _real_len) in mini_chain:
            for k in range(n_mini):
                idx = start + k
                minifat[idx] = ENDOFCHAIN if k == n_mini - 1 else idx + 1

        # ---- 6. Directory entries -----------------------------------------
        def dir_entry(name, etype, color, left, right, child,
                      start, size, clsid=b"\x00" * 16):
            name_utf = name.encode("utf-16-le")
            name_field = name_utf + b"\x00\x00"
            if len(name_field) > 64:
                raise ValueError("name too long")
            name_len = len(name_utf) + 2
            out = name_field + b"\x00" * (64 - len(name_field))  # 64  name
            out += struct.pack("<H", name_len)                   # 64  name length
            out += struct.pack("<B", etype)                      # 66  object type
            out += struct.pack("<B", color)                      # 67  color flag
            out += struct.pack("<I", left)                       # 68  left sibling
            out += struct.pack("<I", right)                      # 72  right sibling
            out += struct.pack("<I", child)                      # 76  child
            out += clsid                                          # 80  CLSID (16)
            out += struct.pack("<I", 0)                          # 96  state flags
            out += b"\x00" * 8                                   # 100 creation time
            out += b"\x00" * 8                                   # 108 modified time
            out += struct.pack("<I", start)                      # 116 starting sector
            out += struct.pack("<I", size)                       # 120 stream size (low)
            assert len(out) == 124, len(out)
            out += b"\x00" * 4                                   # 124 stream size (high, v3=0)
            assert len(out) == 128, len(out)
            return out
        entries = []

        # Directory entries under the root storage are stored as a balanced
        # binary search tree keyed by upper-case UTF-16 name (MS-CFB 2.6.4).
        # olefile walks left/right sibling pointers from root's child, so a
        # flat list with NOSTREAM siblings is invisible. Build a balanced BST.
        # item: (upper-name sort key, original name, mini-start, mini-count, size)
        items = [(n.upper(), n) + mini_chain[i] for i, n in enumerate(names)]
        items.sort(key=lambda t: t[0])
        # maps item index -> (left sid, right sid, color); sid is 1-based (root=0)
        link = {}

        def build_bst(lo, hi):
            if lo > hi:
                return NOSTREAM
            mid = (lo + hi) // 2
            sid = mid + 1  # directory SID (root storage occupies SID 0)
            left_sid = build_bst(lo, mid - 1)
            right_sid = build_bst(mid + 1, hi)
            # color: root of this subtree black (0); a fully balanced tree lets
            # us mark all nodes black and still satisfy traversal needs.
            link[mid] = (left_sid, right_sid, 0)
            return sid

        tree_root_sid = build_bst(0, len(items) - 1) if items else NOSTREAM

        # Root storage: child = BST root sid, start = mini-stream start
        root = dir_entry("Root Entry", STGTY_STORAGE, 0, NOSTREAM, NOSTREAM,
                         tree_root_sid,
                         ministream_start if mini_stream_sectors else ENDOFCHAIN,
                         len(mini_blob))
        entries.append(root)

        for mid, item in enumerate(items):
            _key, name, start, _n_mini, real_len = item
            left_sid, right_sid, color = link[mid]
            entries.append(dir_entry(name, STGTY_STREAM, color, left_sid,
                                     right_sid, NOSTREAM, start, real_len))
        # pad directory to fill its sectors
        while len(entries) < dir_sector_count * 4:
            entries.append(b"\x00" * 128)
        dir_blob = b"".join(entries)
        dir_blob = self._pad(dir_blob, SECTOR_SIZE)

        # ---- 7. Header -----------------------------------------------------
        header = bytearray(SECTOR_SIZE)
        header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
        struct.pack_into("<HH", header, 24, 0x003e, 0x0003)  # minor, major (v3)
        struct.pack_into("<H", header, 26, 0x0003)
        struct.pack_into("<H", header, 28, 0xfffe)            # byte order
        struct.pack_into("<H", header, 30, 9)                 # sector shift 512
        struct.pack_into("<H", header, 32, 6)                 # mini sector shift 64
        struct.pack_into("<I", header, 40, 0)                 # number DIFAT sectors (v3=0)
        struct.pack_into("<I", header, 44, dir_secs[0])       # first directory sector
        struct.pack_into("<I", header, 48, 0)                 # transaction sig
        struct.pack_into("<I", header, 52, MINI_CUTOFF)       # mini stream cutoff 4096
        struct.pack_into("<I", header, 56, minifat_secs[0])   # first mini-FAT sector
        struct.pack_into("<I", header, 60, minifat_sectors)   # mini-FAT sector count
        struct.pack_into("<I", header, 64, ENDOFCHAIN)        # first DIFAT sector
        struct.pack_into("<I", header, 68, 0)                 # DIFAT sector count
        # DIFAT array at offset 76, 109 entries
        for i in range(109):
            val = fat_secs[i] if i < len(fat_secs) else FREESECT
            struct.pack_into("<I", header, 76 + i * 4, val)

        # ---- 8. Serialize sectors (order matches the FAT layout) ----------
        fat_blob = b"".join(struct.pack("<I", v) for v in fat)
        fat_blob = self._pad(fat_blob, SECTOR_SIZE)
        minifat_blob = b"".join(struct.pack("<I", v) for v in minifat)
        minifat_blob = self._pad(minifat_blob, SECTOR_SIZE)

        out = bytes(header) + dir_blob \
            + minifat_blob[:minifat_sectors * SECTOR_SIZE] \
            + mini_blob + fat_blob[:fat_count * SECTOR_SIZE]
        return out


if __name__ == "__main__":
    import sys
    w = CFBWriter()
    w.add_stream("Hello", b"Hi from minicfb " * 10)
    w.add_stream("EncryptionInfo", b"\x04\x00\x04\x00" + b"junk" * 20)
    data = w.build()
    out = sys.argv[1] if len(sys.argv) > 1 else "out.bin"
    with open(out, "wb") as fh:
        fh.write(data)
    print("wrote", out, len(data), "bytes")
