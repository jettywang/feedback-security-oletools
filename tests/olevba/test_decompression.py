"""
Tests for the hardened MS-OVBA decompression path in olevba
(decompress_stream) and the decompression budget shared across the
VBA modules of one document.

All compressed samples are synthesized here (a few hundred bytes at most),
so no binary test fixtures are required.
"""

import struct
import unittest
from io import BytesIO

from oletools import olevba
from oletools.olevba import (
    decompress_stream, DecompressionBudget, DecompressionError,
    DecompressionBudgetError, VBA_Project,
)


# --------------------------------------------------------------------------
# helpers to build MS-OVBA (section 2.4.1) compressed containers by hand
# --------------------------------------------------------------------------

SIGNATURE = b'\x01'
CHUNK_SIGNATURE = 0b011


def compressed_chunk_header(data_length):
    """Header for a compressed chunk whose data is data_length bytes."""
    size_field = data_length + 2 - 3
    value = (1 << 15) | (CHUNK_SIGNATURE << 12) | size_field
    return struct.pack('<H', value)


def raw_chunk_header():
    """Header for an uncompressed (raw) chunk of 4096 data bytes."""
    value = (0 << 15) | (CHUNK_SIGNATURE << 12) | (4098 - 3)
    return struct.pack('<H', value)


def literal_sequence(literals):
    """TokenSequences containing only literal bytes (8 per flag byte)."""
    out = bytearray()
    for i in range(0, len(literals), 8):
        group = literals[i:i + 8]
        out.append(0x00)
        out.extend(group)
    return bytes(out)


def copy_token(offset, length, emitted=3):
    """
    Build a TokenSequence whose literals (emitted plain bytes, all 'A') are
    followed on the next flag-bit slot by a CopyToken. The CopyToken Help bit
    count is derived from emitted, the number of bytes already in the chunk.
    """
    import math
    bit_count = max(int(math.ceil(math.log(emitted, 2))), 4)
    length_mask = 0xFFFF >> bit_count
    offset_field = (offset - 1) << (16 - bit_count)
    length_field = (length - 3) & length_mask
    token = offset_field | length_field
    # the copy token occupies bit (emitted mod 8) of its flag byte
    flag_byte = 1 << (emitted % 8)
    return bytes([flag_byte]) + struct.pack('<H', token)


def literals_then_copy(emitted, offset, length, fill=b'A'):
    """
    A chunk body: emitted literal bytes followed by a CopyToken, all within a
    single TokenSequence (requires emitted < 8). The CopyToken Help bit count
    is derived from emitted, the number of bytes already in the chunk.
    """
    import math
    if emitted >= 8:
        raise ValueError('this helper only supports fewer than 8 literals')
    bit_count = max(int(math.ceil(math.log(emitted, 2))), 4)
    length_mask = 0xFFFF >> bit_count
    offset_field = (offset - 1) << (16 - bit_count)
    length_field = (length - 3) & length_mask
    token = struct.pack('<H', offset_field | length_field)
    # the copy token occupies the bit right after the literals
    flag_byte = 1 << emitted
    return bytes([flag_byte]) + fill * emitted + token


def make_container(chunk_data, raw=False):
    if raw:
        return SIGNATURE + raw_chunk_header() + chunk_data
    return SIGNATURE + compressed_chunk_header(len(chunk_data)) + chunk_data


class TestDecompressStreamHardening(unittest.TestCase):
    """Direct tests of the chunk decoder."""

    def test_literal_round_trip(self):
        """Plain literal tokens are decoded unchanged."""
        payload = b'Attribute VB_Name = "M1"'
        container = make_container(literal_sequence(payload))
        self.assertEqual(decompress_stream(container), payload)

    def test_valid_copy_token_within_chunk(self):
        """A back-reference inside the current chunk expands correctly."""
        chunk = literals_then_copy(emitted=3, offset=3, length=13)
        out = decompress_stream(make_container(chunk))
        self.assertEqual(out, b'A' * 16)
        self.assertLess(len(make_container(chunk)), 40)

    def test_backreference_before_chunk_start_rejected(self):
        """
        A CopyToken whose offset points before the start of the current
        DecompressedChunk must be reported as a corrupt stream, not read from
        another chunk.
        """
        # first chunk emits 4 bytes; second chunk emits 3 bytes then tries to
        # reference offset 8, i.e. 4 bytes before the second chunk start
        first = literal_sequence(b'ABCD')
        bad_second = literals_then_copy(emitted=3, offset=8, length=5)
        container = (SIGNATURE
                     + compressed_chunk_header(len(first)) + first
                     + compressed_chunk_header(len(bad_second)) + bad_second)
        with self.assertRaises(DecompressionError):
            decompress_stream(container)

    def test_declared_chunk_size_mismatch_rejected(self):
        """A chunk header declaring a size beyond the data is corrupt."""
        header = struct.pack('<H', (1 << 15) | (CHUNK_SIGNATURE << 12) | 100)
        container = SIGNATURE + header + b'\x00ABC'
        with self.assertRaises(DecompressionError):
            decompress_stream(container)

    def test_truncated_chunk_header_rejected(self):
        """A stream ending in the middle of a chunk header is corrupt."""
        with self.assertRaises(DecompressionError):
            decompress_stream(SIGNATURE + b'\x03')

    def test_invalid_signature_rejected(self):
        with self.assertRaises(DecompressionError):
            decompress_stream(b'\x02' + b'\x00' * 4)

    def test_empty_stream_rejected(self):
        with self.assertRaises(DecompressionError):
            decompress_stream(b'')

    def test_raw_chunk_requires_full_4096_bytes(self):
        """A raw (uncompressed) chunk must carry exactly 4096 data bytes."""
        header = raw_chunk_header()
        # declare 4098-byte chunk but provide fewer data bytes
        container = SIGNATURE + header + b'X' * 100
        with self.assertRaises(DecompressionError):
            decompress_stream(container)

    def test_raw_chunk_valid(self):
        payload = bytes(bytearray((i % 256 for i in range(4096))))
        out = decompress_stream(make_container(payload, raw=True))
        self.assertEqual(out, payload)

    def test_small_repeated_block_stays_within_chunk_limit(self):
        """
        Repeated small CopyTokens cannot force one chunk past 4096 bytes:
        a token that would overflow the chunk is rejected as corrupt.
        """
        chunk = literals_then_copy(emitted=3, offset=3, length=4094)
        with self.assertRaises(DecompressionError):
            decompress_stream(make_container(chunk))

    def test_truncated_copy_token_rejected(self):
        """A CopyToken header with only one byte left is corrupt."""
        # flag says copy token, but only one token byte follows
        chunk = bytes([0x01, 0x09])
        with self.assertRaises(DecompressionError):
            decompress_stream(make_container(chunk))


class TestDecompressionBudget(unittest.TestCase):
    """Tests of the shared decompression byte budget."""

    def _payload(self, size):
        return b'A' * size

    def test_budget_allows_data_within_limit(self):
        payload = self._payload(64)
        container = make_container(literal_sequence(payload))
        budget = DecompressionBudget(128)
        self.assertEqual(
            decompress_stream(container, budget=budget), payload)
        self.assertEqual(budget.used, 64)
        self.assertFalse(budget.exhausted)

    def test_budget_exceeded_raises_budget_error(self):
        payload = self._payload(64)
        container = make_container(literal_sequence(payload))
        budget = DecompressionBudget(32)
        with self.assertRaises(DecompressionBudgetError) as ctx:
            decompress_stream(container, budget=budget,
                              stream_path='VBA/Mod2', module_name='Mod2',
                              remaining_modules=['Mod3'])
        self.assertTrue(budget.exhausted)
        self.assertIn('Mod3', str(ctx.exception))
        self.assertEqual(ctx.exception.module_name, 'Mod2')

    def test_budget_distinguishes_corruption_from_exhaustion(self):
        # corrupt stream: DecompressionError, budget not marked exhausted
        bad = make_container(bytes([0x01, 0x09]))
        budget = DecompressionBudget(1024)
        with self.assertRaises(DecompressionError) as ctx:
            decompress_stream(bad, budget=budget)
        self.assertNotIsInstance(ctx.exception, DecompressionBudgetError)
        self.assertFalse(budget.exhausted)

    def test_budget_is_shared_across_streams(self):
        """Several legal streams together exhaust one shared budget."""
        payload = self._payload(40)
        container = make_container(literal_sequence(payload))
        budget = DecompressionBudget(100)
        decompress_stream(container, budget=budget, stream_path='VBA/M1')
        self.assertEqual(budget.used, 40)
        # a second legal 40-byte stream would bring the total to 80 (allowed)
        decompress_stream(container, budget=budget, stream_path='VBA/M2')
        self.assertEqual(budget.used, 80)
        # the third one crosses the 100-byte total
        with self.assertRaises(DecompressionBudgetError):
            decompress_stream(container, budget=budget, stream_path='VBA/M3')
        self.assertEqual(len(budget.decompressed_streams), 2)


# --------------------------------------------------------------------------
# minimal fake OLE structure to drive VBA_Project / parse_modules end to end
# --------------------------------------------------------------------------

def _record(record_id, payload):
    return struct.pack('<HL', record_id, len(payload)) + payload


def build_dir_stream(module_specs):
    """
    Build the (uncompressed-then-compressed) dir stream describing the given
    modules. module_specs is a list of (name, stream_name) tuples.
    """
    parts = []
    parts.append(_record(0x0001, struct.pack('<L', 1)))          # PROJECTSYSKIND
    parts.append(_record(0x0002, struct.pack('<L', 0x409)))      # PROJECTLCID
    parts.append(_record(0x0014, struct.pack('<L', 0x409)))      # PROJECTLCIDINVOKE
    parts.append(_record(0x0003, struct.pack('<H', 1252)))       # PROJECTCODEPAGE
    parts.append(_record(0x0004, b'TestProject'))                # PROJECTNAME
    # PROJECTDOCSTRING: record + reserved + unicode size (0)
    parts.append(_record(0x0005, b''))
    parts.append(struct.pack('<H', 0x0040) + struct.pack('<L', 0))
    # PROJECTHELPFILEPATH: record + reserved + duplicate size (0)
    parts.append(_record(0x0006, b''))
    parts.append(struct.pack('<H', 0x003D) + struct.pack('<L', 0))
    parts.append(_record(0x0007, struct.pack('<L', 0)))          # PROJECTHELPCONTEXT
    parts.append(_record(0x0008, struct.pack('<L', 0)))          # PROJECTLIBFLAGS
    # PROJECTVERSION is not a standard Id+Size record: Id + Reserved(=4)
    # + VersionMajor(4) + VersionMinor(2), 12 bytes in total
    parts.append(struct.pack('<HLLH', 0x0009, 4, 0, 0))
    # PROJECTCONSTANTS: record + reserved + unicode size (0)
    parts.append(_record(0x000C, b''))
    parts.append(struct.pack('<H', 0x003C) + struct.pack('<L', 0))
    # 0x000F marks the end of the REFERENCE array; the VBA_Project reference
    # loop consumes this id, so parse_modules starts directly at PROJECTMODULES
    parts.append(struct.pack('<H', 0x000F))

    # PROJECTMODULES: Size(=2) + Count (the 0x000F id is the shared marker)
    parts.append(struct.pack('<LH', 2, len(module_specs)))
    # cookie record; its last read is the 2 cookie bytes, so the first
    # MODULENAME id follows immediately
    parts.append(_record(0x0013, struct.pack('<H', 0)))

    for module_index, (name, stream_name) in enumerate(module_specs):
        name_b = name.encode('cp1252')
        stream_b = stream_name.encode('cp1252')
        stream_u = stream_name.encode('utf-16-le')
        # MODULENAME id + size + payload. After the payload the parser reads
        # the next section id, which is the MODULESTREAMNAME id below.
        parts.append(struct.pack('<HL', 0x0019, len(name_b)))
        parts.append(name_b)
        # MODULESTREAMNAME: id + size + payload, then Reserved + unicode size
        parts.append(struct.pack('<HL', 0x001A, len(stream_b)))
        parts.append(stream_b)
        parts.append(struct.pack('<H', 0x0032))
        parts.append(struct.pack('<L', len(stream_u)))
        parts.append(stream_u)
        # after the unicode name the parser reads MODULEOFFSET id (0x0031),
        # then its size and TextOffset
        parts.append(struct.pack('<HLL', 0x0031, 4, 0))
        # parser then reads MODULETYPE id (0x0021) + its Reserved
        parts.append(struct.pack('<HL', 0x0021, 0))
        # parser then reads the MODULE terminator id (0x002B) + its Reserved
        parts.append(struct.pack('<HL', 0x002B, 0))

    raw_dir = b''.join(parts)
    return _compress_simple(raw_dir)


def _compress_simple(data):
    """Compress data using only literal tokens, one chunk per <=4096 bytes."""
    out = bytearray(b'\x01')
    for offset in range(0, len(data), 4096):
        block = data[offset:offset + 4096]
        body = bytearray()
        for i in range(0, len(block), 8):
            group = block[i:i + 8]
            body.append(0x00)
            body.extend(group)
        out.extend(compressed_chunk_header(len(body)))
        out.extend(body)
    return bytes(out)


class FakeStream(object):
    def __init__(self, data):
        self._buf = BytesIO(data)

    def read(self, size=None):
        if size is None:
            return self._buf.read()
        return self._buf.read(size)

    def __iter__(self):
        return iter(self._buf)


class FakeOle(object):
    """Minimal olefile-like object supporting openstream(path)."""

    def __init__(self, streams):
        self._streams = streams

    def openstream(self, path):
        # the PROJECT text stream is not needed for module extraction
        return FakeStream(self._streams.get(path, b''))


class TestMultiModuleBudget(unittest.TestCase):
    """End-to-end behaviour across several legal/corrupt VBA modules."""

    def _build_project(self, module_streams, budget):
        specs = [(name, name) for name, _ in module_streams]
        streams = {'VBA/dir': build_dir_stream(specs)}
        for name, compressed in module_streams:
            streams['VBA/' + name] = compressed
        ole = FakeOle(streams)
        project = VBA_Project(ole, '', 'PROJECT', 'VBA/dir',
                              decompression_budget=budget)
        project.parse_project_stream()
        return project

    def _literal_stream(self, text):
        return make_container(literal_sequence(text.encode('cp1252')))

    def test_several_legal_modules_exhaust_total_budget(self):
        streams = [
            ('M1', self._literal_stream('A' * 30)),
            ('M2', self._literal_stream('B' * 30)),
            ('M3', self._literal_stream('C' * 30)),
        ]
        project = self._build_project(streams, DecompressionBudget(70))
        extracted = list(project.parse_modules())
        # first module (30 bytes) fits, second would bring the total to 60,
        # third crosses the 70-byte budget
        self.assertEqual([name for _, name, _ in extracted],
                         ['M1.vba', 'M2.vba'])
        self.assertTrue(project.budget_exceeded)
        self.assertIn('M3', project.unchecked_modules)
        # findings from the first two modules are retained
        self.assertEqual(len(extracted), 2)

    def test_corrupt_module_does_not_stop_other_modules(self):
        good_m1 = self._literal_stream('A' * 10)
        corrupt = make_container(bytes([0x01, 0x09]))  # truncated copy token
        good_m3 = self._literal_stream('C' * 10)
        streams = [('M1', good_m1), ('M2', corrupt), ('M3', good_m3)]
        project = self._build_project(streams, DecompressionBudget(1024))
        extracted = list(project.parse_modules())
        names = [name for _, name, _ in extracted]
        self.assertEqual(names, ['M1.vba', 'M3.vba'])
        self.assertEqual(project.corrupt_modules, ['M2'])
        self.assertFalse(project.budget_exceeded)
        statuses = {m.name_str: m.parse_status for m in project.modules}
        self.assertEqual(statuses,
                         {'M1': 'ok', 'M2': 'corrupt', 'M3': 'ok'})

    def test_report_distinguishes_budget_from_corruption(self):
        good = self._literal_stream('X' * 20)
        streams = [('M1', good), ('M2', good), ('M3', good)]
        project = self._build_project(streams, DecompressionBudget(30))
        list(project.parse_modules())
        # M1 fits (20), M2 exceeds -> M2 and M3 unchecked
        self.assertTrue(project.budget_exceeded)
        self.assertEqual(project.corrupt_modules, [])
        self.assertIn('M2', project.unchecked_modules)
        self.assertIn('M3', project.unchecked_modules)
        self.assertIsInstance(project.budget_error, DecompressionBudgetError)


if __name__ == '__main__':
    unittest.main()
