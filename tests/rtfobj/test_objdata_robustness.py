"""
Security regression tests for rtfobj object extraction.

Each sample contains two objects so the tests verify both that the first
(tricky) object is extracted without missing bytes and that parsing does
not bleed into the second object.
"""

import unittest

from oletools import rtfobj


def parse(data):
    rtfp = rtfobj.RtfObjParser(data)
    rtfp.parse()
    return rtfp.objects


def objdata_start(data, hex_start):
    """offset reported by the parser for the object containing hex_start"""
    return data.rindex(b'\\objdata', 0, hex_start) + len(b'\\objdata')


class TestBinBlock(unittest.TestCase):
    def test_bin_binary_block(self):
        # binary payload containing { } and backslash, which would break
        # group parsing if the \bin length were not honoured exactly
        payload = b'\x7b\x7d\x5c\x41'
        data = (b'{\\rtf1'
                b'{\\object{\\objdata 0105\\bin4 ' + payload + b'0206}}'
                b'{\\object{\\objdata 0307}}'
                b'}')
        objects = parse(data)

        self.assertEqual(len(objects), 2)
        self.assertEqual(objects[0].rawdata,
                         b'\x01\x05' + payload + b'\x02\x06')
        self.assertIsNone(objects[0].bin_error)
        # the second object must be intact and at its own offset
        self.assertEqual(objects[1].rawdata, b'\x03\x07')
        self.assertEqual(objects[1].start,
                         objdata_start(data, data.index(b'0307')))
        self.assertLess(objects[0].end, objects[1].start)


class TestEscapedBraces(unittest.TestCase):
    def test_escaped_braces_and_backslash(self):
        # \{ \} and \\ are control symbols, they must not contribute data
        data = (b'{\\rtf1'
                b'{\\object{\\objdata 0105\\{0206\\}\\\\0307}}'
                b'{\\object{\\objdata 0408}}'
                b'}')
        objects = parse(data)

        self.assertEqual(len(objects), 2)
        self.assertEqual(objects[0].rawdata,
                         b'\x01\x05\x02\x06\x03\x07')
        self.assertEqual(objects[1].rawdata, b'\x04\x08')
        self.assertEqual(objects[1].start,
                         objdata_start(data, data.index(b'0408')))


class TestIgnoredDestinationInGroup(unittest.TestCase):
    def test_ignored_destination_in_group(self):
        # the content of the nested \objname destination (which contains
        # hex-looking letters) must not leak into the object data
        data = (b'{\\rtf1'
                b'{\\object{\\objdata 0105{\\objname NotHexDataZZ}0206}}'
                b'{\\object{\\objdata 0307}}'
                b'}')
        objects = parse(data)

        self.assertEqual(len(objects), 2)
        self.assertEqual(objects[0].rawdata, b'\x01\x05\x02\x06')
        self.assertEqual(objects[1].rawdata, b'\x03\x07')
        self.assertEqual(objects[1].start,
                         objdata_start(data, data.index(b'0307}')))


class TestSplitHex(unittest.TestCase):
    def test_split_hex_content(self):
        # hex digits split by whitespaces/newlines, even in the middle of
        # a byte, must be reassembled without missing bytes
        data = (b'{\\rtf1'
                b'{\\object{\\objdata 0\n1\n05\r\n02 06\t0307}}'
                b'{\\object{\\objdata 0408}}'
                b'}')
        objects = parse(data)

        self.assertEqual(len(objects), 2)
        self.assertEqual(objects[0].rawdata,
                         b'\x01\x05\x02\x06\x03\x07')
        self.assertEqual(objects[1].rawdata, b'\x04\x08')
        self.assertEqual(objects[1].start,
                         objdata_start(data, data.index(b'0408')))


class TestBinLengthBeyondEndOfFile(unittest.TestCase):
    def test_bin_length_exceeding_remaining_data(self):
        # \bin declares 9999 bytes but the file ends long before that;
        # the corrupted object must get a clear failure result and the
        # following object must still be extracted independently
        data = (b'{\\rtf1'
                b'{\\object{\\objdata 0105\\bin9999 XYZ}}'
                b'{\\object{\\objdata 0206}}'
                b'}')
        objects = parse(data)

        self.assertEqual(len(objects), 2)

        corrupted = objects[0]
        # reported offset points at the start of the object data
        self.assertEqual(corrupted.start,
                         objdata_start(data, data.index(b'0105')))
        # the bogus \bin contributed no bytes: only the hex 0105 extracted
        self.assertEqual(corrupted.rawdata, b'\x01\x05')
        self.assertEqual(len(corrupted.rawdata), 2)
        # clear failure result: reason recorded, object not valid OLE
        self.assertIsNotNone(corrupted.bin_error)
        self.assertIn('9999', corrupted.bin_error)
        self.assertFalse(corrupted.is_ole)

        # the following object is still identified independently
        following = objects[1]
        self.assertEqual(following.rawdata, b'\x02\x06')
        self.assertIsNone(following.bin_error)
        self.assertEqual(following.start,
                         objdata_start(data, data.index(b'0206')))
        self.assertLess(corrupted.end, following.start)


if __name__ == '__main__':
    unittest.main()
