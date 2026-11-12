"""
Tests for statement-level macro behaviour evidence
(:mod:`oletools.macro_evidence`) and its olevba integration.
"""

import json
import os
import unittest
from os.path import join

from oletools import macro_evidence, olevba
from oletools.macro_evidence import (
    MacroEvidenceAnalyzer, RESOLUTION_CODE, RESOLUTION_STRING,
    RESOLUTION_COMMENT, RESOLUTION_STATIC, RESOLUTION_UNKNOWN)
from tests.test_utils import DATA_BASE_DIR, call_and_capture

EVIDENCE_DIR = join(DATA_BASE_DIR, 'olevba', 'evidence')


def _index(evidence, keyword=None, kind=None):
    """Filter KeywordEvidence objects by keyword and/or context kind."""
    result = evidence
    if keyword is not None:
        result = [item for item in result
                  if item.keyword.lower() == keyword.lower()]
    if kind is not None:
        result = [item for item in result if item.kind == kind]
    return result


def _index_calls(calls, callee):
    return [item for item in calls if item.callee.lower() == callee.lower()]


class TestLexing(unittest.TestCase):
    """Low-level lexer/parser behaviour."""

    def test_comment_string_and_code_separation(self):
        analyzer = MacroEvidenceAnalyzer()
        code = (
            "Sub Demo()\n"
            "    ' Shell only in a comment\n"
            "    Dim x As String\n"
            "    x = \"Shell only in a string\"\n"
            "    Shell \"cmd.exe\"\n"
            "End Sub\n")
        analyzer.add_module('Mod', code)
        evidence = analyzer.keyword_evidence(('Shell',))
        contexts = sorted(item.kind for item in evidence)
        self.assertEqual(contexts,
                         [RESOLUTION_CODE, RESOLUTION_COMMENT,
                          RESOLUTION_STRING])
        code_item = _index(evidence, 'Shell', RESOLUTION_CODE)[0]
        self.assertEqual(code_item.line, 5)
        self.assertIn('Shell "cmd.exe"', code_item.statement)
        self.assertEqual(_index(evidence, 'Shell', RESOLUTION_COMMENT)[0].line,
                         2)
        self.assertEqual(_index(evidence, 'Shell', RESOLUTION_STRING)[0].line,
                         4)

    def test_line_continuation_keeps_first_line_number(self):
        analyzer = MacroEvidenceAnalyzer()
        code = (
            "Sub Demo()\n"
            "    Open path For _\n"
            "        Binary As #1\n"
            "End Sub\n")
        analyzer.add_module('Mod', code)
        evidence = analyzer.keyword_evidence(('Binary', 'Open'))
        # both keywords are on the logical line starting at physical line 2
        for item in evidence:
            self.assertEqual(item.line, 2)
        statements = {item.statement for item in evidence}
        self.assertTrue(any('Binary' in stmt for stmt in statements))

    def test_colon_separated_statements(self):
        analyzer = MacroEvidenceAnalyzer()
        code = "Sub Demo()\n    Kill f : MkDir d\nEnd Sub\n"
        analyzer.add_module('Mod', code)
        evidence = analyzer.keyword_evidence(('Kill', 'MkDir'))
        self.assertEqual({item.line for item in evidence}, {2})
        statements = {item.statement for item in evidence}
        self.assertIn('Kill f', statements)
        self.assertIn('MkDir d', statements)

    def test_quote_does_not_start_comment_inside_string(self):
        analyzer = MacroEvidenceAnalyzer()
        code = "Sub Demo()\n    x = \"it's a Write test\"\nEnd Sub\n"
        analyzer.add_module('Mod', code)
        evidence = analyzer.keyword_evidence(('Write',))
        self.assertEqual(len(evidence), 1)
        self.assertEqual(evidence[0].kind, RESOLUTION_STRING)

    def test_rem_is_treated_as_comment(self):
        analyzer = MacroEvidenceAnalyzer()
        code = ("Rem Shell Write in a rem comment\n"
                "Sub Demo()\n"
                "    x = 1 : Rem Kill MkDir here\n"
                "End Sub\n")
        analyzer.add_module('Mod', code)
        evidence = analyzer.keyword_evidence(
            ('Shell', 'Write', 'Kill', 'MkDir'))
        self.assertTrue(evidence)
        self.assertEqual({item.kind for item in evidence},
                         {RESOLUTION_COMMENT})

    def test_punctuated_keyword_keeps_word_boundary(self):
        analyzer = MacroEvidenceAnalyzer()
        code = "Sub Demo()\n    Print #1, \"x\"\nEnd Sub\n"
        analyzer.add_module('Mod', code)
        evidence = analyzer.keyword_evidence(('Print #',))
        self.assertEqual(len(evidence), 1)
        self.assertEqual(evidence[0].kind, RESOLUTION_CODE)
        self.assertEqual(evidence[0].line, 2)


class TestStatementKeywords(unittest.TestCase):
    """Statement keywords must head a statement to be executable evidence."""

    def test_variable_named_like_statement_is_not_code_evidence(self):
        analyzer = MacroEvidenceAnalyzer()
        code = (
            "Sub Demo()\n"
            "    Dim shell As Object\n"
            "    Set shell = CreateObject(\"WScript.Shell\")\n"
            "    Dim w As Integer\n"
            "    w = 1\n"
            "End Sub\n")
        analyzer.add_module('Mod', code)
        evidence = analyzer.keyword_evidence(('Shell', 'Write'))
        self.assertEqual(_index(evidence, 'Shell', RESOLUTION_CODE), [])
        self.assertEqual(_index(evidence, 'Write', RESOLUTION_CODE), [])
        # the string occurrence of Shell is still reported as a string
        self.assertTrue(_index(evidence, 'Shell', RESOLUTION_STRING))

    def test_statement_head_and_call_prefix_are_code_evidence(self):
        analyzer = MacroEvidenceAnalyzer()
        code = (
            "Sub Demo()\n"
            "    Shell \"a\"\n"
            "    Call Shell(\"b\")\n"
            "End Sub\n")
        analyzer.add_module('Mod', code)
        evidence = analyzer.keyword_evidence(('Shell',))
        code_items = _index(evidence, 'Shell', RESOLUTION_CODE)
        self.assertEqual(sorted(item.line for item in code_items), [2, 3])


class TestProceduresAndCalls(unittest.TestCase):
    """Procedure mapping and static call resolution."""

    def _two_module_analyzer(self):
        module_a = (
            "Sub UniqueTarget()\n"
            "    Write #1, \"go\"\n"
            "End Sub\n"
            "Sub Shared()\n"
            "End Sub\n")
        module_b = (
            "Sub Entry()\n"
            "    UniqueTarget \"x\"\n"
            "    Call Shared\n"
            "    Application.Run \"DynamicName\"\n"
            "    CallByName obj, \"Method\", 1\n"
            "End Sub\n")
        module_c = (
            "Sub Shared()\n"
            "End Sub\n")
        analyzer = MacroEvidenceAnalyzer()
        analyzer.add_module('ModA', module_a)
        analyzer.add_module('ModB', module_b)
        analyzer.add_module('ModC', module_c)
        return analyzer

    def test_procedure_map_locates_every_definition(self):
        analyzer = self._two_module_analyzer()
        procedures = {info.name: info for info in analyzer.all_procedures()}
        self.assertEqual(procedures['UniqueTarget'].module, 'ModA')
        self.assertEqual(procedures['UniqueTarget'].line, 1)
        self.assertEqual(procedures['Entry'].module, 'ModB')

    def test_static_cross_module_call_is_resolved(self):
        analyzer = self._two_module_analyzer()
        calls = _index_calls(analyzer.procedure_calls(), 'UniqueTarget')
        self.assertEqual(len(calls), 1)
        call = calls[0]
        self.assertEqual(call.resolution, RESOLUTION_STATIC)
        self.assertEqual(call.module, 'ModB')
        self.assertEqual(call.line, 2)
        self.assertEqual(call.target_module, 'ModA')
        self.assertEqual(call.target_line, 1)

    def test_same_named_procedure_is_unknown(self):
        analyzer = self._two_module_analyzer()
        calls = _index_calls(analyzer.procedure_calls(), 'Shared')
        self.assertEqual(len(calls), 1)
        call = calls[0]
        self.assertEqual(call.resolution, RESOLUTION_UNKNOWN)
        self.assertIn('2 modules', call.reason)
        self.assertIsNone(call.target_module)

    def test_dynamic_dispatch_is_unknown(self):
        analyzer = self._two_module_analyzer()
        calls = analyzer.procedure_calls()
        for name in ('Run', 'CallByName'):
            matched = _index_calls(calls, name)
            self.assertEqual(len(matched), 1, name)
            self.assertEqual(matched[0].resolution, RESOLUTION_UNKNOWN)
            self.assertIn('runtime', matched[0].reason)

    def test_undeclared_bare_names_are_not_invented(self):
        analyzer = MacroEvidenceAnalyzer()
        analyzer.add_module('Mod', "Sub Demo()\n    Mystery 1\nEnd Sub\n")
        self.assertEqual(analyzer.procedure_calls(), [])


class TestOlevbaIntegration(unittest.TestCase):
    """End-to-end behaviour through VBA_Parser and the olevba CLI."""

    def _parser(self, filename):
        path = join(EVIDENCE_DIR, filename)
        return olevba.VBA_Parser(path)

    def test_real_malware_has_code_evidence_and_static_calls(self):
        parser = self._parser('real_malware.bas')
        report = parser.build_macro_evidence()
        code_evidence = [item for item in report.keyword_evidence
                         if item.kind == RESOLUTION_CODE]
        code_keywords = {item.keyword.lower() for item in code_evidence}
        # the behaviour named in the task must be located in real code
        for expected in ('auto_open', 'createobject', 'open', 'put', 'run'):
            self.assertIn(expected, code_keywords)
        # every item points at the module and a positive line number
        for item in code_evidence:
            self.assertEqual(item.module, 'real_malware')
            self.assertGreater(item.line, 0)
            self.assertTrue(item.statement.strip())
        # static in-module calls are resolved to their definition
        download = _index_calls(report.calls, 'DownloadPayload')
        self.assertEqual(len(download), 1)
        self.assertEqual(download[0].resolution, RESOLUTION_STATIC)
        self.assertEqual(download[0].target_module, 'real_malware')
        # object dispatch stays unknown
        run = _index_calls(report.calls, 'Run')
        self.assertEqual(run[0].resolution, RESOLUTION_UNKNOWN)
        parser.close()

    def test_plain_text_control_has_no_code_evidence(self):
        parser = self._parser('plain_text_control.bas')
        report = parser.build_macro_evidence()
        self.assertEqual(
            [item for item in report.keyword_evidence
             if item.kind == RESOLUTION_CODE], [])
        # but the same words are accounted for as strings/comments
        contexts = {item.kind for item in report.keyword_evidence}
        self.assertTrue({RESOLUTION_STRING, RESOLUTION_COMMENT}
                        <= contexts)
        # the words that classic scanning flags are not silently dropped
        words = {item.keyword.lower() for item in report.keyword_evidence}
        for expected in ('shell', 'write', 'auto_open'):
            self.assertIn(expected, words)
        parser.close()

    def test_report_is_json_serializable(self):
        parser = self._parser('real_malware.bas')
        report = parser.build_macro_evidence().as_dict()
        # round-trip must not raise
        json.loads(json.dumps(report))
        self.assertIn('keyword_evidence', report)
        self.assertIn('procedure_calls', report)
        parser.close()

    def test_cli_evidence_flag_text_output(self):
        path = join(EVIDENCE_DIR, 'real_malware.bas')
        output, ret_code = call_and_capture(
            'olevba', args=(path, '--evidence'))
        self.assertEqual(ret_code, 0)
        self.assertIn('MACRO BEHAVIOUR EVIDENCE', output)
        self.assertIn('Keyword evidence in executable code', output)
        self.assertIn('Procedure calls', output)

    def test_cli_evidence_flag_json_output(self):
        path = join(EVIDENCE_DIR, 'plain_text_control.bas')
        output, ret_code = call_and_capture(
            'olevba', args=(path, '--evidence', '-j'), exclude_stderr=True)
        self.assertEqual(ret_code, 0)
        data = json.loads(output)
        report = None
        for entry in data:
            if 'macro_evidence' in entry:
                report = entry['macro_evidence']
        self.assertIsNotNone(report)
        self.assertEqual(
            [item for item in report['keyword_evidence']
             if item['context'] == RESOLUTION_CODE], [])

    def test_without_evidence_flag_keeps_legacy_output(self):
        path = join(EVIDENCE_DIR, 'real_malware.bas')
        output, ret_code = call_and_capture('olevba', args=(path, ))
        self.assertEqual(ret_code, 0)
        self.assertNotIn('MACRO BEHAVIOUR EVIDENCE', output)


if __name__ == '__main__':
    unittest.main()
