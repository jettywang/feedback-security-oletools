"""
Static, statement-level evidence for VBA macro behaviour.

This module complements the keyword-based detection in
:mod:`oletools.olevba`.  The classic scanner only reports that a keyword such
as ``AutoExec``, ``Write`` or ``Execute`` appears somewhere in the merged
source code, without saying:

* whether the keyword is actual code, a string literal or just a comment;
* which module and which statement contains it;
* whether a called procedure is defined in another module or only resolved at
  runtime (dynamic dispatch).

The analysis here stays deliberately small and purely static.  Anything that
depends on runtime values (``CallByName``, ``Application.Run``, objects whose
type is not known, ...) is reported with the resolution status
:data:`RESOLUTION_UNKNOWN` instead of being guessed.
"""

from __future__ import print_function

import re

# Resolution statuses used for keyword evidence and procedure calls.
RESOLUTION_CODE = 'code'
RESOLUTION_STRING = 'string'
RESOLUTION_COMMENT = 'comment'
RESOLUTION_STATIC = 'static'
RESOLUTION_UNKNOWN = 'unknown'

#: statement kinds that cannot be resolved without executing the macro.
DYNAMIC_DISPATCH = ('CallByName', 'Run', 'OnTime', 'OnAction')

#: VBA built-in statements among the olevba keyword tables.  These are only
#: executable evidence when they head a statement (or follow "Call"); using the
#: same word as a variable or member name is not an invocation.
STATEMENT_KEYWORDS = frozenset((
    'open', 'write', 'put', 'output', 'binary', 'kill', 'mkdir', 'chdir',
    'chdrive', 'shell', 'lib', 'sendkeys', 'appactivate', 'beep', 'print #',
    'filecopy', 'set', 'get', 'close', 'load', 'unload', 'erase', 'randomize',
    'call', 'declare',
))

# A VBA identifier: letters/digits/underscore, not starting with a digit.
IDENTIFIER_RE = re.compile(r'[A-Za-z_][A-Za-z0-9_]*')

# Sub/Function/Property declarations, optionally scoped and/or static.
# Examples:
#   Sub Auto_Open()
#   Private Function DoStuff(x As Integer) As String
#   Public Property Get Value() As Long
PROCEDURE_RE = re.compile(
    r'(?im)^\s*(?:(?P<scope>public|private|friend|global)\s+)?'
    r'(?:static\s+)?(?P<kind>sub|function|property\s+(?:get|let|set))\s+'
    r'(?P<name>[A-Za-z_][A-Za-z0-9_]*)')

# A procedure *call*: an identifier at the start of a statement or after a
# colon (multi-statement line), or following the optional "Call" keyword.
# Member access (obj.Method) and built-in statements (If, For, ...) are
# filtered out by the caller.
CALL_RE = re.compile(
    r'(?i)(?:(?:^|(?<=[\s:]))call\s+|^|(?<=[\s:]))'
    r'(?P<name>[A-Za-z_][A-Za-z0-9_]*)(?=\s*(?:\(|$|[\s:=,]))')

# VBA reserved words that look like calls but are control statements.
CONTROL_KEYWORDS = frozenset((
    'if', 'then', 'else', 'elseif', 'for', 'each', 'while', 'wend', 'do',
    'loop', 'until', 'select', 'case', 'with', 'end', 'exit', 'goto', 'gosub',
    'return', 'on', 'error', 'resume', 'next', 'dim', 'redim', 'const', 'set',
    'let', 'call', 'sub', 'function', 'property', 'declare', 'type', 'enum',
    'option', 'open', 'close', 'print', 'write', 'put', 'get', 'kill', 'rmdir',
    'mkdir', 'chdir', 'chdrive', 'sendkeys', 'beep', 'stop', 'rem', 'load',
    'unload', 'erase', 'randomize', 'line', 'circle', 'pset', 'scale', 'as',
    'and', 'or', 'not', 'mod', 'to', 'step', 'in', 'is', 'like', 'new',
    'byval', 'byref', 'optional', 'paramarray', 'preserve', 'alias', 'lib',
    'handles', 'implements', 'begin', 'attribute', 'vba_name', 'friend',
    'global', 'public', 'private', 'static', 'class', 'me',
))


def _collapse_continuations(vba_code):
    """
    Join physical lines split with the VBA line-continuation marker `` _``.

    :return: tuple ``(logical_lines, line_map)`` where ``line_map[i]`` is the
        1-based number of the *first* physical line that makes up logical line
        ``i``.  Continuation lines keep their original line numbers, which
        matters when locating evidence.
    """
    physical_lines = vba_code.splitlines()
    logical_lines = []
    line_map = []
    buffer_parts = []
    start_line = None
    for index, physical_line in enumerate(physical_lines, start=1):
        if start_line is None:
            start_line = index
        stripped = physical_line.rstrip()
        if stripped.endswith(' _'):
            # keep content before the marker, join split parts with a space
            buffer_parts.append(stripped[:-2].rstrip())
            continue
        buffer_parts.append(physical_line)
        logical_lines.append(' '.join(buffer_parts))
        line_map.append(start_line)
        buffer_parts = []
        start_line = None
    if buffer_parts:
        # file ended with a dangling continuation marker
        logical_lines.append(' '.join(buffer_parts))
        fallback = start_line if start_line is not None else len(physical_lines)
        line_map.append(fallback)
    return logical_lines, line_map


def _split_statements(logical_line):
    """
    Split a logical line into statements on ``:`` while keeping string
    literals and comments intact.

    A colon inside a quoted string or after a comment marker does not split.
    :return: list of statement strings (without the separating colons).
    """
    statements = []
    current = []
    in_string = False
    in_comment = False
    for char in logical_line:
        if in_comment:
            current.append(char)
            continue
        if char == '"':
            in_string = not in_string
            current.append(char)
        elif char == "'" and not in_string:
            in_comment = True
            current.append(char)
        elif char == ':' and not in_string:
            statements.append(''.join(current))
            current = []
        else:
            current.append(char)
    statements.append(''.join(current))
    return statements


def _strip_line_label(statement):
    """Remove a leading line label or number (e.g. ``10:`` / ``Retry:``)."""
    match = re.match(r'\s*[A-Za-z_][A-Za-z0-9_]*\s*:', statement)
    if match:
        return statement[match.end():]
    match = re.match(r'\s*\d+\s*:', statement)
    if match:
        return statement[match.end():]
    return statement


def lex_statement_segments(statement):
    """
    Split one VBA statement into classified segments.

    :return: list of ``(kind, text)`` tuples, with kind in
        :data:`RESOLUTION_CODE`, :data:`RESOLUTION_STRING` and
        :data:`RESOLUTION_COMMENT`.  Concatenating the texts rebuilds the
        original statement.
    """
    # "Rem" introduces a comment when it heads a statement (the whole rest of
    # the statement is then a comment, like an apostrophe comment).
    rem_match = re.match(r'\s*Rem(?=\s|$)', statement, re.IGNORECASE)
    if rem_match:
        return [(RESOLUTION_COMMENT, statement)]
    segments = []
    buffer_code = []
    buffer_text = []
    in_string = False

    def flush(kind, parts):
        text = ''.join(parts)
        if text:
            segments.append((kind, text))
        parts[:] = []

    for position, char in enumerate(statement):
        if in_string:
            buffer_text.append(char)
            if char == '"':
                in_string = False
                flush(RESOLUTION_STRING, buffer_text)
            continue
        if char == '"':
            flush(RESOLUTION_CODE, buffer_code)
            in_string = True
            buffer_text.append(char)
            continue
        if char == "'":
            flush(RESOLUTION_CODE, buffer_code)
            segments.append((RESOLUTION_COMMENT, statement[position:]))
            return segments
        buffer_code.append(char)
    if in_string:
        # unterminated string literal: treat what we have as a string
        flush(RESOLUTION_STRING, buffer_text)
    flush(RESOLUTION_CODE, buffer_code)
    return segments


def _keyword_in_code_pattern(keyword):
    """
    Case-insensitive regex for a plain (non-regex) keyword inside code.

    A standard word boundary is used: VBA is case-insensitive, so a variable
    named ``shell`` is the same identifier as the ``Shell`` statement and
    cannot be told apart by token shape alone.  Statement-position filtering
    is applied separately in :meth:`MacroEvidenceAnalyzer.keyword_evidence`.
    """
    return re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)


def _keyword_in_string_pattern(keyword):
    """
    Pattern matching a keyword anywhere inside a quoted string literal.

    Unlike code matching, no word boundary is required on the quote side, so a
    keyword embedded in a longer string (e.g. ``"run payload"``) is still
    classified as a string occurrence, not code.
    """
    return re.compile(r'"[^"]*' + re.escape(keyword) + r'[^"]*"',
                      re.IGNORECASE)


class MacroEvidenceAnalyzer(object):
    """
    Collect statement-level evidence for one or more VBA modules.

    Typical usage::

        analyzer = MacroEvidenceAnalyzer()
        analyzer.add_module('Module1', code_module_1)
        analyzer.add_module('Module2', code_module_2)
        evidence = analyzer.keyword_evidence(('Shell', 'Write'))
        calls = analyzer.procedure_calls()
    """

    def __init__(self):
        # module name -> list of Statement objects
        self.statements = {}
        # module name -> {procedure name: ProcedureInfo}
        self.procedures = {}

    #-- module ingestion ----------------------------------------------------

    def add_module(self, module_name, vba_code):
        """
        Parse one VBA module: collapse continuations, split statements and
        record every procedure declaration.

        :param module_name: str, name of the VBA module (used in every result)
        :param vba_code: str, raw VBA source code of the module
        """
        logical_lines, line_map = _collapse_continuations(vba_code)
        statements = []
        procedures = {}
        for logical_index, logical_line in enumerate(logical_lines):
            physical_line = line_map[logical_index]
            for statement_text in _split_statements(logical_line):
                statement_text = _strip_line_label(statement_text)
                if not statement_text.strip():
                    continue
                statement = Statement(module_name, physical_line,
                                      statement_text.strip())
                statements.append(statement)
                self._register_procedure(statement, procedures)
        self.statements[module_name] = statements
        self.procedures[module_name] = procedures

    @staticmethod
    def _register_procedure(statement, procedures):
        match = PROCEDURE_RE.match(statement.text)
        if not match:
            return
        name = match.group('name')
        kind = re.sub(r'\s+', ' ', match.group('kind').lower())
        scope = (match.group('scope') or '').lower() or 'default'
        procedures[name] = ProcedureInfo(
            module=statement.module, name=name, kind=kind, scope=scope,
            line=statement.line)

    #-- keyword evidence ----------------------------------------------------

    def keyword_evidence(self, keywords):
        """
        Locate every occurrence of the given plain keywords.

        Each occurrence is classified as code, string literal or comment, and
        located to its module and physical line.  Duplicate reports for the
        same ``(module, line, keyword, kind)`` are removed.

        For statement keywords (e.g. ``Shell``, ``Write``, ``Open``,
        ``Kill`` ...), an occurrence only counts as executable evidence when
        it sits at the head of a statement or after ``Call``: a variable or
        member named like the keyword (``Dim shell``, ``shell.Run``) is not an
        invocation.  Keywords that are not VBA statements (``Auto_Open``,
        ``CreateObject``, ``WScript.Shell`` ...) are matched anywhere in code.

        :param keywords: iterable of plain (non-regex) keyword strings
        :return: list of :class:`KeywordEvidence`, ordered by module then line
        """
        keywords = [keyword for keyword in keywords if keyword]
        code_patterns = {kw: _keyword_in_code_pattern(kw) for kw in keywords}
        string_patterns = {
            kw: _keyword_in_string_pattern(kw) for kw in keywords}
        seen = set()
        evidence = []
        for module_name in sorted(self.statements):
            for statement in self.statements[module_name]:
                segments = lex_statement_segments(statement.text)
                for kind, segment_text in segments:
                    for keyword in keywords:
                        if kind == RESOLUTION_STRING:
                            found = bool(string_patterns[keyword].search(
                                segment_text))
                        elif kind == RESOLUTION_COMMENT:
                            found = bool(code_patterns[keyword].search(
                                segment_text))
                        else:
                            found = self._keyword_in_code(
                                keyword, code_patterns[keyword],
                                segment_text)
                        if not found:
                            continue
                        identity = (module_name, statement.line,
                                    keyword.lower(), kind)
                        if identity in seen:
                            continue
                        seen.add(identity)
                        evidence.append(KeywordEvidence(
                            module=module_name, line=statement.line,
                            keyword=self._matched_keyword(
                                keyword, kind, segment_text,
                                code_patterns[keyword],
                                string_patterns[keyword]),
                            kind=kind,
                            statement=statement.text))
        return evidence

    @staticmethod
    def _matched_keyword(keyword, kind, segment_text, code_pattern,
                         string_pattern):
        """Return the keyword as it actually appears in the source."""
        pattern = string_pattern if kind == RESOLUTION_STRING else code_pattern
        match = pattern.search(segment_text)
        if kind == RESOLUTION_STRING and match is not None:
            inner = re.search(re.escape(keyword), match.group(),
                              re.IGNORECASE)
            if inner is not None:
                return inner.group()
        if match is not None:
            return match.group()
        return keyword

    @staticmethod
    def _keyword_in_code(keyword, pattern, code_segment):
        """
        Check whether ``keyword`` is used as executable code in a code segment.

        Statement keywords must appear at the segment head (or after
        ``Call``); other keywords are evidence anywhere (expressions,
        arguments, member access).
        """
        for match in pattern.finditer(code_segment):
            position = match.start()
            prefix = code_segment[:position]
            # nothing but whitespace before the keyword: statement head
            if not prefix.strip():
                return True
            # keyword follows an explicit "Call" statement
            if re.match(r'(?is)^.*\bCall\s+$', prefix):
                return True
            # non-statement keywords are evidence anywhere
            if keyword.lower() not in STATEMENT_KEYWORDS:
                return True
        return False

    #-- procedure map and calls ---------------------------------------------

    def all_procedures(self):
        """
        Flatten the procedure map across modules.

        :return: list of :class:`ProcedureInfo` ordered by module and line
        """
        result = []
        for module_name in sorted(self.procedures):
            result.extend(sorted(self.procedures[module_name].values(),
                                 key=lambda info: info.line))
        return result

    def locate_procedure(self, name):
        """
        Find every definition of a procedure name (same-named procedures may
        exist in different modules).

        :return: list of :class:`ProcedureInfo`, empty if the name is not
            defined anywhere in the analyzed modules
        """
        name_lower = name.lower()
        result = []
        for module_name in sorted(self.procedures):
            for procedure in self.procedures[module_name].values():
                if procedure.name.lower() == name_lower:
                    result.append(procedure)
        return result

    def procedure_calls(self):
        """
        Statically resolve procedure calls found in code segments.

        A bare call (``Foo`` / ``Call Foo``) to a name defined in exactly one
        module is resolved to that definition.  Names defined in several
        modules and runtime dispatch forms (``Application.Run``,
        ``CallByName``, ...) are marked :data:`RESOLUTION_UNKNOWN`.  Member
        calls such as ``obj.Method`` are left to the keyword layer and are not
        reported here.

        :return: list of :class:`CallEvidence` ordered by module then line
        """
        result = []
        seen = set()
        for module_name in sorted(self.statements):
            for statement in self.statements[module_name]:
                result.extend(self._calls_in_statement(statement, seen))
        return result

    def _calls_in_statement(self, statement, seen):
        segments = lex_statement_segments(statement.text)
        calls = []
        for kind, text in segments:
            if kind != RESOLUTION_CODE:
                continue
            calls.extend(self._scan_code_segment(statement, text, seen))
        return calls

    def _scan_code_segment(self, statement, code, seen):
        # Procedure declarations define calls rather than issuing them.
        if PROCEDURE_RE.match(statement.text):
            return []
        calls = []
        # Dynamic dispatch: the target name is computed at runtime.
        for dynamic_name in DYNAMIC_DISPATCH:
            for match in re.finditer(r'(?i)\b' + dynamic_name + r'\b', code):
                identity = (statement.module, statement.line,
                            dynamic_name.lower())
                if identity in seen:
                    continue
                seen.add(identity)
                calls.append(CallEvidence(
                    module=statement.module, line=statement.line,
                    callee=dynamic_name, resolution=RESOLUTION_UNKNOWN,
                    target_module=None, target_line=None,
                    reason='target name is resolved at runtime',
                    statement=statement.text))
        # Static, bare procedure calls.
        for match in CALL_RE.finditer(code):
            name = match.group('name')
            if name.lower() in CONTROL_KEYWORDS:
                continue
            identity = (statement.module, statement.line, name.lower())
            if identity in seen:
                continue
            definitions = self.locate_procedure(name)
            if not definitions:
                # unknown built-in / library call / variable: out of scope
                continue
            seen.add(identity)
            unique_modules = {info.module for info in definitions}
            if len(unique_modules) == 1 and len(definitions) == 1:
                target = definitions[0]
                calls.append(CallEvidence(
                    module=statement.module, line=statement.line,
                    callee=name, resolution=RESOLUTION_STATIC,
                    target_module=target.module, target_line=target.line,
                    reason=None, statement=statement.text))
            else:
                calls.append(CallEvidence(
                    module=statement.module, line=statement.line,
                    callee=name, resolution=RESOLUTION_UNKNOWN,
                    target_module=None, target_line=None,
                    reason='same-named procedure defined in %d modules'
                           % len(unique_modules),
                    statement=statement.text))
        return calls


class MacroEvidenceReport(object):
    """
    Immutable-style container aggregating the evidence produced for one VBA
    project (all modules).
    """

    __slots__ = ('keyword_evidence', 'procedures', 'calls')

    def __init__(self, keyword_evidence, procedures, calls):
        self.keyword_evidence = keyword_evidence
        self.procedures = procedures
        self.calls = calls

    def as_dict(self):
        """Serialize the report to plain dictionaries/lists (JSON-ready)."""
        return dict(
            keyword_evidence=[item.as_dict()
                              for item in self.keyword_evidence],
            procedures=[item.as_dict() for item in self.procedures],
            procedure_calls=[item.as_dict() for item in self.calls])

    def __repr__(self):
        return ('MacroEvidenceReport(keywords=%d, procedures=%d, calls=%d)'
                % (len(self.keyword_evidence), len(self.procedures),
                   len(self.calls)))


class Statement(object):
    """One logical VBA statement located in a module."""

    __slots__ = ('module', 'line', 'text')

    def __init__(self, module, line, text):
        self.module = module
        self.line = line
        self.text = text

    def __repr__(self):
        return 'Statement(%r, %d, %r)' % (self.module, self.line, self.text)


class ProcedureInfo(object):
    """A Sub/Function/Property declaration located in a module."""

    __slots__ = ('module', 'name', 'kind', 'scope', 'line')

    def __init__(self, module, name, kind, scope, line):
        self.module = module
        self.name = name
        self.kind = kind
        self.scope = scope
        self.line = line

    def as_dict(self):
        return dict(module=self.module, name=self.name, kind=self.kind,
                    scope=self.scope, line=self.line)

    def __repr__(self):
        return 'ProcedureInfo(%r, %r, line=%d)' % (
            self.module, self.name, self.line)


class KeywordEvidence(object):
    """A single keyword occurrence with its context classification."""

    __slots__ = ('module', 'line', 'keyword', 'kind', 'statement')

    def __init__(self, module, line, keyword, kind, statement):
        self.module = module
        self.line = line
        self.keyword = keyword
        self.kind = kind
        self.statement = statement

    def as_dict(self):
        return dict(module=self.module, line=self.line,
                    keyword=self.keyword, context=self.kind,
                    statement=self.statement)

    def __repr__(self):
        return 'KeywordEvidence(%r, %d, %r, %r)' % (
            self.module, self.line, self.keyword, self.kind)


class CallEvidence(object):
    """A procedure call and its static resolution status."""

    __slots__ = ('module', 'line', 'callee', 'resolution',
                 'target_module', 'target_line', 'reason', 'statement')

    def __init__(self, module, line, callee, resolution,
                 target_module, target_line, reason, statement):
        self.module = module
        self.line = line
        self.callee = callee
        self.resolution = resolution
        self.target_module = target_module
        self.target_line = target_line
        self.reason = reason
        self.statement = statement

    def as_dict(self):
        return dict(module=self.module, line=self.line, callee=self.callee,
                    resolution=self.resolution,
                    target_module=self.target_module,
                    target_line=self.target_line, reason=self.reason,
                    statement=self.statement)

    def __repr__(self):
        return 'CallEvidence(%r, %d, %r, %r)' % (
            self.module, self.line, self.callee, self.resolution)
