"""Fixed, isolated pytest set with a process/import guard and redacted JUnit."""
from contextlib import redirect_stdout, redirect_stderr
import hashlib
import io
import json
import ntpath
import os
from pathlib import Path
import re
import sys
import time
import xml.etree.ElementTree as ET

from . import windows_preflight as pre

TESTS = (
    'tests/test_excel2024_windows_preflight.py',
    'tests/test_excel2024_preopen_input.py', 'tests/test_excel2024_preopen_diagnostic.py',
    'tests/test_excel2024_functions24.py', 'tests/test_excel2024_e1.py',
    'tests/test_excel2024_validation.py', 'tests/test_excel2024_calculation_wait.py',
    'tests/test_excel2024_doctor_aggregation.py', 'tests/test_excel2024_cleanup.py',
    'tests/test_p1.py', 'tests/test_p2a.py',
)
FORBIDDEN_IMPORTS = {'pythoncom', 'pywintypes', 'win32com', 'comtypes', 'clr'}


class Guard:
    """Fail closed on actual imports/launches; collect only fixed reason codes."""
    def __init__(self):
        self.violations = []
        self.linux = sys.platform.startswith('linux')
        self.windows = sys.platform == 'win32'
        self.repo = os.fspath(pre.REPO)

    def allowed_popen(self, args):
        """Closed audit shapes, not a shell parser or a basename allowlist."""
        if type(args) is not tuple or len(args) != 4:
            return False
        executable, command, cwd, env = args
        if env is not None or type(executable) not in (str, type(None)):
            return False
        commands = (('git', 'rev-parse', 'HEAD'), ('git', 'rev-parse', '--abbrev-ref', 'HEAD'))
        if type(command) is str:
            # CPython Windows list2cmdline + lpApplicationName=None is normal.
            if not self.windows or executable not in (None, 'git'):
                return False
            git = command in ('git rev-parse HEAD', 'git rev-parse --abbrev-ref HEAD')
        elif type(command) in (list, tuple) and all(type(item) is str for item in command):
            git = executable == 'git' and tuple(command) in commands
            # Preserve only the existing Linux OS-description query.
            if (self.linux and not self.windows and executable == 'uname'
                    and tuple(command) == ('uname', '-p') and cwd is None):
                return True
        else:
            return False
        if not git or type(cwd) not in (str, type(pre.REPO)):
            return False
        actual = os.fspath(cwd)
        return ntpath.normcase(actual) == ntpath.normcase(self.repo) if self.windows else actual == self.repo

    def __call__(self, event, args):
        reason = None
        if event == 'import':
            if (type(args) is not tuple or not args or type(args[0]) is not str
                    or args[0].split('.')[0] in FORBIDDEN_IMPORTS):
                reason = 'COM_IMPORT'
        if event == 'subprocess.Popen':
            if not self.allowed_popen(args):
                reason = 'PROCESS_LAUNCH'
        if event in ('os.system', 'os.startfile', 'os.exec', 'os.posix_spawn', 'os.spawn'):
            reason = 'OS_ACTIVATION'
        if event == 'ctypes.dlopen':
            if type(args) is not tuple or len(args) != 1 or type(args[0]) not in (str, type(None)):
                reason = 'COM_LIBRARY'
            elif args[0] is not None:
                library = ntpath.basename(args[0]).casefold().removesuffix('.dll')
                if library in ('ole32', 'oleaut32', 'combase'):
                    reason = 'COM_LIBRARY'
        if reason:
            if len(self.violations) < 20: self.violations.append(reason)
            raise RuntimeError('NO_EXCEL_GUARD_' + reason)


class Sink(io.TextIOBase):
    # pytest tracebacks may contain private paths; do not persist raw stdout.
    def write(self, value): return len(value)
    def flush(self): pass


class Results:
    def __init__(self): self.records = {}

    def pytest_runtest_logreport(self, report):
        self.records.setdefault(report.nodeid, []).append((report.when, report.outcome, report.duration))

    def pytest_collectreport(self, report):
        if report.failed:
            self.records.setdefault(report.nodeid, []).append(('collect', 'failed', 0))

    def xml(self):
        counts = dict(passed=0, failed=0, errors=0, skipped=0)
        root = ET.Element('testsuites')
        suite = ET.SubElement(root, 'testsuite', name='no_excel_preflight_tests')
        for index, (node, phases) in enumerate(self.records.items(), 1):
            failed = [phase for phase, outcome, _ in phases if outcome == 'failed']
            status = ('errors' if any(p != 'call' for p in failed) else 'failed') if failed else (
                'skipped' if any(o == 'skipped' for _, o, _ in phases) else
                'passed' if any(p == 'call' and o == 'passed' for p, o, _ in phases) else 'errors')
            counts[status] += 1
            name = node.split('::')[-1].split('[')[0]
            name = name if re.fullmatch('[A-Za-z_][A-Za-z_0-9]*', name) else 'case'
            case = ET.SubElement(suite, 'testcase', name=f'{name}__{index:04}',
                                 id=hashlib.sha256(node.encode()).hexdigest()[:16],
                                 time=f'{sum(d for _, _, d in phases):.6f}')
            if status != 'passed':
                ET.SubElement(case, {'failed': 'failure', 'errors': 'error', 'skipped': 'skipped'}[status],
                              message='DETAILS_REDACTED_PM_REVIEW' if status != 'skipped' else 'NOT_VERIFIED')
        suite.attrib.update(tests=str(sum(counts.values())), failures=str(counts['failed']),
                            errors=str(counts['errors']), skipped=str(counts['skipped']))
        return ET.tostring(root, encoding='utf-8', xml_declaration=True), counts


def execute(scratch, expected, confirmed=False, *, add_guard=sys.addaudithook, pytest_main=None):
    value = dict(schema='G2_NO_EXCEL_TESTS_V1', status='BLOCKED', tests=list(TESTS),
                 excel_started=False, com_connected=False, approval_id_used=False,
                 gate2_acceptance=False, next_action='STOP_PM_REVIEW', tests_host=sys.platform)
    safe = None
    guard = None
    started = time.monotonic()
    try:
        pre.require(__debug__, 'ASSERTS_REQUIRED')
        pre.require(confirmed, 'OWNER_LOCAL_UNSYNCED_CONFIRMATION_REQUIRED')
        candidate = pre.path_check(scratch, outside=True)
        pre.require(candidate.is_dir() and not any(candidate.iterdir()) and len(str(candidate)) <= 180, 'EMPTY_SCRATCH_REQUIRED')
        if sys.platform == 'win32':
            value['environment'] = pre.windows_environment()
            value['filesystem'] = pre.local_ntfs(candidate)
        safe = candidate
        value['release_source_sha256'] = pre.check_source(expected)['release_source_sha256']
        pre.require(not FORBIDDEN_IMPORTS.intersection(sys.modules), 'COM_MODULE_ALREADY_LOADED')
        guard = Guard()
        add_guard(guard)
        for name in list(os.environ):
            if ('RECALCULATED' in name or name.startswith(('P2A_S1_', 'S1_RECALCULATED', 'P2B_MEXT_'))
                    or name in ('PYTEST_ADDOPTS', 'PYTEST_PLUGINS')):
                os.environ.pop(name)
        os.environ['PYTEST_DISABLE_PLUGIN_AUTOLOAD'] = '1'
        with (safe / 'pytest.ini').open('x', encoding='utf-8') as stream:
            stream.write('[pytest]\n')
        collector = Results()
        command = ['-q', '--capture=sys', '--noconftest', '-c', str(safe / 'pytest.ini'),
                   '--rootdir=' + str(pre.REPO), '--basetemp=' + str(safe / 'pytest-temp')]
        command += [str(pre.REPO / p) for p in TESTS]
        with redirect_stdout(Sink()), redirect_stderr(Sink()):
            if pytest_main is None:
                import pytest
                pytest_main = pytest.main
            code = int(pytest_main(command, plugins=[collector]))
        xml, counts = collector.xml()
        with (safe / 'tests.xml').open('xb') as stream:
            stream.write(xml); stream.flush(); os.fsync(stream.fileno())
        value.update(counts=counts, pytest_exit_code=code, guard_violations=guard.violations,
                     junit='tests.xml', junit_sha256=pre.digest(xml))
        pre.require(code == 0 and counts['passed'] > 0 and counts['failed'] == counts['errors'] == 0
                    and not guard.violations, 'TESTS_OR_GUARD_FAILED')
        value['status'] = 'NO_EXCEL_TESTS_COMPLETE'
    except BaseException as error:
        value.update(reason=str(error) if type(error) is pre.Blocked else 'TEST_CHECK_FAILED',
                     error_type=type(error).__name__)
    if guard is not None:
        value['guard_violations'] = guard.violations
    value['elapsed_seconds'] = round(time.monotonic() - started, 3)
    if safe is not None:
        try:
            pre.save_new(safe / 'tests-summary.json', value)
        except BaseException:
            value.update(status='BLOCKED', report_write='FAILED')
    return value


def main():
    parser = pre.SafeParser(description=__doc__)
    parser.add_argument('--scratch', type=Path, required=True)
    parser.add_argument('--expected-source-sha', required=True)
    parser.add_argument('--local-unsynced-confirmed', action='store_true')
    args = parser.parse_args()
    result = execute(args.scratch, args.expected_source_sha, args.local_unsynced_confirmed)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result['status'] == 'NO_EXCEL_TESTS_COMPLETE' else 2


if __name__ == '__main__':
    raise SystemExit(main())
