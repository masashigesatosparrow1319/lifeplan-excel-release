"""Non-Excel checks only. Windows APIs, registry and profile facts are mocked."""
import ast
import builtins
import errno
import importlib
import json
import os
from pathlib import Path
from types import SimpleNamespace
import sys
import subprocess
import xml.etree.ElementTree as ET

import pytest

from tools.excel2024_validation import windows_preflight as pre, no_excel_tests as safe_tests
from tools.excel2024_validation import preopen_input, run, session


@pytest.fixture
def setup(tmp_path, monkeypatch):
    profile = tmp_path / 'existing-profile.json'
    scratch = tmp_path / 'scratch'; scratch.mkdir()
    installation = dict(installation_path=r'c:\office', product_release_ids=['HomeBusiness2024Retail'],
                        version_to_report='16.0.1.2', platform='x64')
    value = dict(schema=pre.SCHEMA, licensed=True, ui_confirmed_on_owned_pid=True,
                 product='Microsoft Office Home and Business 2024',
                 facts=dict(pid=12345, path=r'C:\Office\EXCEL.EXE', version='16.0', exe_sha256='a' * 64,
                            file_version='16.0.1.2', build='1.2', installation=installation))
    profile.write_text(json.dumps(value), encoding='utf-8')
    monkeypatch.setattr(pre, 'windows_environment', lambda: dict(os='MOCK_WINDOWS_NOT_ACTUAL'))
    monkeypatch.setattr(pre, 'local_ntfs', lambda p: dict(filesystem='MOCK_NOT_ACTUAL'))
    monkeypatch.setattr(pre, 'read_installation', lambda p: installation)
    monkeypatch.setattr(pre, 'image_digest', lambda p: 'a' * 64)
    monkeypatch.setattr(pre, 'git_identity', lambda: dict(status='NOT_VERIFIABLE_ZIP', commit=None, branch=None))
    return profile, scratch, value, pre.source_fingerprint()['release_source_sha256']


def test_mock_preflight_never_writes_profile_or_ledger_and_only_dummy_io(setup):
    profile, scratch, _, expected = setup
    before = profile.read_bytes()
    result = pre.execute(profile, scratch, expected, True)
    assert result['status'] == 'PREFLIGHT_CHECKED' and result['next_action'] == 'STOP_PM_REVIEW'
    assert not result['excel_started'] and not result['com_connected'] and not result['approval_id_used']
    assert not result['gate2_acceptance'] and not result['product_acceptance']
    assert result['checks']['profile']['next_pid_must_equal_saved_pid'] is False
    assert result['checks']['profile']['current_license'] == 'NOT_CHECKED'
    assert profile.read_bytes() == before and not (profile.parent / pre.LEDGER).exists()
    assert sorted(p.name for p in scratch.iterdir()) == ['dummy-existing.bin', 'dummy-new.bin',
                                                       'dummy-published.bin', 'preflight.json']
    assert json.loads((scratch / 'preflight.json').read_text(encoding='utf-8')) == result
    assert str(scratch) not in json.dumps(result) and str(profile) not in json.dumps(result)


@pytest.mark.parametrize('fault', ['source', 'schema', 'license', 'ui', 'product', 'image', 'registry',
                                 'permission', 'nonempty', 'unconfirmed', 'environment', 'unexpected_private_error'])
def test_preflight_stops_before_dummy_io_and_redacts_private_errors(setup, monkeypatch, fault):
    profile, scratch, value, expected = setup
    if fault in ('schema', 'license', 'ui', 'product'):
        value[{'schema': 'schema', 'license': 'licensed', 'ui': 'ui_confirmed_on_owned_pid', 'product': 'product'}[fault]] = (
            'Microsoft 365' if fault == 'product' else False)
        profile.write_text(json.dumps(value), encoding='utf-8')
    if fault == 'image': monkeypatch.setattr(pre, 'image_digest', lambda p: 'b' * 64)
    if fault == 'registry': monkeypatch.setattr(pre, 'read_installation', lambda p: {})
    if fault == 'permission': monkeypatch.setattr(pre.os, 'access', lambda *a: False)
    if fault == 'nonempty': (scratch / 'KEEP').write_bytes(b'unchanged')
    if fault == 'environment':
        monkeypatch.setattr(pre, 'windows_environment', lambda: pre.require(False, 'WINDOWS_NOT_AVAILABLE'))
    if fault == 'unexpected_private_error':
        def fail(*a): raise OSError(r'C:\Users\PRIVATE_USER\SECRET_FILE')
        monkeypatch.setattr(pre, 'read_installation', fail)
    before = profile.read_bytes()
    result = pre.execute(profile, scratch, '0' * 64 if fault == 'source' else expected, fault != 'unconfirmed')
    assert result['status'] == 'BLOCKED'
    assert not list(scratch.glob('dummy-*')) and profile.read_bytes() == before
    assert 'PRIVATE_USER' not in json.dumps(result) and 'SECRET_FILE' not in json.dumps(result)
    if fault == 'nonempty': assert [p.name for p in scratch.iterdir()] == ['KEEP']


@pytest.mark.parametrize('pid', [12345, 99999])
def test_saved_pid_is_historical_not_a_requirement_for_next_process(setup, pid):
    profile, scratch, value, expected = setup
    value['facts']['pid'] = pid
    profile.write_text(json.dumps(value), encoding='utf-8')
    assert pre.execute(profile, scratch, expected, True)['status'] == 'PREFLIGHT_CHECKED'


@pytest.mark.parametrize('place', ['repo', 'ledger', 'onedrive', 'env_sync', 'traversal', 'unc'])
def test_unsafe_placements_rejected(tmp_path, monkeypatch, place):
    path = tmp_path
    if place == 'repo':
        with pytest.raises(pre.Blocked): pre.path_check(pre.REPO, outside=True)
        return
    if place in ('ledger', 'onedrive', 'env_sync'):
        path = tmp_path / {'ledger': pre.LEDGER.upper(), 'onedrive': 'OneDrive - Business', 'env_sync': 'custom-sync'}[place]
        path.mkdir()
    if place == 'env_sync': monkeypatch.setenv('OneDrive', str(path))
    if place == 'traversal': path = tmp_path / '..' / 'other'
    if place == 'unc': path = Path('//server/share')
    with pytest.raises(pre.Blocked): pre.path_check(path)


def test_reparse_guard_without_creating_privileged_links(tmp_path, monkeypatch):
    original = Path.lstat
    monkeypatch.setattr(Path, 'lstat', lambda p: SimpleNamespace(st_mode=0o100644, st_file_attributes=0x400)
                        if p == tmp_path else original(p))
    with pytest.raises(pre.Blocked, match='REPARSE'): pre.path_check(tmp_path)


def test_source_digest_agrees_with_existing_runtime_but_has_portable_pin():
    fingerprint = pre.source_fingerprint()
    assert fingerprint['source_sha256'] == run.source_identity()['source_sha256']
    assert all('\\' not in name and not Path(name).is_absolute() for name in fingerprint['files'])
    assert 'tests/test_excel2024_preopen_input.py' in fingerprint['test_and_dependency_files']
    assert 'requirements-excel2024.txt' in fingerprint['test_and_dependency_files']
    assert pre.check_source(fingerprint['release_source_sha256'])['original_sha256'] == pre.ORIGINAL_SHA
    assert {s.casefold() for s in pre.SUPPORTED_IDS} == set(session.EXCEL_2024_IDS)
    assert {s.casefold() for s in pre.COMPANION_IDS} == set(session.NON_EXCEL_IDS)


def test_dummy_collision_permission_error_is_blocked_not_success(tmp_path, monkeypatch):
    def denied(*a): raise PermissionError(errno.EACCES, 'PRIVATE_PATH')
    monkeypatch.setattr(preopen_input, '_move_new', denied)
    with pytest.raises(pre.Blocked, match='COLLISION_ACCESS'): pre.dummy_probe(tmp_path)
    assert (tmp_path / 'dummy-source.bin').read_bytes() == b'SOURCE'
    assert (tmp_path / 'dummy-existing.bin').read_bytes() == b'KEEP'


def test_dummy_overwrite_by_faulty_native_move_is_detected(tmp_path, monkeypatch):
    def broken(source, target): target.write_bytes(source.read_bytes())
    monkeypatch.setattr(preopen_input, '_move_new', broken)
    with pytest.raises(pre.Blocked, match='COLLISION_NOT_REJECTED'): pre.dummy_probe(tmp_path)


def test_existing_summary_never_overwritten(tmp_path):
    path = tmp_path / 'preflight.json'; path.write_bytes(b'KEEP')
    with pytest.raises(FileExistsError): pre.save_new(path, {})
    assert path.read_bytes() == b'KEEP'


def test_argument_errors_do_not_echo_private_values(capsys):
    parser = pre.SafeParser()
    with pytest.raises(SystemExit) as error: parser.parse_args(['--PRIVATE_USER_SECRET_PATH'])
    assert error.value.code == 2
    text = capsys.readouterr().err
    assert text == 'BLOCKED: ARGUMENTS_INVALID\n'


@pytest.mark.parametrize('fault', ['os', 'old_windows', 'python', 'dependencies'])
def test_platform_and_metadata_gate_without_importing_com(monkeypatch, fault):
    monkeypatch.setattr(pre, 'sys', SimpleNamespace(platform='linux' if fault == 'os' else 'win32',
        version_info=(3, 11, 0) if fault == 'python' else (3, 12, 10), maxsize=2**63,
        getwindowsversion=lambda: SimpleNamespace(major=10, build=19000 if fault == 'old_windows' else 26100, product_type=1)))
    monkeypatch.setattr(pre, 'version', lambda name: 'wrong' if fault == 'dependencies' else pre.DEPENDENCIES[name])
    with pytest.raises(pre.Blocked): pre.windows_environment()


@pytest.mark.parametrize('fault', ['none', 'network', 'filesystem', 'api', 'unc'])
def test_volume_api_is_readonly_metadata_with_no_com(monkeypatch, fault):
    calls = []
    class Function:
        def __init__(self, function): self.function = function
        def __call__(self, *args): calls.append(args); return self.function(*args)
    def volume(*args):
        args[6].value = 'FAT32' if fault == 'filesystem' else 'NTFS'
        return 0 if fault == 'api' else 1
    kernel = SimpleNamespace(GetDriveTypeW=Function(lambda *a: 4 if fault == 'network' else 3),
                             GetVolumeInformationW=Function(volume))
    monkeypatch.setattr(pre.ctypes, 'WinDLL', lambda *a, **kw: kernel, raising=False)
    path = Path(r'\\server\share\folder' if fault == 'unc' else 'C:/scratch')
    if fault == 'none': assert pre.local_ntfs(path) == dict(drive_type='FIXED_LOCAL', filesystem='NTFS')
    else:
        with pytest.raises(pre.Blocked): pre.local_ntfs(path)
    assert all(a[0] == 'C:\\' for a in calls)


@pytest.mark.parametrize('fault', ['none', 'missing_view', 'subscription', 'unknown', 'conflict', 'missing_value', 'access'])
def test_registry_reader_does_not_write_or_expose_private_paths(monkeypatch, fault):
    ids = ['O365ProPlusRetail'] if fault == 'subscription' else ['UNKNOWN'] if fault == 'unknown' else ['HomeBusiness2024Retail']
    def open_key(hive, name, reserved, flags):
        assert hive == 1 and reserved == 0
        assert flags in (0x20019 | 0x100, 0x20019 | 0x200)  # KEY_READ plus one WOW64 view
        assert name == r'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        if fault == 'missing_view' and flags & 0x100: raise FileNotFoundError('PRIVATE')
        if fault == 'access': raise PermissionError('PRIVATE')
        class Key:
            view_flags = flags
            def __enter__(self): return flags
            def __exit__(self, *a): pass
        return Key()
    def query(key, name):
        if fault == 'missing_value' and name == 'Platform': raise FileNotFoundError('PRIVATE')
        values = dict(InstallationPath=r'C:\Office', ProductReleaseIds=ids,
                      VersionToReport='16.0.2' if fault == 'conflict' and key.view_flags & 0x100 else '16.0.1', Platform='x64')
        return values[name], 1
    fake = SimpleNamespace(HKEY_LOCAL_MACHINE=1, KEY_READ=0x20019, KEY_WOW64_64KEY=0x100, KEY_WOW64_32KEY=0x200,
                           OpenKey=open_key, QueryValueEx=query)
    monkeypatch.setitem(sys.modules, 'winreg', fake)
    if fault in ('none', 'missing_view'):
        record = pre.read_installation(r'C:\Office\EXCEL.EXE')
        assert record['product_release_ids'] == ['HomeBusiness2024Retail']
    else:
        with pytest.raises((pre.Blocked, OSError)): pre.read_installation(r'C:\Office\EXCEL.EXE')


def test_image_hash_is_a_bounded_read_not_an_execution(tmp_path):
    path = tmp_path / 'dummy-image.bin'; path.write_bytes(b'NOT AN EXECUTABLE')
    assert pre.image_digest(path) == pre.digest(path.read_bytes())


def test_independent_modules_import_without_com_session_or_execute(monkeypatch):
    original = builtins.__import__
    def guarded(name, *args, **kwargs):
        assert name.split('.')[0] not in {'pythoncom', 'win32com', 'comtypes', 'psutil', 'win32api'}
        assert name not in {'session', 'worker', 'run'}
        return original(name, *args, **kwargs)
    with monkeypatch.context() as context:
        context.setattr(builtins, '__import__', guarded)
        importlib.reload(pre)
        importlib.reload(safe_tests)
    for module in (pre, safe_tests):
        tree = ast.parse(Path(module.__file__).read_text(encoding='utf-8'))
        calls = [n for n in ast.walk(tree) if isinstance(n, ast.Call)]
        assert not any(isinstance(n.func, ast.Attribute) and n.func.attr in ('prepare', 'connect', 'Dispatch', 'DispatchEx') for n in calls)


@pytest.mark.parametrize('event,args', [
    ('import', ('pythoncom',)), ('import', ('win32com.client',)), ('import', ('comtypes',)),
    ('ctypes.dlopen', ('ole32.dll',)), ('os.startfile', ('PRIVATE_PATH',)),
    ('os.system', ('PRIVATE_COMMAND',)), ('subprocess.Popen', ('excel.exe', ['excel.exe'])),
    ('subprocess.Popen', ('git', ['git', 'push'])),
])
def test_test_runner_guard_blocks_actual_activation_even_if_caught(event, args):
    guard = safe_tests.Guard()
    with pytest.raises(RuntimeError): guard(event, args)
    assert len(guard.violations) == 1 and 'PRIVATE' not in json.dumps(guard.violations)


def test_guard_allows_only_readonly_git_and_kernel_metadata():
    guard = safe_tests.Guard()
    guard('subprocess.Popen', ('git', ['git', 'rev-parse', 'HEAD'], pre.REPO, None))
    guard('ctypes.dlopen', ('kernel32',))
    guard('ctypes.dlopen', (None,))
    assert not guard.violations


def test_uname_query_is_linux_metadata_only_never_a_windows_exception():
    guard = safe_tests.Guard()
    guard.linux, guard.windows = True, False
    guard('subprocess.Popen', ('uname', ['uname', '-p'], None, None))
    guard.linux = False
    with pytest.raises(RuntimeError): guard('subprocess.Popen', ('uname', ['uname', '-p'], None, None))


def audit_guard(windows):
    guard = safe_tests.Guard()
    guard.windows, guard.linux = windows, not windows
    return guard


@pytest.mark.parametrize('command', [('git', 'rev-parse', 'HEAD'), ('git', 'rev-parse', '--abbrev-ref', 'HEAD')])
@pytest.mark.parametrize('shape', ['list', 'tuple', 'windows_default', 'windows_explicit'])
def test_audit_accepts_only_two_readonly_git_commands_in_proven_shapes(command, shape):
    windows = shape.startswith('windows')
    guard = audit_guard(windows)
    vector = list(command) if shape == 'list' else command
    payload = subprocess.list2cmdline(command) if windows else vector
    executable = None if shape == 'windows_default' else 'git'
    cwd = str(pre.REPO) if windows else pre.REPO
    guard('subprocess.Popen', (executable, payload, cwd, None))
    assert guard.violations == []


@pytest.mark.parametrize('payload', [
    'git push', 'git --help', 'git -c x=y rev-parse HEAD', 'git rev-parse HEAD extra',
    'git rev-parse HEAD && git push', 'git rev-parse HEAD & x', 'git rev-parse HEAD | x',
    'git rev-parse HEAD > x', 'git rev-parse HEAD; x', 'git rev-parse HEAD\n',
    ' git rev-parse HEAD', 'git rev-parse HEAD ', 'git  rev-parse HEAD', 'git\trev-parse HEAD',
    '"git" rev-parse HEAD', 'git rev-parse "HEAD"', 'git rev-parse HE^AD', 'git rev-parse %HEAD%',
    'git rev-parse HEAD\x00', 'cmd /c git rev-parse HEAD', 'powershell git rev-parse HEAD',
    'python -c git', 'GIT rev-parse HEAD', 'git.exe rev-parse HEAD',
])
def test_windows_command_strings_are_not_parsed_or_loosened(payload):
    guard = audit_guard(True)
    with pytest.raises(RuntimeError, match='PROCESS_LAUNCH'):
        guard('subprocess.Popen', (None, payload, str(pre.REPO), None))
    assert guard.violations == ['PROCESS_LAUNCH']


@pytest.mark.parametrize('executable', ['excel.exe', 'cmd.exe', 'powershell.exe', 'python.exe', 'git.exe',
                                     r'C:\Windows\System32\cmd.exe', r'C:\Git\git.exe', b'git', 0])
def test_audit_rejects_executable_substitution_even_with_allowed_command(executable):
    guard = audit_guard(True)
    with pytest.raises(RuntimeError):
        guard('subprocess.Popen', (executable, 'git rev-parse HEAD', str(pre.REPO), None))


@pytest.mark.parametrize('fault', ['cwd_missing', 'cwd_relative', 'cwd_other', 'cwd_bytes', 'env_empty', 'env_changed'])
def test_audit_requires_repo_cwd_and_inherited_environment_only(tmp_path, fault):
    cwd, env = pre.REPO, None
    if fault == 'cwd_missing': cwd = None
    if fault == 'cwd_relative': cwd = '.'
    if fault == 'cwd_other': cwd = tmp_path
    if fault == 'cwd_bytes': cwd = bytes(str(pre.REPO), 'utf-8')
    if fault == 'env_empty': env = {}
    if fault == 'env_changed': env = {'PATH': 'PRIVATE'}
    guard = audit_guard(False)
    with pytest.raises(RuntimeError): guard('subprocess.Popen', ('git', ['git', 'rev-parse', 'HEAD'], cwd, env))
    assert 'PRIVATE' not in json.dumps(guard.violations)


@pytest.mark.parametrize('payload', [None, (), ('git',), ('git', ['git', 'rev-parse', 'HEAD']),
                                  ('git', ['git', 'rev-parse', 'HEAD'], None),
                                  ['git', ['git', 'rev-parse', 'HEAD'], None, None],
                                  ('git', ['git', 'rev-parse', 'HEAD'], None, None, 'extra')])
def test_unrecognized_popen_structures_are_recorded_and_blocked(payload):
    guard = audit_guard(False)
    with pytest.raises(RuntimeError, match='PROCESS_LAUNCH'): guard('subprocess.Popen', payload)
    assert guard.violations == ['PROCESS_LAUNCH']


def test_string_forms_are_windows_only_and_null_executable_is_not_a_vector_fallback():
    guard = audit_guard(False)
    with pytest.raises(RuntimeError): guard('subprocess.Popen', (None, 'git rev-parse HEAD', str(pre.REPO), None))
    with pytest.raises(RuntimeError): guard('subprocess.Popen', (None, ['git', 'rev-parse', 'HEAD'], pre.REPO, None))


def test_custom_equality_objects_cannot_impersonate_commands():
    class Liar(str):
        def __eq__(self, other): return True
    guard = audit_guard(False)
    with pytest.raises(RuntimeError): guard('subprocess.Popen', ('git', [Liar('anything')], pre.REPO, None))
    with pytest.raises(RuntimeError): guard('subprocess.Popen', (Liar('anything'), ['git', 'rev-parse', 'HEAD'], pre.REPO, None))


@pytest.mark.parametrize('event,payload', [('import', ()), ('import', (b'pythoncom',)),
                                         ('ctypes.dlopen', ()), ('ctypes.dlopen', (b'ole32',))])
def test_malformed_security_events_fail_closed(event, payload):
    guard = safe_tests.Guard()
    with pytest.raises(RuntimeError): guard(event, payload)
    assert len(guard.violations) == 1


def test_violation_cap_and_forbidden_activations_are_unchanged():
    guard = audit_guard(True)
    for _ in range(25):
        with pytest.raises(RuntimeError): guard('subprocess.Popen', (None, 'git push', str(pre.REPO), None))
    assert guard.violations == ['PROCESS_LAUNCH'] * 20


@pytest.mark.parametrize('missing_git', [False, True])
def test_source_identity_handles_synthetic_windows_audit_and_absent_git(monkeypatch, missing_git):
    guard = audit_guard(True)
    seen = []
    def mock_git(command, **kwargs):
        seen.append(tuple(command))
        guard('subprocess.Popen', (None, subprocess.list2cmdline(command), str(kwargs['cwd']), kwargs.get('env')))
        if missing_git: raise FileNotFoundError('mock Git absent')
        return SimpleNamespace(returncode=0, stdout='a' * 40)
    monkeypatch.setattr(run.subprocess, 'run', mock_git)
    value = run.source_identity()
    assert seen == [('git', 'rev-parse', 'HEAD')] and not guard.violations
    assert value['commit'] == (None if missing_git else 'a' * 40)
    assert value['source_sha256'] == pre.source_fingerprint()['source_sha256']


@pytest.mark.parametrize('scenario', ['A', 'F24', 'E1', 'doctor_mock', 'F24_residual_failure'])
def test_selected_monitor_mocks_under_synthetic_windows_git_events(tmp_path, monkeypatch, scenario):
    from tests import test_excel2024_preopen_diagnostic as a
    from tests import test_excel2024_functions24 as f
    from tests import test_excel2024_e1 as e
    from tests import test_excel2024_doctor_aggregation as d
    guard = audit_guard(True)
    calls = []
    # This test models Windows Git events, not Linux platform's uname query.
    monkeypatch.setattr(run.platform, 'platform', lambda: 'SYNTHETIC_WINDOWS_AUDIT_TEST')
    def mock_git(command, **kwargs):
        guard('subprocess.Popen', (None, subprocess.list2cmdline(command), str(kwargs['cwd']), kwargs.get('env')))
        calls.append(True)
        raise FileNotFoundError('mock Git absent; no process started')
    monkeypatch.setattr(run.subprocess, 'run', mock_git)
    if scenario == 'A': result, _ = a.execute(tmp_path, monkeypatch)
    elif scenario == 'F24': result, _ = f.execute(tmp_path, monkeypatch)
    elif scenario == 'E1': result, _ = e.execute(tmp_path, monkeypatch)
    elif scenario == 'doctor_mock': result, _ = d.execute_doctor(tmp_path)
    else: result, _ = f.execute(tmp_path, monkeypatch, residual=lambda r: True)
    wanted = 'CONFIGURED' if scenario == 'doctor_mock' else 'FAIL' if scenario.endswith('failure') else 'DIAGNOSTIC_PASS'
    assert result['status'] == wanted and calls and not guard.violations
    assert result['excel_recalculation_verified'] is False


def test_redacted_junit_retains_outcomes_without_host_path_or_messages():
    collector = safe_tests.Results()
    for index, when, outcome in [(1, 'call', 'passed'), (2, 'call', 'failed'), (3, 'setup', 'failed'), (4, 'setup', 'skipped')]:
        collector.pytest_runtest_logreport(SimpleNamespace(nodeid=f'tests/test_x.py::test_x[PRIVATE_PATH{index}]',
                                                          when=when, outcome=outcome, duration=0.1))
    raw, counts = collector.xml()
    assert counts == dict(passed=1, failed=1, errors=1, skipped=1)
    assert b'PRIVATE' not in raw and b'hostname' not in raw
    assert ET.fromstring(raw).find('testsuite').get('tests') == '4'


@pytest.mark.parametrize('fault', ['none', 'test_fail', 'guard', 'zero_tests', 'nonempty'])
def test_safe_test_entry_has_no_diagnostic_calls_and_private_output_is_discarded(tmp_path, monkeypatch, fault):
    for key in list(os.environ):
        if key.startswith(('PYTEST_', 'P2A_', 'P2B_')) or 'RECALCULATED' in key:
            monkeypatch.delenv(key, raising=False)
    monkeypatch.setenv('PYTEST_DISABLE_PLUGIN_AUTOLOAD', '1')
    scratch = tmp_path / 'tests'; scratch.mkdir()
    if fault == 'nonempty': (scratch / 'KEEP').write_bytes(b'KEEP')
    guards = []
    def pytest_main(command, plugins):
        assert '--noconftest' in command and any(p.startswith('--basetemp=') for p in command)
        print(r'C:\Users\PRIVATE_USER\EXCEL.EXE')
        if fault == 'guard':
            with pytest.raises(RuntimeError): guards[0]('import', ('pythoncom',))
        if fault != 'zero_tests':
            plugins[0].pytest_runtest_logreport(SimpleNamespace(nodeid='tests/test_safe.py::test_safe', when='call',
                outcome='failed' if fault == 'test_fail' else 'passed', duration=0.1))
        return 1 if fault == 'test_fail' else 0
    result = safe_tests.execute(scratch, pre.source_fingerprint()['release_source_sha256'], True,
                                add_guard=guards.append, pytest_main=pytest_main)
    assert result['status'] == ('NO_EXCEL_TESTS_COMPLETE' if fault == 'none' else 'BLOCKED')
    assert 'PRIVATE_USER' not in json.dumps(result)
    assert not result['excel_started'] and not result['approval_id_used']
    if fault == 'nonempty': assert [p.name for p in scratch.iterdir()] == ['KEEP']
    else:
        assert json.loads((scratch / 'tests-summary.json').read_text(encoding='utf-8')) == result
        assert b'PRIVATE_USER' not in (scratch / 'tests.xml').read_bytes()
