"""SYNTHETIC_TEST_ONLY: fresh tiny3 mocks; never starts Excel/COM/doctor."""
import ast
import json
import os
from pathlib import Path
import shutil
import threading
import time
from types import SimpleNamespace
import zipfile
import xml.etree.ElementTree as ET

import pytest

from tools.excel2024_validation import E1_DIAGNOSTIC, E1_PROTOCOL, SCHEMA
from tools.excel2024_validation import common, fresh_tiny3 as fresh
from tools.excel2024_validation import fresh_tiny3_worker as wrapper
from tools.excel2024_validation import functions24, run as legacy_run, worker as legacy_worker


PIN = 'a' * 64
SOURCE = 'b' * 64
RUN_ID = '11111111-1111-1111-1111-111111111111'
APPROVAL = 'G2-20261017-017-WIN-99'
M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def profile():
    installation = dict(installation_path='C:\\Office',
                        product_release_ids=['HomeBusiness2024Retail'],
                        version_to_report='16.0', platform='x64')
    facts = dict(pid=900, path='C:\\Office\\EXCEL.EXE', created=10.0,
                 version='16.0', build='17932', exe_sha256='e' * 64,
                 file_version='16.0.17932.20000', installation=installation)
    return dict(schema=SCHEMA, product='Microsoft Office Home and Business 2024',
                licensed=True, facts=facts, ui_confirmed_on_owned_pid=True)


def process(pid, ppid, created):
    return dict(pid=pid, ppid=ppid, created=created, path=os.fspath(Path(os.sys.executable)))


CONTROLLER = process(100, 50, 1.0)
LAUNCHER = process(101, 100, 2.0)
WORKER = dict(**LAUNCHER, sys_executable=LAUNCHER['path'])


def rewrite_saved(source, target, *, cache=None, formula=None, added_value=None):
    cache = cache or {'B1': 10, 'B2': 21, 'B3': 31}
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    sheet = ET.fromstring(parts['xl/worksheets/sheet1.xml'])
    for cell in sheet.findall('.//' + M + 'c'):
        address = cell.get('r')
        if address == 'A1':
            value = cell.find(M + 'v')
            value.text = '7'
        if address in cache:
            value = cell.find(M + 'v')
            if value is None:
                value = ET.SubElement(cell, M + 'v')
            value.text = str(cache[address])
        if formula and address == formula[0]:
            cell.find(M + 'f').text = formula[1]
    if added_value:
        data = sheet.find(M + 'sheetData')
        row = ET.SubElement(data, M + 'row', {'r': '5'})
        cell = ET.SubElement(row, M + 'c', {'r': added_value[0], 't': 'n'})
        ET.SubElement(cell, M + 'v').text = str(added_value[1])
        sheet.find(M + 'dimension').set('ref', 'A1:C5')
    parts['xl/worksheets/sheet1.xml'] = ET.tostring(sheet)
    workbook = ET.fromstring(parts['xl/workbook.xml'])
    calc = workbook.find(M + 'calcPr')
    calc.set('calcMode', 'auto')
    parts['xl/workbook.xml'] = ET.tostring(workbook)
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, raw in parts.items():
            archive.writestr(name, raw)


def events(directory, request, *, interactive=True, worker_pid=101):
    names = ['ownership.verified', 'identity.verified', 'settings.after_call',
             'final.full_rebuild.before_call', 'final.full_rebuild.after_call',
             'final.automatic_read.before_call', 'final.automatic_read.after_call',
             'calculation.wait_done', 'save_as.after_call', 'reopen.after_call',
             'calculation.wait_done', 'cleanup.quit.after_call', 'cleanup.release.after_call']
    records = []
    for number, name in enumerate(names, 1):
        records.append(dict(schema=SCHEMA, run_id=request['run_id'], case_id='tiny3',
                            sequence=number, utc='2026-10-17T00:00:00+00:00',
                            worker_pid=worker_pid, event=name))
    for prop, value in [('Ready', True), ('Interactive', interactive), ('DisplayAlerts', True)]:
        records.append(dict(schema=SCHEMA, run_id=request['run_id'], case_id='tiny3',
                            sequence=len(records) + 1, utc='2026-10-17T00:00:00+00:00',
                            worker_pid=worker_pid, event='cleanup.observed', property=prop, value=value))
    (directory / 'progress.jsonl').write_text(
        ''.join(json.dumps(record, ensure_ascii=False) + '\n' for record in records), encoding='utf-8')


def e1_evidence(directory, manifest, values):
    phases = ['opened', 'before_flag', 'after_flag', 'after_full_rebuild', 'after_automatic', 'reopened']
    result = []
    for phase in phases:
        reopened = phase == 'reopened'
        path = directory / ('results' if reopened else 'work') / manifest['file']
        mode = -4135 if phase in ('before_flag', 'after_flag', 'after_full_rebuild') else -4105
        force = phase in ('opened', 'before_flag')
        sample = dict(phase='e1.' + phase, state_before=0, state=0, mode=mode,
                      values=values, differences=[])
        result.append(dict(phase=phase, utc='2026-10-17T00:00:00+00:00', force_full=force,
                           workbooks=[dict(name=manifest['file'], path=str(path.resolve()), read_only=reopened)],
                           sample=sample))
    return result


def fake_monitor(directory, request, *, problem=None):
    fresh._write_new_json(directory / 'request.json', request)
    if problem == 'timeout':
        return dict(status='TIMEOUT', worker_still_running=True, cleanup_pending=True,
                    force_termination=False)
    manifest = common.read_json(directory / 'case.json')
    saved = directory / 'results' / manifest['file']
    cache = {'B1': 10, 'B2': 21, 'B3': 999 if problem == 'cache' else 31}
    rewrite_saved(directory / 'inputs' / manifest['file'], saved, cache=cache,
                  formula=('B3', 'SUM(B1:B1)') if problem == 'formula' else None,
                  added_value=('C5', 1) if problem == 'nonformula' else None)
    values = {'診断!B1': 10, '診断!B2': 21, '診断!B3': 31}
    evidence = e1_evidence(directory, manifest, values)
    if problem == 'documents':
        evidence[0]['workbooks'].append(dict(name='other.xlsx', path='C:\\other.xlsx', read_only=False))
    samples = [dict(phase=name, state_before=0, state=0, mode=-4105,
                    values=values, differences=[]) for name in ('before_save', 'reopened')]
    proof_worker = WORKER.copy()
    proof_launcher = LAUNCHER.copy()
    if problem == 'pid':
        proof_worker['pid'] = 102
    result = dict(schema=SCHEMA, run_id=request['run_id'], case_id='tiny3',
                  request_sha256=common.sha(directory / 'request.json'),
                  manifest_sha256=common.sha(directory / 'case.json'), status='SAVED_REOPENED',
                  first_error=None, cleanup_errors=['close failed'] if problem == 'close' else [],
                  progress_log_errors=[], identity=profile()['facts'].copy(), observed_identity=profile()['facts'].copy(),
                  full_rebuild_calls=1, samples=samples, saved_sha256=common.sha(saved),
                  e1=dict(diagnostic=E1_DIAGNOSTIC, protocol=E1_PROTOCOL, evidence=evidence),
                  worker_pid=101, worker_still_running=problem == 'running',
                  cleanup_pending=problem == 'running', force_termination=False,
                  python_processes=dict(controller=CONTROLLER, launcher=proof_launcher,
                                        worker=proof_worker, worker_exit=dict(outcome='PROCESS_EXITED')))
    if problem == 'identity':
        result['identity']['build'] = 'UNEXPECTED'
    events(directory, request, interactive=problem != 'modal')
    return result


@pytest.fixture
def packet(tmp_path):
    config = tmp_path / 'profile.json'
    config.write_text(json.dumps(profile()), encoding='utf-8')
    return SimpleNamespace(config=config, output=tmp_path / 'fresh-run')


def execute(packet, *, problem=None, confirmed=True, approval=APPROVAL, source=PIN, monitor=None):
    monitor = monitor or (lambda directory, request: fake_monitor(directory, request, problem=problem))
    def residual(runtime):
        runtime['residual_observation'] = dict(outcome='PROCESS_EXITED')
        return False
    return fresh.execute(packet.config, packet.output, source, approval, confirmed=confirmed,
                         environment=lambda: {'windows': 'SYNTHETIC'},
                         source_check=lambda expected: dict(release_source_sha256=PIN, source_sha256=SOURCE),
                         monitor=monitor, residual=residual,
                         process_identity=lambda pid: CONTROLLER.copy(), run_id_factory=lambda: RUN_ID,
                         filesystem_check=lambda path: dict(drive_type='FIXED_LOCAL', filesystem='NTFS'))


def test_fresh_tiny3_happy_path_is_single_case_diagnostic_only(packet):
    original = Path('output/s1_tiny_diagnostic/tiny_calc.xlsx').read_bytes()
    result = execute(packet)
    assert result['status'] == 'FRESH_TINY3_DIAGNOSTIC_PASS' and result['exit_code'] == 0
    assert result['success'] is True and result['diagnostic_recalculation_verified'] is True
    assert result['p1_status'] == result['annual1_status'] == 'NOT_RUN'
    assert result['gate2_acceptance'] is result['product_acceptance'] is False
    assert result['python_process_model'] == 'DIRECT_FRESH_WORKER'
    assert result['verification']['numeric_comparison'] == 'matched'
    assert result['verification']['all_nonformula_values'] == 'matched'
    assert result['progress']['cleanup_observed'] == {'Ready': True, 'Interactive': True, 'DisplayAlerts': True}
    assert set(result['filesystem']) == {'source_root', 'profile', 'output_parent'}
    assert all(value == {'drive_type': 'FIXED_LOCAL', 'filesystem': 'NTFS'}
               for value in result['filesystem'].values())
    assert common.read_json(packet.output / 'run.json')['status'] == result['status']
    assert Path('output/s1_tiny_diagnostic/tiny_calc.xlsx').read_bytes() == original
    assert sorted(path.name for path in (packet.output / 'tiny3').iterdir()) == [
        'case.json', 'inputs', 'progress.jsonl', 'request.json', 'results', 'work']


@pytest.mark.parametrize('problem,expected', [
    ('pid', 'BLOCKED'), ('identity', 'BLOCKED'), ('modal', 'BLOCKED'), ('documents', 'FAIL'), ('close', 'BLOCKED'),
    ('running', 'BLOCKED'), ('cache', 'FAIL'), ('formula', 'FAIL'), ('nonformula', 'BLOCKED')])
def test_runtime_process_modal_document_cleanup_and_saved_changes_fail_closed(packet, problem, expected):
    result = execute(packet, problem=problem)
    assert result['status'] == expected and result['success'] is False
    assert result['exit_code'] == (2 if expected == 'BLOCKED' else 1)
    assert result['gate2_acceptance'] is False


def test_timeout_is_not_success_and_never_forces_worker(packet):
    result = execute(packet, problem='timeout')
    assert result['status'] == 'TIMEOUT' and result['exit_code'] == 1
    assert result['force_termination'] is False and result['gate2_acceptance'] is False


@pytest.mark.parametrize('condition', ['owner', 'approval', 'source', 'profile', 'output'])
def test_preconditions_stop_before_monitor_or_approval_use(packet, condition):
    calls = []
    kwargs = {}
    if condition == 'owner': kwargs['confirmed'] = False
    elif condition == 'approval': kwargs['approval'] = 'G2-INVALID'
    elif condition == 'source': kwargs['source'] = 'bad'
    elif condition == 'profile': packet.config.write_text('{}', encoding='utf-8')
    else: packet.output.mkdir()
    result = execute(packet, monitor=lambda *args: calls.append(args), **kwargs)
    assert result['status'] == 'BLOCKED' and result['exit_code'] == 2 and not calls
    assert result['excel_started'] is result['com_connected'] is False


def test_source_mismatch_and_bad_fixture_stop_before_monitor(packet, monkeypatch):
    calls = []
    value = fresh.execute(packet.config, packet.output, PIN, APPROVAL, confirmed=True,
                          environment=lambda: None,
                          source_check=lambda expected: dict(release_source_sha256='c' * 64),
                          monitor=lambda *args: calls.append(args),
                          process_identity=lambda pid: CONTROLLER, run_id_factory=lambda: RUN_ID,
                          filesystem_check=lambda path: dict(drive_type='FIXED_LOCAL', filesystem='NTFS'))
    assert value['reason'] == 'SOURCE_SHA_MISMATCH' and not calls
    packet.output = packet.output.parent / 'second'
    monkeypatch.setattr(fresh.cases, 'prepare', lambda *args: dict(run_id=RUN_ID, case_id='annual1'))
    value = execute(packet, monitor=lambda *args: calls.append(args))
    assert value['status'] == 'BLOCKED' and value['reason'] == 'TINY3_MANIFEST_MISMATCH' and not calls


@pytest.mark.parametrize('blocked_check', [1, 2, 3], ids=['source-root', 'profile', 'output-parent'])
def test_nonfixed_or_non_ntfs_path_stops_before_approval_and_monitor(packet, blocked_check):
    filesystem_calls, monitor_calls = [], []
    def filesystem(path):
        filesystem_calls.append(Path(path))
        if len(filesystem_calls) == blocked_check:
            return dict(drive_type='REMOTE_OR_REMOVABLE', filesystem='NTFS')
        return dict(drive_type='FIXED_LOCAL', filesystem='NTFS')
    result = fresh.execute(
        packet.config, packet.output, PIN, APPROVAL, confirmed=True,
        environment=lambda: {'windows': 'SYNTHETIC'},
        source_check=lambda expected: dict(release_source_sha256=PIN, source_sha256=SOURCE),
        monitor=lambda *args: monitor_calls.append(args), process_identity=lambda pid: CONTROLLER,
        run_id_factory=lambda: RUN_ID, filesystem_check=filesystem)
    assert result['status'] == 'BLOCKED' and result['reason'] == 'LOCAL_FIXED_NTFS_UNVERIFIED'
    assert result['approval_id_used'] is False and not monitor_calls
    assert len(filesystem_calls) == blocked_check


def test_native_filesystem_rejection_is_redacted_and_stops_before_monitor(packet):
    monitor_calls = []
    def rejected(path):
        raise fresh.pre.Blocked('FIXED_LOCAL_DRIVE_REQUIRED')
    result = fresh.execute(
        packet.config, packet.output, PIN, APPROVAL, confirmed=True,
        environment=lambda: {'windows': 'SYNTHETIC'},
        source_check=lambda expected: dict(release_source_sha256=PIN, source_sha256=SOURCE),
        monitor=lambda *args: monitor_calls.append(args), process_identity=lambda pid: CONTROLLER,
        run_id_factory=lambda: RUN_ID, filesystem_check=rejected)
    assert result['status'] == 'BLOCKED' and result['reason'] == 'LOCAL_FIXED_NTFS_UNVERIFIED'
    assert result['approval_id_used'] is False and not monitor_calls


def test_approval_id_is_spent_once_without_retry(packet):
    assert execute(packet)['status'] == 'FRESH_TINY3_DIAGNOSTIC_PASS'
    packet.output = packet.output.parent / 'retry'
    calls = []
    result = execute(packet, monitor=lambda *args: calls.append(args))
    assert result['status'] == 'FAIL' and result['reason'] == 'FRESH_TINY3_VALIDATION_FAILED'
    assert not calls and result['approval_id_used'] is False


def test_validate_manifest_pins_original_and_fixed_expectations(tmp_path):
    directory = tmp_path / 'case'
    directory.mkdir()
    manifest = fresh.validate_manifest(fresh.cases.prepare('tiny3', directory, RUN_ID))
    assert manifest['sha256'] == fresh.ORIGINAL_SHA
    assert manifest['probes'] == [['診断', 'B1', 10], ['診断', 'B2', 21], ['診断', 'B3', 31]]
    assert manifest['edits'] == [['診断', 'A1', 7], ['診断', 'A2', 3]]


def wrapper_packet(tmp_path):
    directory = tmp_path / 'tiny3'; directory.mkdir()
    manifest = fresh.cases.prepare('tiny3', directory, RUN_ID)
    controller = process(200, 50, 1.0)
    profile_path = tmp_path / 'profile.json'
    profile_path.write_text('{}', encoding='utf-8')
    approval = dict(schema=fresh.FRESH_SCHEMA, approval_id=APPROVAL, run_id=RUN_ID,
                    source_sha256=SOURCE, original_sha256=fresh.ORIGINAL_SHA,
                    profile_sha256=common.sha(profile_path), created_utc='2026-10-17T00:00:00+00:00')
    ledger = tmp_path / fresh.FRESH_LEDGER; ledger.mkdir()
    marker = ledger / (APPROVAL + '.json')
    fresh._write_new_json(marker, approval)
    request = dict(run_id=RUN_ID, case_id='tiny3', diagnostic=E1_DIAGNOSTIC, protocol=E1_PROTOCOL,
                   profile={}, profile_sha256=approval['profile_sha256'],
                   manifest_sha256=common.sha(directory / 'case.json'), source_sha256=SOURCE,
                   fresh_tiny3_control=dict(**approval, controller=controller,
                                            profile_path=str(profile_path),
                                            approval_marker=str(marker), approval_marker_sha256=common.sha(marker)))
    fresh._write_new_json(directory / 'request.json', request)
    return directory, request, controller


def test_wrapper_requires_actual_parent_and_authorization_before_worker(tmp_path):
    directory, request, controller = wrapper_packet(tmp_path)
    worker = process(201, controller['pid'], 2.0)
    calls = []
    def identities(pid):
        return controller if pid == os.getppid() else worker
    def authorize(seconds):
        proof = directory / 'python-worker.json'
        fresh._write_new_json(directory / 'launch-authorization.json',
                              dict(schema=fresh.FRESH_SCHEMA,
                                   request_sha256=common.sha(directory / 'request.json'),
                                   worker_sha256=common.sha(proof)))
    code = wrapper.main(directory / 'request.json', process_identity=identities,
                        worker_main=lambda path: calls.append(path) or 0, pause=authorize)
    assert code == 0 and calls == [directory / 'request.json']
    assert common.read_json(directory / 'python-worker.json')['process']['pid'] == worker['pid']


def test_wrapper_parent_mismatch_or_stop_never_calls_worker(tmp_path):
    directory, request, controller = wrapper_packet(tmp_path)
    calls = []
    wrong = process(999, 50, 1.0)
    with pytest.raises(fresh.Blocked, match='ACTUAL_PARENT_MISMATCH'):
        wrapper.main(directory / 'request.json', process_identity=lambda pid: wrong,
                     worker_main=lambda path: calls.append(path))
    assert not calls and not (directory / 'python-worker.json').exists()
    worker = process(201, controller['pid'], 2.0)
    (directory / 'STOP').touch()
    with pytest.raises(fresh.Blocked, match='PARENT_STOP_BEFORE_COM'):
        wrapper.main(directory / 'request.json',
                     process_identity=lambda pid: controller if pid == os.getppid() else worker,
                     worker_main=lambda path: calls.append(path))
    assert not calls and not (directory / 'launch-authorization.json').exists()


def test_supervisor_rejects_worker_executable_mismatch_before_authorization(tmp_path, monkeypatch):
    directory = tmp_path / 'case'; directory.mkdir()
    controller = process(os.getpid(), 1, 1.0)
    launcher = process(321, controller['pid'], 2.0)
    request = dict(run_id=RUN_ID, case_id='tiny3', fresh_tiny3_control={'controller': controller})
    class Proc:
        pid = 321
        returncode = 2
        def poll(self): return self.returncode
    def popen(command, **kwargs):
        proof = dict(schema=fresh.FRESH_SCHEMA, request_sha256=common.sha(directory / 'request.json'),
                     process=dict(**launcher, sys_executable=os.fspath(Path('/different/python'))))
        fresh._write_new_json(directory / 'python-worker.json', proof)
        return Proc()
    monkeypatch.setattr(functions24, 'worker_exit_observation',
                        lambda process_value: {'outcome': 'PROCESS_EXITED'})
    result = fresh.supervise(
        directory, request, popen=popen,
        process_identity=lambda pid: controller if pid == os.getpid() else launcher)
    assert result['status'] == 'BLOCKED'
    assert result['reason'] == 'PRECOM_FAILURE'
    assert result['failure_stage'] == 'PARENT_EXECUTABLE_VERIFY'
    assert result['reason_code'] == 'EXECUTABLE_MISMATCH'
    assert result['first_error'] == {'type': 'PreComFailure', 'message': 'EXECUTABLE_MISMATCH'}
    assert not (directory / 'launch-authorization.json').exists()
    assert (directory / 'STOP').exists() and result['force_termination'] is False


def test_full_handshake_invokes_real_legacy_worker_interface_and_records_exit(tmp_path, monkeypatch):
    directory, request, controller = wrapper_packet(tmp_path)
    (directory / 'request.json').unlink()
    launcher = process(321, controller['pid'], 2.0)
    worker_process = dict(**launcher, sys_executable=launcher['path'])
    calls = []

    class FakeSession:
        def __init__(self, log, profile_value):
            self.log, self.facts = log, profile_value.get('facts', profile()['facts'])
            self.owned, self.owner, self.observed_facts = False, None, None
        def connect(self): calls.append('connect')
        def close(self): calls.append('close')
        def release(self): calls.append('release')

    def fake_execute(session, manifest, case_directory, log, books, **kwargs):
        calls.append(('execute_case', manifest['case_id'], kwargs['diagnostic']))
        return dict(status='SAVED_REOPENED', samples=[], full_rebuild_calls=1,
                    saved_sha256='f' * 64)

    class FakeProgress(common.Progress):
        def emit(self, event, **details):
            self.sequence += 1
            record = dict(schema=SCHEMA, run_id=self.run_id, case_id=self.case_id,
                          sequence=self.sequence, utc=common.utc(), worker_pid=launcher['pid'],
                          event=event, **details)
            with self.path.open('a', encoding='utf-8') as stream:
                stream.write(json.dumps(record, ensure_ascii=False, default=str) + '\n')

    monkeypatch.setattr(legacy_worker, 'Session', FakeSession)
    monkeypatch.setattr(legacy_worker, 'execute_case', fake_execute)
    monkeypatch.setattr(legacy_worker, 'Progress', FakeProgress)
    monkeypatch.setattr(legacy_run, 'source_identity', lambda: {'source_sha256': SOURCE})
    monkeypatch.setattr(functions24, 'worker_exit_observation',
                        lambda process_value: {'outcome': 'PROCESS_EXITED'})

    class Proc:
        pid = launcher['pid']
        returncode = None
        def __init__(self, request_path):
            def invoke():
                self.returncode = wrapper.main(
                    request_path,
                    process_identity=lambda pid: controller if pid == os.getppid() else launcher,
                    worker_main=legacy_worker.main, pause=time.sleep)
            self.thread = threading.Thread(target=invoke, daemon=True)
            self.thread.start()
        def poll(self):
            return None if self.thread.is_alive() else self.returncode

    result = fresh.supervise(
        directory, request, popen=lambda command, **kwargs: Proc(Path(command[-1])),
        process_identity=lambda pid: controller if pid == os.getpid() else launcher,
        pause=lambda seconds: time.sleep(.001), overall_timeout=2)
    assert result['status'] == 'SAVED_REOPENED', json.dumps(result, ensure_ascii=False, default=str)
    assert result['python_processes']['worker'] == worker_process
    assert result['python_processes']['worker_exit']['outcome'] == 'PROCESS_EXITED'
    assert calls == ['connect', ('execute_case', 'tiny3', E1_DIAGNOSTIC), 'close', 'release']
    assert (directory / 'launch-authorization.json').is_file()
    assert not (directory / 'STOP').exists()


def test_supervisor_timeout_writes_stop_without_kill_or_terminate(tmp_path):
    directory = tmp_path / 'case'; directory.mkdir()
    controller = process(os.getpid(), 1, 1.0)
    launcher = process(321, controller['pid'], 2.0)
    request = dict(run_id=RUN_ID, case_id='tiny3', fresh_tiny3_control={'controller': controller})
    now = [0.0]
    class Proc:
        pid = 321
        returncode = None
        def poll(self): return None
    def popen(command, **kwargs):
        proof = dict(schema=fresh.FRESH_SCHEMA, request_sha256=common.sha(directory / 'request.json'),
                     process=dict(**launcher, sys_executable=launcher['path']))
        fresh._write_new_json(directory / 'python-worker.json', proof)
        return Proc()
    def clock(): return now[0]
    def pause(seconds): now[0] += seconds
    result = fresh.supervise(directory, request, popen=popen,
                             process_identity=lambda pid: controller if pid == os.getpid() else launcher,
                             clock=clock, pause=pause, overall_timeout=.2)
    assert result['status'] == 'TIMEOUT' and result['worker_still_running'] is True
    assert result['force_termination'] is False and (directory / 'STOP').exists()
    tree = ast.parse(Path(fresh.__file__).read_text(encoding='utf-8'))
    calls = {ast.unparse(node.func) for node in ast.walk(tree) if isinstance(node, ast.Call)}
    assert not any(name.endswith(('.kill', '.terminate')) for name in calls)


def test_cli_redacts_private_arguments_and_keeps_gate2_false(monkeypatch, capsys):
    private = 'C:\\PRIVATE\\profile.json'
    assert fresh.main(['--config', private]) == 2
    captured = capsys.readouterr()
    value = json.loads(captured.out)
    assert value['status'] == 'BLOCKED' and value['gate2_acceptance'] is False
    assert private not in captured.out and captured.err == ''


def test_modules_do_not_import_com_or_reference_other_cases():
    source = Path(fresh.__file__).read_text(encoding='utf-8')
    worker_source = Path(wrapper.__file__).read_text(encoding='utf-8')
    combined = source + worker_source
    assert all(token not in combined for token in ('win32com', 'pythoncom', "prepare('annual1'", "prepare('functions24'"))
    assert 'subprocess.Popen' in source and 'fresh_tiny3_worker' in source
    assert 'Gate 2 acceptance' in source and 'STOP_PM_REVIEW' in source


def read_failure(directory):
    value = common.read_json(directory / 'precom-failure.json')
    assert set(value) == fresh.PRECOM_KEYS
    assert value['schema'] == fresh.PRECOM_SCHEMA
    assert value['stage'] in fresh.PRECOM_STAGES
    assert value['reason_code'] in fresh.PRECOM_REASONS
    assert value['is_precom'] is True and value['com_start_attempted'] is False
    assert value['success'] is False
    return value


def test_precom_schema_is_finite_redacted_and_first_failure_is_immutable(tmp_path):
    private = 'SECRET_USER_PATH_TOKEN_1234'
    value = fresh.write_precom_failure(tmp_path, private, private)
    assert value == fresh.precom_failure('UNKNOWN', 'PRECOM_UNKNOWN')
    before = (tmp_path / 'precom-failure.json').read_bytes()
    fresh.write_precom_failure(tmp_path, 'PARENT_PROCESS_START', 'PROCESS_START_FAILED')
    assert (tmp_path / 'precom-failure.json').read_bytes() == before
    assert private not in before.decode('utf-8') and len(before) < 1024
    assert read_failure(tmp_path) == value


@pytest.mark.parametrize('case,stage,reason', [
    ('request_file', 'WORKER_REQUEST_PARSE', 'REQUEST_FILE_INVALID'),
    ('request_json', 'WORKER_REQUEST_PARSE', 'REQUEST_JSON_INVALID'),
    ('control', 'WORKER_CONTROL_VALIDATE', 'CONTROL_BINDING_INVALID'),
    ('manifest', 'WORKER_MANIFEST_VALIDATE', 'MANIFEST_INVALID'),
    ('marker', 'WORKER_MARKER_VALIDATE', 'APPROVAL_MARKER_INVALID'),
    ('profile', 'WORKER_PROFILE_VALIDATE', 'PROFILE_BINDING_INVALID'),
    ('parent', 'WORKER_PROCESS_IDENTITY', 'PARENT_IDENTITY_MISMATCH'),
    ('identity_error', 'WORKER_PROCESS_IDENTITY', 'PROCESS_IDENTITY_UNAVAILABLE'),
    ('proof_collision', 'WORKER_PROOF_CREATE', 'PROOF_CREATE_FAILED'),
])
def test_wrapper_precom_failures_are_fixed_and_never_call_worker(tmp_path, case, stage, reason):
    directory, request, controller = wrapper_packet(tmp_path)
    request_path = directory / 'request.json'
    worker = process(201, controller['pid'], 2.0)
    identities = lambda pid: controller if pid == os.getppid() else worker
    if case == 'request_file':
        request_path.unlink()
    elif case == 'request_json':
        request_path.write_text('{', encoding='utf-8')
    elif case == 'control':
        request['fresh_tiny3_control'].pop('run_id')
        request_path.write_text(json.dumps(request), encoding='utf-8')
    elif case == 'manifest':
        case_path = directory / 'case.json'
        case_path.write_text('{', encoding='utf-8')
        request['manifest_sha256'] = common.sha(case_path)
        request_path.write_text(json.dumps(request), encoding='utf-8')
    elif case == 'marker':
        marker = Path(request['fresh_tiny3_control']['approval_marker'])
        marker.write_text('{', encoding='utf-8')
        request['fresh_tiny3_control']['approval_marker_sha256'] = common.sha(marker)
        request_path.write_text(json.dumps(request), encoding='utf-8')
    elif case == 'profile':
        control = request['fresh_tiny3_control']
        profile_path = Path(control['profile_path'])
        profile_path.write_text('{"stored":1}', encoding='utf-8')
        new_sha = common.sha(profile_path)
        request['profile'] = {'requested': 2}
        request['profile_sha256'] = control['profile_sha256'] = new_sha
        marker = Path(control['approval_marker'])
        approval_keys = {'schema', 'approval_id', 'run_id', 'source_sha256', 'original_sha256',
                         'profile_sha256', 'created_utc'}
        marker.write_text(json.dumps({key: control[key] for key in approval_keys}), encoding='utf-8')
        control['approval_marker_sha256'] = common.sha(marker)
        request_path.write_text(json.dumps(request), encoding='utf-8')
    elif case == 'parent':
        identities = lambda pid: process(999, 50, 99.0)
    elif case == 'identity_error':
        def identities(pid):
            raise RuntimeError('SECRET_PROCESS_DETAIL')
    elif case == 'proof_collision':
        (directory / 'python-worker.json').write_text('{}', encoding='utf-8')
    calls = []
    with pytest.raises(BaseException):
        wrapper.main(request_path, process_identity=identities,
                     worker_main=lambda path: calls.append(path))
    assert not calls and not (directory / 'launch-authorization.json').exists()
    assert read_failure(directory) == fresh.precom_failure(stage, reason)
    assert 'SECRET_' not in (directory / 'precom-failure.json').read_text(encoding='utf-8')


@pytest.mark.parametrize('case,stage,reason', [
    ('stop', 'WORKER_AUTHORIZATION_WAIT', 'PARENT_STOPPED'),
    ('timeout', 'WORKER_AUTHORIZATION_WAIT', 'AUTHORIZATION_TIMEOUT'),
    ('invalid', 'WORKER_AUTHORIZATION_VALIDATE', 'AUTHORIZATION_INVALID'),
])
def test_wrapper_authorization_failures_are_precom_and_no_legacy_call(
        tmp_path, monkeypatch, case, stage, reason):
    directory, request, controller = wrapper_packet(tmp_path)
    worker = process(201, controller['pid'], 2.0)
    calls = []
    if case == 'stop':
        (directory / 'STOP').touch()
        pause = lambda seconds: None
    elif case == 'timeout':
        times = iter([0.0, 16.0])
        monkeypatch.setattr(wrapper.time, 'monotonic', lambda: next(times))
        pause = lambda seconds: None
    else:
        def pause(seconds):
            path = directory / 'launch-authorization.json'
            if not path.exists():
                fresh._write_new_json(path, {'schema': 'WRONG'})
    with pytest.raises(BaseException):
        wrapper.main(directory / 'request.json',
                     process_identity=lambda pid: controller if pid == os.getppid() else worker,
                     worker_main=lambda path: calls.append(path), pause=pause)
    assert not calls
    assert read_failure(directory) == fresh.precom_failure(stage, reason)


@pytest.mark.parametrize('signal,expected', [
    ('stop', fresh.precom_failure('WORKER_HANDOFF_CHECK', 'PARENT_STOPPED')),
    ('failure', fresh.precom_failure('WORKER_MARKER_VALIDATE', 'APPROVAL_MARKER_INVALID')),
])
def test_wrapper_rechecks_stop_and_first_failure_after_authorization(
        tmp_path, signal, expected):
    directory, request, controller = wrapper_packet(tmp_path)
    worker = process(201, controller['pid'], 2.0)
    calls = []

    def write_new(path, value):
        fresh._write_new_json(path, value)
        if Path(path).name == 'python-worker.json':
            fresh._write_new_json(directory / 'launch-authorization.json', dict(
                schema=fresh.FRESH_SCHEMA,
                request_sha256=common.sha(directory / 'request.json'),
                worker_sha256=common.sha(path)))
            if signal == 'stop':
                (directory / 'STOP').touch()
            else:
                fresh.write_precom_failure(directory, 'WORKER_MARKER_VALIDATE',
                                           'APPROVAL_MARKER_INVALID')

    with pytest.raises(fresh.Blocked):
        wrapper.main(directory / 'request.json',
                     process_identity=lambda pid: controller if pid == os.getppid() else worker,
                     worker_main=lambda path: calls.append(path), write_new=write_new)
    assert not calls and (directory / 'launch-authorization.json').is_file()
    assert read_failure(directory) == expected


def test_unknown_child_exception_is_redacted_as_unknown(tmp_path, monkeypatch):
    directory, request, controller = wrapper_packet(tmp_path)
    calls = []
    monkeypatch.setattr(wrapper, 'validate_control',
                        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError('SECRET_TRACE')))
    with pytest.raises(RuntimeError, match='SECRET_TRACE'):
        wrapper.main(directory / 'request.json', process_identity=lambda pid: controller,
                     worker_main=lambda path: calls.append(path))
    assert not calls
    assert read_failure(directory) == fresh.precom_failure(
        'WORKER_CONTROL_VALIDATE', 'PRECOM_UNKNOWN')


def supervisor_packet(tmp_path):
    directory = tmp_path / 'case'; directory.mkdir()
    controller = process(os.getpid(), 1, 1.0)
    launcher = process(321, controller['pid'], 2.0)
    request = dict(run_id=RUN_ID, case_id='tiny3', fresh_tiny3_control={'controller': controller})
    return directory, request, controller, launcher


@pytest.mark.parametrize('case,stage,reason', [
    ('request_write', 'PARENT_REQUEST_CREATE', 'REQUEST_CREATE_FAILED'),
    ('controller_missing', 'PARENT_CONTROLLER_IDENTITY', 'PROCESS_IDENTITY_UNAVAILABLE'),
    ('controller_mismatch', 'PARENT_CONTROLLER_IDENTITY', 'CONTROLLER_IDENTITY_MISMATCH'),
    ('popen', 'PARENT_PROCESS_START', 'PROCESS_START_FAILED'),
    ('launcher_missing', 'PARENT_LAUNCHER_IDENTITY', 'LAUNCHER_IDENTITY_UNAVAILABLE'),
    ('launcher_parent', 'PARENT_LAUNCHER_IDENTITY', 'LAUNCHER_PARENT_MISMATCH'),
    ('child_exit', 'PARENT_WRAPPER_PROOF_WAIT', 'CHILD_EXITED_WITHOUT_PROOF'),
    ('proof_timeout', 'PARENT_WRAPPER_PROOF_WAIT', 'WRAPPER_PROOF_TIMEOUT'),
    ('proof_json', 'PARENT_WRAPPER_PROOF_READ', 'PROOF_JSON_INVALID'),
    ('proof_shape', 'PARENT_WRAPPER_PROOF_VERIFY', 'PROOF_BINDING_MISMATCH'),
    ('proof_request', 'PARENT_WRAPPER_PROOF_VERIFY', 'PROOF_BINDING_MISMATCH'),
    ('worker_process', 'PARENT_WRAPPER_PROOF_VERIFY', 'WORKER_PROCESS_MISMATCH'),
    ('authorization', 'PARENT_AUTHORIZATION', 'AUTHORIZATION_CREATE_FAILED'),
])
def test_supervisor_parent_precom_matrix(tmp_path, monkeypatch, case, stage, reason):
    directory, request, controller, launcher = supervisor_packet(tmp_path)
    now = [0.0]
    class Proc:
        pid = 321
        returncode = 2 if case == 'child_exit' else None
        def poll(self): return self.returncode
    def popen(command, **kwargs):
        if case == 'popen': raise OSError('SECRET_LAUNCH_PATH')
        if case == 'proof_json':
            (directory / 'python-worker.json').write_text('{', encoding='utf-8')
        elif case == 'proof_shape':
            fresh._write_new_json(directory / 'python-worker.json', [])
        elif case in ('proof_request', 'worker_process', 'authorization'):
            observed = launcher.copy()
            if case == 'worker_process': observed['created'] = 999.0
            fresh._write_new_json(directory / 'python-worker.json', dict(
                schema=fresh.FRESH_SCHEMA,
                request_sha256='0' * 64 if case == 'proof_request' else common.sha(directory / 'request.json'),
                process=dict(**observed, sys_executable=observed['path'])))
        return Proc()
    def identity(pid):
        if pid == os.getpid():
            if case == 'controller_missing': raise ProcessLookupError('SECRET_CONTROLLER')
            if case == 'controller_mismatch': return dict(controller, created=999.0)
            return controller
        if case == 'launcher_missing': raise ProcessLookupError('SECRET_PID')
        if case == 'launcher_parent': return dict(launcher, ppid=999)
        return launcher
    def clock(): return now[0]
    def pause(seconds): now[0] += seconds
    writes = []
    def write_new(path, value):
        writes.append(Path(path).name)
        if case == 'request_write' and Path(path).name == 'request.json':
            raise OSError('SECRET_REQUEST_PATH')
        if case == 'authorization' and Path(path).name == 'launch-authorization.json':
            raise OSError('SECRET_AUTH_PATH')
        fresh._write_new_json(path, value)
    monkeypatch.setattr(functions24, 'worker_exit_observation',
                        lambda value: {'outcome': 'PROCESS_EXITED'})
    result = fresh.supervise(directory, request, popen=popen, process_identity=identity,
                             clock=clock, pause=pause, write_new=write_new)
    assert result['status'] == 'BLOCKED' and result['reason'] == 'PRECOM_FAILURE'
    assert result['failure_stage'] == stage and result['reason_code'] == reason
    assert result['first_error'] == {'type': 'PreComFailure', 'message': reason}
    assert result['precom_failure'] == fresh.precom_failure(stage, reason)
    assert not (directory / 'launch-authorization.json').exists()
    assert (directory / 'STOP').exists() and result['force_termination'] is False
    public_text = json.dumps(result['precom_failure']) + json.dumps(result['first_error'])
    assert 'SECRET_' not in public_text


@pytest.mark.parametrize('case,stage,reason', [
    ('proof_failure', 'WORKER_PROFILE_VALIDATE', 'PROFILE_BINDING_INVALID'),
    ('proof_failure_and_stop', 'WORKER_PROFILE_VALIDATE', 'PROFILE_BINDING_INVALID'),
    ('proof_stop', 'PARENT_AUTHORIZATION', 'PARENT_STOPPED'),
    ('proof_dead', 'PARENT_AUTHORIZATION', 'CHILD_EXITED_BEFORE_AUTHORIZATION'),
    ('identity_changed', 'PARENT_AUTHORIZATION', 'WORKER_PROCESS_MISMATCH'),
    ('identity_missing', 'PARENT_AUTHORIZATION', 'PROCESS_IDENTITY_UNAVAILABLE'),
    ('race_stop', 'PARENT_AUTHORIZATION', 'PARENT_STOPPED'),
    ('race_failure', 'WORKER_MARKER_VALIDATE', 'APPROVAL_MARKER_INVALID'),
    ('invalid_failure', 'PARENT_AUTHORIZATION', 'PRECOM_UNKNOWN'),
])
def test_supervisor_never_authorizes_after_proof_when_failure_stop_or_exit_observed(
        tmp_path, monkeypatch, case, stage, reason):
    directory, request, controller, launcher = supervisor_packet(tmp_path)
    calls = {'worker_identity': 0}

    class Proc:
        pid = 321
        returncode = 2 if case == 'proof_dead' else None
        def poll(self): return self.returncode

    def popen(command, **kwargs):
        fresh._write_new_json(directory / 'python-worker.json', dict(
            schema=fresh.FRESH_SCHEMA, request_sha256=common.sha(directory / 'request.json'),
            process=dict(**launcher, sys_executable=launcher['path'])))
        if case in ('proof_failure', 'proof_failure_and_stop'):
            fresh.write_precom_failure(directory, 'WORKER_PROFILE_VALIDATE',
                                       'PROFILE_BINDING_INVALID')
            if case == 'proof_failure_and_stop': (directory / 'STOP').touch()
        elif case == 'proof_stop':
            (directory / 'STOP').touch()
        elif case == 'invalid_failure':
            (directory / 'precom-failure.json').write_text('{"unsafe":"SECRET"}', encoding='utf-8')
        return Proc()

    def identity(pid):
        if pid == os.getpid(): return controller
        calls['worker_identity'] += 1
        if calls['worker_identity'] == 1:
            return launcher
        if case == 'identity_changed': return dict(launcher, created=999.0)
        if case == 'identity_missing': raise ProcessLookupError('SECRET_GONE')
        if case == 'race_stop': (directory / 'STOP').touch()
        if case == 'race_failure':
            fresh.write_precom_failure(directory, 'WORKER_MARKER_VALIDATE',
                                       'APPROVAL_MARKER_INVALID')
        return launcher

    monkeypatch.setattr(functions24, 'worker_exit_observation',
                        lambda value: {'outcome': 'PROCESS_EXITED'})
    result = fresh.supervise(directory, request, popen=popen, process_identity=identity)
    assert result['status'] == 'BLOCKED' and result['reason'] == 'PRECOM_FAILURE'
    assert result['failure_stage'] == stage and result['reason_code'] == reason
    assert result['first_error']['message'] == reason
    assert not (directory / 'launch-authorization.json').exists()
    assert (directory / 'STOP').exists() and result['force_termination'] is False
    assert 'SECRET_' not in json.dumps(result['precom_failure'])


def test_supervisor_uses_valid_child_failure_and_preserves_it_on_exit_observation_error(
        tmp_path, monkeypatch):
    directory, request, controller, launcher = supervisor_packet(tmp_path)
    class Proc:
        pid, returncode = 321, 2
        def poll(self): return self.returncode
    def popen(command, **kwargs):
        fresh.write_precom_failure(directory, 'WORKER_MANIFEST_VALIDATE', 'MANIFEST_INVALID')
        return Proc()
    result = fresh.supervise(directory, request, popen=popen,
                             process_identity=lambda pid: controller if pid == os.getpid() else launcher)
    assert result['failure_stage'] == 'WORKER_MANIFEST_VALIDATE'
    assert result['reason_code'] == 'MANIFEST_INVALID'
    assert result['first_error']['message'] == 'MANIFEST_INVALID'


def test_secondary_exit_observation_does_not_overwrite_first_failure(tmp_path, monkeypatch):
    directory, request, controller, launcher = supervisor_packet(tmp_path)
    class Proc:
        pid, returncode = 321, 2
        def poll(self): return self.returncode
    def popen(command, **kwargs):
        fresh._write_new_json(directory / 'python-worker.json', dict(
            schema=fresh.FRESH_SCHEMA, request_sha256=common.sha(directory / 'request.json'),
            process=dict(**launcher, sys_executable='/different/python')))
        return Proc()
    monkeypatch.setattr(functions24, 'worker_exit_observation',
                        lambda value: (_ for _ in ()).throw(RuntimeError('SECONDARY_SECRET')))
    result = fresh.supervise(directory, request, popen=popen,
                             process_identity=lambda pid: controller if pid == os.getpid() else launcher)
    assert result['failure_stage'] == 'PARENT_EXECUTABLE_VERIFY'
    assert result['reason_code'] == result['first_error']['message'] == 'EXECUTABLE_MISMATCH'
    assert result['python_processes']['worker_exit']['outcome'] == 'EXIT_OBSERVATION_FAILED'
    assert 'SECONDARY_SECRET' not in json.dumps(result)


def test_supervisor_malformed_child_failure_falls_back_without_inference(tmp_path):
    directory, request, controller, launcher = supervisor_packet(tmp_path)
    class Proc:
        pid, returncode = 321, 2
        def poll(self): return self.returncode
    def popen(command, **kwargs):
        (directory / 'precom-failure.json').write_text('{"stage":"SECRET_PATH"}', encoding='utf-8')
        return Proc()
    result = fresh.supervise(directory, request, popen=popen,
                             process_identity=lambda pid: controller if pid == os.getpid() else launcher)
    assert result['failure_stage'] == 'PARENT_WRAPPER_PROOF_WAIT'
    assert result['reason_code'] == 'CHILD_EXITED_WITHOUT_PROOF'


def test_execute_and_cli_preserve_sanitized_first_precom_failure(packet, monkeypatch, capsys):
    failure = fresh.precom_failure('PARENT_PROCESS_START', 'PROCESS_START_FAILED')
    result = execute(packet, monitor=lambda directory, request: dict(
        status='BLOCKED', reason='PRECOM_FAILURE', precom_failure=failure,
        first_error={'type': 'PreComFailure', 'message': 'PROCESS_START_FAILED'},
        worker_still_running=False))
    assert result['status'] == 'BLOCKED' and result['runtime_status'] == 'BLOCKED'
    assert result['reason'] == 'PRECOM_FAILURE'
    assert result['failure_stage'] == 'PARENT_PROCESS_START'
    assert result['reason_code'] == 'PROCESS_START_FAILED'
    assert result['gate2_acceptance'] is False and result['exit_code'] == 2


def test_cli_precom_summary_exposes_only_fixed_codes(monkeypatch, capsys):
    private = 'C:\\SECRET_USER\\profile.json'
    value = dict(schema=fresh.FRESH_SCHEMA, status='BLOCKED', reason='PRECOM_FAILURE',
                 success=False, exit_code=2, case_id='tiny3', runtime_status='BLOCKED',
                 failure_stage='PARENT_PROCESS_START', reason_code='PROCESS_START_FAILED',
                 diagnostic_only=True, gate2_acceptance=False, product_acceptance=False,
                 private_path=private, next_action='STOP_PM_REVIEW')
    monkeypatch.setattr(fresh, 'execute', lambda *args, **kwargs: value)
    assert fresh.main(['--config', private, '--output', 'C:\\SECRET_USER\\out',
                       '--expected-source-sha', 'a' * 64,
                       '--approval-id', APPROVAL, '--owner-confirmed']) == 2
    captured = capsys.readouterr()
    public = json.loads(captured.out)
    assert public['failure_stage'] == 'PARENT_PROCESS_START'
    assert public['reason_code'] == 'PROCESS_START_FAILED'
    assert private not in captured.out and 'private_path' not in public and captured.err == ''


def test_successful_wrapper_handshake_creates_no_failure_diagnostic(tmp_path):
    directory, request, controller = wrapper_packet(tmp_path)
    worker = process(201, controller['pid'], 2.0)
    calls = []
    def authorize(seconds):
        proof = directory / 'python-worker.json'
        path = directory / 'launch-authorization.json'
        if not path.exists():
            fresh._write_new_json(path, dict(schema=fresh.FRESH_SCHEMA,
                                  request_sha256=common.sha(directory / 'request.json'),
                                  worker_sha256=common.sha(proof)))
    assert wrapper.main(directory / 'request.json',
                        process_identity=lambda pid: controller if pid == os.getppid() else worker,
                        worker_main=lambda path: calls.append(path) or 0, pause=authorize) == 0
    assert calls == [directory / 'request.json']
    assert not (directory / 'precom-failure.json').exists()
