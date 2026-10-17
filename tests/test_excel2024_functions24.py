"""Linux protocol mocks only; opaque saves are NOT XLSX/cache evidence."""
from pathlib import Path
from types import SimpleNamespace
import json
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

import pytest

from tools.excel2024_validation import SCHEMA, F24_DIAGNOSTIC, F24_PROTOCOL, E1_DIAGNOSTIC
from tools.excel2024_validation import cases, common, functions24 as f24, run, worker


def expected(challenge=False):
    # Independent literal acceptance constants, not a calculation from Excel.
    values = {}
    for row, duplicate, count in zip(range(8, 14), [0, 1, 0, 1, 0, 1], [2, 2, 0, 2, 1, 4]):
        values.update({f'確認!I{row}': duplicate, f'確認!J{row}': '一致',
                       f'確認!L{row}': count, f'確認!M{row}': '一致'})
    if challenge:
        values.update({'確認!L8': 1, '確認!M8': '不一致'})
    return values


def profile():
    return dict(schema=SCHEMA, product='Microsoft Office Home and Business 2024', licensed=True,
                ui_confirmed_on_owned_pid=True,
                facts=dict(pid=123, created=1000, path=r'C:\Office\EXCEL.EXE', version='16.0',
                           exe_sha256='MOCK', file_version='16.0.1.2', build='1',
                           installation=dict(product_release_ids=['HomeBusiness2024Retail'])))


def python_proof():
    parent = dict(pid=1, ppid=0, created=1, path='supervisor.exe')
    launcher = dict(pid=2, ppid=1, created=2, path='venv-python.exe')
    child = dict(pid=3, ppid=2, created=3, path='python.exe', sys_executable='venv-python.exe')
    return dict(worker_pid=2, python_process=child.copy(),
                python_processes=dict(supervisor=parent, launcher=launcher, worker=child.copy(),
                                      worker_exit=dict(outcome='PROCESS_EXITED')))


class FakeSession:
    """Protocol stimulus; never evaluates a workbook or writes formula caches."""
    def __init__(self, problem=None):
        self.problem, self.calls, self.current = problem, [], None
        self.b8, self.rebuilds, self.saved_flag = 100, 0, None
        self.values = expected()
        owner = self

        class Cell:
            def __init__(self, cell): self.cell = cell
            @property
            def Value2(self): return owner.b8 if self.cell == 'B8' else owner.values['確認!' + self.cell]
            @Value2.setter
            def Value2(self, value):
                assert self.cell == 'B8'
                owner.calls.append(('edit', value))
                owner.b8 = value

        class Book:
            Date1904 = False
            def __init__(self, path, readonly):
                self.FullName, self.Name, self.ReadOnly = path, Path(path).name, readonly
                if owner.problem == 'readonly' and not readonly: self.ReadOnly = True
                self.flag = owner.saved_flag if readonly else True
                if owner.problem == 'reopened_true' and readonly: self.flag = True
                self.Worksheets = SimpleNamespace(Item=lambda s: SimpleNamespace(Range=Cell))
            @property
            def ForceFullCalculation(self): return self.flag
            @ForceFullCalculation.setter
            def ForceFullCalculation(self, value):
                owner.calls.append(('force', value))
                if owner.problem == 'refuse_false': return
                if owner.problem == 'integer_zero': self.flag = 0; return
                if owner.problem == 'second_false' and owner.rebuilds == 1: self.flag = True; return
                self.flag = value
            def SaveAs(self, path, FileFormat):
                assert FileFormat == 51
                owner.calls.append('save')
                if owner.problem == 'save': raise OSError('SaveAs failed')
                owner.saved_flag = self.flag
                Path(path).write_bytes(b'MOCK ONLY / NOT XLSX / NO FORMULA CACHES')
            def Close(self, save):
                assert save is False
                owner.calls.append('close')
                if owner.problem == 'close': raise RuntimeError('secondary Close failure')
                owner.current = None

        class Collection:
            @property
            def Count(self): return 2 if owner.problem == 'extra_book' else int(owner.current is not None)
            def Item(self, i): return owner.current
            def Open(self, path, **kwargs):
                owner.calls.append(('open', kwargs['ReadOnly']))
                owner.current = Book(path, kwargs['ReadOnly'])
                return owner.current

        class Excel:
            Workbooks = Collection()
            def __init__(self): self.mode = -4105
            @property
            def Calculation(self): return self.mode
            @Calculation.setter
            def Calculation(self, value):
                owner.calls.append(('mode', value))
                if owner.problem == 'manual' and value == -4135: return
                if owner.problem == 'automatic' and value == -4105: return
                self.mode = value
            @property
            def CalculationState(self):
                if owner.rebuilds and owner.problem in ('pending', 'calculating', 'unknown_state'):
                    return {'pending': 2, 'calculating': 1, 'unknown_state': 3}[owner.problem]
                return 0
            def CalculateFullRebuild(self):
                owner.calls.append('full_rebuild')
                if owner.problem == 'rpc':
                    error = RuntimeError('RPC rejected'); error.hresult = -2147418111; raise error
                owner.rebuilds += 1
                owner.values = expected(challenge=owner.b8 == 0)
                if owner.problem == 'stale_challenge': owner.values = expected()
                if owner.problem == 'stale_restore' and owner.b8 == 100: owner.values = expected(True)
        self.excel = Excel()
    def confirm_owner(self):
        self.calls.append('identity')
        if self.problem == 'ownership': raise common.GateError('BLOCKED', 'ownership mismatch')


def execute_worker(directory, session=None):
    manifest = cases.prepare('functions24', directory, 'test')
    s = session or FakeSession()
    evidence = []
    result = worker.execute_case(s, manifest, directory, common.Progress(directory, 'test', 'functions24'), [],
                                 diagnostic=F24_DIAGNOSTIC, evidence=evidence)
    result['functions24'] = dict(diagnostic=F24_DIAGNOSTIC, protocol=F24_PROTOCOL, evidence=evidence)
    return result, manifest, s


def runtime(directory, request):
    common.write_json(directory / 'request.json', request)
    manifest = common.read_json(directory / 'case.json')
    # Exercise actual worker orchestration against an in-memory COM mock.
    s = FakeSession()
    evidence = []
    result = worker.execute_case(s, manifest, directory,
                                 common.Progress(directory, request['run_id'], 'functions24'), [],
                                 diagnostic=F24_DIAGNOSTIC, evidence=evidence)
    result.update(schema=SCHEMA, run_id=request['run_id'], case_id='functions24',
                  identity=profile()['facts'], owned_pid=123, first_error=None, cleanup_errors=[],
                  progress_log_errors=[], worker_still_running=False, force_termination=False,
                  request_sha256=common.sha(directory / 'request.json'),
                  manifest_sha256=common.sha(directory / 'case.json'), **python_proof())
    result['functions24'] = dict(diagnostic=F24_DIAGNOSTIC, protocol=F24_PROTOCOL, evidence=evidence)
    return result


def execute(tmp_path, monkeypatch, monitor=runtime, residual=lambda r: False, units=lambda r: dict(returncode=0)):
    config = tmp_path / 'profile.json'
    common.write_json(config, profile())
    def fake_verify(directory, manifest):
        return dict(saved_sha256=common.sha(directory / 'results' / manifest['file']), numeric_comparison='MOCK-ONLY')
    monkeypatch.setattr(run, 'verify_saved', fake_verify)
    return run.execute('functions24-false', config, tmp_path / 'run', check=lambda: None,
                       monitor=monitor, residual=residual, units=units)


def test_standalone_only_functions24_never_a_gate2_acceptance(tmp_path, monkeypatch):
    calls = []
    def monitor(directory, request):
        calls.append(request)
        assert request['diagnostic'] == F24_DIAGNOSTIC
        return runtime(directory, request)
    value, root = execute(tmp_path, monkeypatch, monitor)
    assert [r['case_id'] for r in calls] == ['functions24']
    assert value['status'] == 'DIAGNOSTIC_PASS' and value['exit_code'] == 0
    assert value['diagnostic_recalculation_verified'] is True
    assert value['gate2_acceptance'] is False and value['product_acceptance'] is False
    assert value['excel_recalculation_verified'] is False and value['diagnostic_only'] is True
    assert [s['status'] for s in value['stages']] == ['NOT_RUN', 'PASS', 'NOT_RUN']
    assert not (root / 'tiny3').exists() and not (root / 'annual1').exists()
    assert common.read_json(root / 'run.json') == value


def test_false_both_phases_same_order_24_literal_values_and_original_preserved(tmp_path):
    original = cases.REPO / 'output/s1_compatibility/01_function_gate.xlsx'
    before = original.read_bytes()
    result, manifest, s = execute_worker(tmp_path)
    protocol = [c for c in s.calls if c == 'full_rebuild' or isinstance(c, tuple) and c[0] in ('mode', 'force', 'edit')]
    assert protocol == [('mode', -4135), ('edit', 0), ('force', False), 'full_rebuild', ('mode', -4105),
                        ('mode', -4135), ('edit', 100), ('force', False), 'full_rebuild', ('mode', -4105)]
    assert result['full_rebuild_calls'] == 2 and s.calls.count('save') == 1 and s.calls.count('close') == 2
    assert result['samples'][1]['values'] == expected(True)
    assert result['samples'][2]['values'] == expected() == result['samples'][3]['values']
    assert len(expected()) == 24
    assert original.read_bytes() == before == (tmp_path / 'inputs' / manifest['file']).read_bytes()
    assert (tmp_path / 'work' / manifest['file']).read_bytes() == before
    assert common.sha(original) == 'ea158ca2052c1c2ba4d5b7bd30e4304846770e82392bea4229506d973f131e21'
    with zipfile.ZipFile(original) as z:
        calc = ET.fromstring(z.read('xl/workbook.xml')).find('m:calcPr', cases.NS)
        assert calc.get('forceFullCalc') == '1'


def test_normal_functions24_keeps_true_twice_without_diagnostic_observations(tmp_path):
    manifest = cases.prepare('functions24', tmp_path, 'test')
    s = FakeSession()
    result = worker.execute_case(s, manifest, tmp_path, common.Progress(tmp_path, 'test', 'functions24'), [])
    assert result['status'] == 'SAVED_REOPENED' and result['full_rebuild_calls'] == 2
    assert [c for c in s.calls if isinstance(c, tuple) and c[0] == 'force'] == [('force', True), ('force', True)]
    assert not any(e['event'].startswith(('functions24.', 'e1.')) for e in run.events(tmp_path / 'progress.jsonl'))


@pytest.mark.parametrize('problem', ['refuse_false', 'integer_zero', 'manual', 'automatic', 'rpc', 'ownership',
                                     'readonly', 'extra_book', 'stale_challenge', 'stale_restore', 'save', 'close',
                                     'reopened_true', 'unknown_state', 'second_false'])
def test_worker_failures_stop_without_next_phase_or_false_success(tmp_path, problem):
    manifest = cases.prepare('functions24', tmp_path, 'test')
    s = FakeSession(problem)
    with pytest.raises((common.GateError, RuntimeError, OSError)) as error:
        worker.execute_case(s, manifest, tmp_path, common.Progress(tmp_path, 'test', 'functions24'), [],
                            diagnostic=F24_DIAGNOSTIC, evidence=[])
    if problem in ('refuse_false', 'integer_zero', 'manual', 'readonly', 'extra_book', 'ownership'):
        assert 'full_rebuild' not in s.calls
    if problem == 'stale_challenge':
        assert ('edit', 100) not in s.calls and s.rebuilds == 1
    if problem == 'second_false': assert s.rebuilds == 1
    if problem not in ('reopened_true', 'close', 'save'):
        assert 'save' not in s.calls
    if problem == 'rpc': assert error.value.hresult == -2147418111


@pytest.mark.parametrize('problem', ['pending', 'calculating'])
def test_pending_or_calculating_even_with_correct_values_times_out_without_restore(tmp_path, monkeypatch, problem):
    now = [0.0]
    monkeypatch.setattr(worker.time, 'monotonic', lambda: now[0])
    monkeypatch.setattr(worker.time, 'sleep', lambda n: now.__setitem__(0, now[0] + 60))
    monkeypatch.setattr(worker, 'pump_messages', lambda: None)
    s = FakeSession(problem)
    with pytest.raises(common.GateError) as caught:
        execute_worker(tmp_path, s)
    assert caught.value.status == 'TIMEOUT' and now[0] == 180
    assert s.values == expected(True) and s.rebuilds == 1
    assert ('edit', 100) not in s.calls and 'save' not in s.calls


@pytest.mark.parametrize('problem', ['state', 'state_before', 'mode', 'value', 'cleanup', 'primary', 'progress',
                                     'rebuild', 'request', 'manifest', 'worker_alive', 'flag', 'flag_type',
                                     'phase_order', 'missing_phase', 'probe', 'difference', 'book', 'book_path',
                                     'readonly', 'protocol', 'diagnostic', 'sku', 'license', 'owned_pid',
                                     'python_pid', 'python_path', 'python_missing', 'saved_sha', 'python_alive'])
def test_parent_rejects_incomplete_or_tampered_evidence(tmp_path, monkeypatch, problem):
    def monitor(directory, request):
        r = runtime(directory, request)
        item = r['functions24']['evidence'][2]
        if problem in ('state', 'state_before', 'mode'): r['samples'][1][problem] = -4135 if problem == 'mode' else 2
        if problem == 'value': r['samples'][1]['values']['確認!L8'] = 2
        if problem == 'cleanup': r['cleanup_errors'] = [dict(message='quit failure')]
        if problem == 'primary': r['first_error'] = dict(message='COM failure')
        if problem == 'progress': r['progress_log_errors'] = [dict(message='logging failure')]
        if problem == 'rebuild': r['full_rebuild_calls'] = 1
        if problem in ('request', 'manifest'): r[problem + '_sha256'] = 'wrong'
        if problem == 'worker_alive': r['worker_still_running'] = True
        if problem == 'flag': item['force_full'] = True
        if problem == 'flag_type': item['force_full'] = 0
        if problem == 'phase_order': r['functions24']['evidence'].reverse()
        if problem == 'missing_phase': r['functions24']['evidence'].pop()
        if problem == 'probe': item['sample']['values'].pop('確認!I8')
        if problem == 'difference': item['sample']['differences'] = [dict(invented=True)]
        if problem == 'book': item['workbooks'].append(item['workbooks'][0])
        if problem == 'book_path': item['workbooks'][0]['path'] = 'other.xlsx'
        if problem == 'readonly': item['workbooks'][0]['read_only'] = True
        if problem == 'protocol': r['functions24']['protocol'] = 'E0'
        if problem == 'diagnostic': r['functions24']['diagnostic'] = 'E2'
        if problem == 'sku': r['identity']['installation']['product_release_ids'] = ['O365ProPlusRetail']
        if problem == 'license': request['profile']['licensed'] = False
        if problem == 'owned_pid': r['owned_pid'] = 999
        if problem == 'python_pid': r['python_process']['ppid'] = 999
        if problem == 'python_path': r['python_process']['path'] = 'other.exe'
        if problem == 'python_missing': del r['python_processes']['worker']
        if problem == 'python_alive': r['python_processes']['worker_exit']['outcome'] = 'WORKER_STILL_RUNNING'
        if problem == 'saved_sha': r['saved_sha256'] = 'tampered'
        return r
    value, _ = execute(tmp_path, monkeypatch, monitor)
    assert value['status'] in ('FAIL', 'BLOCKED') and value['exit_code'] != 0
    assert value['diagnostic_recalculation_verified'] is False
    assert value['stages'][0]['status'] == value['stages'][2]['status'] == 'NOT_RUN'


@pytest.mark.parametrize('failure', ['remains', 'access', 'invalid', 'cache', 'interrupt'])
def test_cache_and_process_exit_are_mandatory(tmp_path, monkeypatch, failure):
    def residual(result):
        if failure == 'access': raise PermissionError('denied')
        if failure == 'interrupt': raise KeyboardInterrupt()
        return None if failure == 'invalid' else failure == 'remains'
    def monitor(directory, request):
        r = runtime(directory, request)
        if failure == 'cache':
            def reject(*a): raise ValueError('missing or wrong cache')
            monkeypatch.setattr(run, 'verify_saved', reject)
        return r
    value, _ = execute(tmp_path, monkeypatch, monitor, residual)
    assert value['status'] == ('INTERRUPTED' if failure == 'interrupt' else 'FAIL') and value['exit_code'] == 1


@pytest.mark.parametrize('status', ['TIMEOUT', 'BLOCKED', 'FAIL', 'INTERRUPTED'])
def test_original_failure_kept_with_secondary_cleanup_and_residual(tmp_path, monkeypatch, status):
    def monitor(directory, request):
        r = runtime(directory, request)
        r.update(status=status, first_error=dict(message='original failure'),
                 cleanup_errors=[dict(message='secondary Close failure')])
        return r
    value, _ = execute(tmp_path, monkeypatch, monitor, lambda r: True)
    stage = value['stages'][1]
    assert value['status'] == status and 'residual_error' in stage
    assert stage['runtime']['first_error']['message'] == 'original failure'
    assert stage['runtime']['cleanup_errors'][0]['message'] == 'secondary Close failure'


@pytest.mark.parametrize('reason', ['preflight', 'profile', 'units', 'internal_profile'])
def test_preconditions_no_excel_launch(tmp_path, reason):
    config = tmp_path / 'profile.json'
    if reason != 'profile': common.write_json(config, profile())
    if reason == 'internal_profile': config = cases.REPO / 'profile.json'
    def forbidden(*a): pytest.fail('must not start Excel')
    value, _ = run.execute('functions24-false', config, tmp_path / 'run',
                           check=lambda: 'no Windows' if reason == 'preflight' else None,
                           monitor=forbidden, units=lambda r: dict(returncode=1))
    assert value['status'] in ('BLOCKED', 'FAIL')
    assert value['stages'][1]['status'] == value['status']
    assert value['stages'][0]['status'] == value['stages'][2]['status'] == 'NOT_RUN'


def test_real_verifier_rejects_uncalculated_original(tmp_path):
    manifest = cases.prepare('functions24', tmp_path, 'test')
    (tmp_path / 'results' / manifest['file']).write_bytes((tmp_path / 'inputs' / manifest['file']).read_bytes())
    with pytest.raises((AssertionError, ValueError)):
        cases.verify_saved(tmp_path, manifest)


@pytest.mark.parametrize('case_id,diagnostic', [('tiny3', F24_DIAGNOSTIC), ('doctor', F24_DIAGNOSTIC),
                                               ('annual1', F24_DIAGNOSTIC), ('functions24', 'E2'),
                                               ('functions24', E1_DIAGNOSTIC)])
def test_invalid_case_diagnostic_rejected_before_connect(tmp_path, monkeypatch, case_id, diagnostic):
    class Session:
        owned = False
        facts = None
        def __init__(self, *a): pass
        def connect(self): pytest.fail('invalid request must not connect')
        def close(self): pass
        def release(self): pass
    monkeypatch.setattr(worker, 'Session', Session)
    request = dict(run_id='test', case_id=case_id, diagnostic=diagnostic,
                   source_sha256=run.source_identity()['source_sha256'])
    common.write_json(tmp_path / 'request.json', request)
    assert worker.main(tmp_path / 'request.json') == 1
    assert common.read_json(tmp_path / 'worker.json')['status'] == 'FAIL'


@pytest.mark.parametrize('key,value', [('edits', [['確認', 'B8', 200]]), ('challenges', []),
                                     ('formula_count', 23), ('probes', []), ('file', 'other.xlsx')])
def test_manifest_scope_guard_rejects_before_edit(tmp_path, key, value):
    manifest = cases.prepare('functions24', tmp_path, 'test')
    manifest[key] = value
    s = FakeSession()
    with pytest.raises(ValueError):
        worker.execute_case(s, manifest, tmp_path, common.Progress(tmp_path, 'test', 'functions24'), [],
                            diagnostic=F24_DIAGNOSTIC, evidence=[])
    assert not s.calls


def test_process_mapping_supports_direct_and_launcher_without_changing_monitor():
    r = python_proof()
    assert f24.validate_processes(r) == 'LAUNCHER_CHILD_WORKER'
    direct = dict(**r['python_processes']['launcher'], sys_executable='venv-python.exe')
    r['python_process'] = direct.copy(); r['python_processes']['worker'] = direct.copy()
    assert f24.validate_processes(r) == 'DIRECT_WORKER'


@pytest.mark.parametrize('problem', ['parent', 'created', 'direct_fingerprint', 'missing_exit'])
def test_process_mapping_rejects_unknown_relationship(problem):
    r = python_proof()
    if problem == 'parent':
        r['python_process']['ppid'] = r['python_processes']['worker']['ppid'] = 999
    if problem == 'created':
        r['python_process']['created'] = r['python_processes']['worker']['created'] = 1
    if problem == 'direct_fingerprint':
        r['python_process']['pid'] = r['python_processes']['worker']['pid'] = 2
    if problem == 'missing_exit': del r['python_processes']['worker_exit']
    with pytest.raises(ValueError): f24.validate_processes(r)


@pytest.mark.parametrize('problem,outcome', [('missing', 'PROCESS_EXITED'), ('reuse', 'PID_REUSED'),
                                           ('live', 'WORKER_STILL_RUNNING'), ('path', None), ('access', None)])
def test_worker_leaf_exit_is_read_only_and_fingerprint_checked(problem, outcome):
    def process(pid):
        assert pid == 3
        if problem == 'missing': raise ProcessLookupError()
        if problem == 'access': raise PermissionError('denied')
        return SimpleNamespace(create_time=lambda: 4 if problem == 'reuse' else 3,
                               exe=lambda: 'other.exe' if problem == 'path' else 'python.exe')
    if outcome:
        assert f24.worker_exit_observation(python_proof()['python_process'], process=process)['outcome'] == outcome
    else:
        with pytest.raises((ValueError, PermissionError)):
            f24.worker_exit_observation(python_proof()['python_process'], process=process)


def test_supervisor_collects_only_own_parent_launcher_and_progress_worker(tmp_path, monkeypatch):
    proof = python_proof()
    own = dict(proof['python_processes']['supervisor'], pid=os.getpid())
    launcher = dict(proof['python_processes']['launcher'], ppid=os.getpid())
    seen = []
    def observe(pid):
        seen.append(pid)
        return own if pid == os.getpid() else launcher
    monkeypatch.setattr(f24, 'process_identity', observe)
    monkeypatch.setattr(f24, 'worker_exit_observation', lambda observed: dict(outcome='PROCESS_EXITED'))
    common.Progress(tmp_path, 'r', 'functions24').emit('python.worker_started', process=dict(proof['python_process'], pid=os.getpid()))
    common.write_json(tmp_path / 'worker.json', dict(status='SAVED_REOPENED'))
    proc = SimpleNamespace(pid=2, returncode=0, poll=lambda: 0)
    r = run.supervise(tmp_path, dict(run_id='r', case_id='functions24', diagnostic=F24_DIAGNOSTIC),
                      popen=lambda *a, **k: proc)
    assert r['status'] == 'SAVED_REOPENED'
    assert seen == [os.getpid(), 2]
    assert r['python_processes']['worker']['pid'] == os.getpid()


@pytest.mark.parametrize('failure', ['still_alive', 'lookup_error', 'timeout', 'interrupt'])
def test_supervisor_never_confuses_launcher_exit_with_worker_exit(tmp_path, monkeypatch, failure):
    monkeypatch.setattr(f24, 'process_identity', lambda pid: dict(pid=pid))
    def observation(*a):
        if failure == 'lookup_error': raise PermissionError('denied')
        if failure == 'interrupt': raise KeyboardInterrupt()
        return dict(outcome='WORKER_STILL_RUNNING')
    monkeypatch.setattr(f24, 'worker_exit_observation', observation)
    common.Progress(tmp_path, 'r', 'functions24').emit('python.worker_started', process=dict(pid=os.getpid()))
    common.write_json(tmp_path / 'worker.json', dict(status='SAVED_REOPENED'))
    proc = SimpleNamespace(pid=2, returncode=0, poll=lambda: None if failure == 'timeout' else 0)
    now = [0.0]
    r = run.supervise(tmp_path, dict(run_id='r', case_id='functions24', diagnostic=F24_DIAGNOSTIC),
                      popen=lambda *a, **k: proc, clock=lambda: now[0],
                      pause=lambda n: now.__setitem__(0, 601))
    assert r['status'] == ('TIMEOUT' if failure == 'timeout' else 'INTERRUPTED' if failure == 'interrupt' else 'FAIL')
    assert (tmp_path / 'STOP').exists() and r['force_termination'] is False
    if failure not in ('lookup_error', 'interrupt'): assert r['worker_still_running'] is True and r['cleanup_pending'] is True


def test_worker_main_records_pid_and_preserves_primary_and_cleanup_failures(tmp_path, monkeypatch):
    manifest = cases.prepare('functions24', tmp_path, 'r')
    calls = []
    class Session(FakeSession):
        def __init__(self, *a):
            super().__init__('stale_challenge')
            self.owned, self.facts = False, None
        def connect(self):
            self.owned = True
            self.facts = profile()['facts']; self.owner = self.facts
        def close(self): calls.append('quit'); raise RuntimeError('secondary Quit failure')
        def release(self): calls.append('release')
    monkeypatch.setattr(worker, 'Session', Session)
    monkeypatch.setattr(f24, 'worker_process', lambda: dict(python_proof()['python_process'], pid=os.getpid()))
    request = dict(run_id='r', case_id='functions24', diagnostic=F24_DIAGNOSTIC, profile=profile(),
                   source_sha256=run.source_identity()['source_sha256'], manifest_sha256=common.sha(tmp_path / 'case.json'))
    common.write_json(tmp_path / 'request.json', request)
    assert worker.main(tmp_path / 'request.json') == 1
    result = common.read_json(tmp_path / 'worker.json')
    assert result['status'] == 'FAIL' and result['first_error']['type'] == 'GateError'
    assert result['cleanup_errors'][0]['message'] == 'secondary Quit failure'
    assert result['python_process']['pid'] == os.getpid()
    assert run.events(tmp_path / 'progress.jsonl')[0]['event'] == 'python.worker_started'
    assert calls == ['quit', 'release']
    assert not (tmp_path / 'results' / manifest['file']).exists()


def test_stop_after_full_rebuild_prevents_restore_save_and_more_com(tmp_path):
    manifest = cases.prepare('functions24', tmp_path, 'test')
    s = FakeSession()
    rebuild = s.excel.CalculateFullRebuild
    def stop():
        rebuild()
        (tmp_path / 'STOP').touch()
    s.excel.CalculateFullRebuild = stop
    with pytest.raises(common.GateError) as caught:
        worker.execute_case(s, manifest, tmp_path, common.Progress(tmp_path, 'test', 'functions24'), [],
                            diagnostic=F24_DIAGNOSTIC, evidence=[])
    assert caught.value.status == 'INTERRUPTED'
    assert s.rebuilds == 1 and ('edit', 100) not in s.calls and 'save' not in s.calls


def test_cli_reports_diagnostic_not_gate2(tmp_path, monkeypatch, capsys):
    value, root = execute(tmp_path, monkeypatch)
    monkeypatch.setattr(run, 'execute', lambda *a: (value, root))
    monkeypatch.setattr(sys, 'argv', ['run', 'functions24-false'])
    assert run.main() == 0
    result = json.loads(capsys.readouterr().out)
    assert result['status'] == 'DIAGNOSTIC_PASS' and result['gate2_acceptance'] is False
    assert result['diagnostic_only'] is True
