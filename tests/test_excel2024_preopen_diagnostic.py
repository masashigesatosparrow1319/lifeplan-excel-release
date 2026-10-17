"""Linux COM mocks only. Opaque saves/verification stubs are not Excel evidence."""
import copy
import json
import os
from types import SimpleNamespace

import pytest

from tools.excel2024_validation import A_DIAGNOSTIC, A_PROTOCOL, F24_DIAGNOSTIC, E1_DIAGNOSTIC, SCHEMA
from tools.excel2024_validation import cases, common, functions24 as f24, preopen_input as builder, run, worker
from tests.test_excel2024_functions24 import FakeSession, expected, profile, python_proof
from tests.test_excel2024_preopen_input import APPROVAL, prepared


class ASession(FakeSession):
    def __init__(self, problem=None, initial=False):
        super().__init__(problem)
        original_open = self.excel.Workbooks.Open
        def open_book(path, **kwargs):
            book = original_open(path, **kwargs)
            if not kwargs['ReadOnly']: book.flag = initial
            return book
        self.excel.Workbooks.Open = open_book


def execute_worker(directory, manifest, session=None):
    session = session or ASession()
    evidence = []
    result = worker.execute_case(session, manifest, directory,
                                 common.Progress(directory, manifest['run_id'], 'functions24'), [],
                                 diagnostic=A_DIAGNOSTIC, evidence=evidence)
    result['functions24'] = dict(diagnostic=A_DIAGNOSTIC, protocol=A_PROTOCOL, evidence=evidence)
    return result, session


def runtime(directory, request):
    common.write_json(directory / 'request.json', request)
    manifest = common.read_json(directory / 'case.json')
    # Parent/worker/input checks are real; Excel execution and numeric cache checker are mocked.
    proof = builder.validate_bundle(directory, manifest, request)
    result, _ = execute_worker(directory, manifest)
    result.update(schema=SCHEMA, run_id=request['run_id'], case_id='functions24',
                  identity=profile()['facts'], owned_pid=123, first_error=None, cleanup_errors=[],
                  progress_log_errors=[], worker_still_running=False, force_termination=False,
                  request_sha256=common.sha(directory / 'request.json'),
                  manifest_sha256=common.sha(directory / 'case.json'), preopen_control=proof, **python_proof())
    return result


def execute(tmp_path, monkeypatch, monitor=runtime, residual=lambda r: False, units=lambda r: dict(returncode=0),
            approval=APPROVAL):
    config = tmp_path / 'profile.json'
    common.write_json(config, profile())
    monkeypatch.setattr(run, 'verify_saved', lambda d, m: dict(
        saved_sha256=common.sha(d / 'results' / m['file']), numeric_comparison='MOCK-ONLY'))
    return run.execute('functions24-preopen-force0', config, tmp_path / 'run',
                       check=lambda: None, monitor=monitor, residual=residual, units=units, approval_id=approval)


def test_a_only_functions24_diagnostic_not_formal_acceptance(tmp_path, monkeypatch):
    calls = []
    def monitor(directory, request):
        calls.append(request)
        assert request['diagnostic'] == A_DIAGNOSTIC and request['protocol'] == A_PROTOCOL
        return runtime(directory, request)
    value, root = execute(tmp_path, monkeypatch, monitor)
    assert value['status'] == 'DIAGNOSTIC_PASS' and value['exit_code'] == 0
    assert [r['case_id'] for r in calls] == ['functions24']
    assert [s['status'] for s in value['stages']] == ['NOT_RUN', 'PASS', 'NOT_RUN']
    assert value['gate2_acceptance'] is value['product_acceptance'] is value['excel_recalculation_verified'] is False
    assert value['diagnostic_only'] is value['diagnostic_recalculation_verified'] is True
    assert not (root / 'tiny3').exists() and not (root / 'annual1').exists()
    assert value['stages'][1]['control_verification']['verified'] is True
    assert common.read_json(root / 'run.json') == value


def test_a_keeps_existing_two_phase_protocol_and_all_24_values(tmp_path):
    directory, manifest = prepared(tmp_path)
    result, s = execute_worker(directory, manifest)
    f24.validate_evidence(dict(result, identity=profile()['facts'], owned_pid=123,
                               force_termination=False, **python_proof()), manifest, directory, diagnostic=A_DIAGNOSTIC)
    protocol = [c for c in s.calls if c == 'full_rebuild' or isinstance(c, tuple) and c[0] in ('mode', 'force', 'edit')]
    assert protocol == [('mode', -4135), ('edit', 0), ('force', False), 'full_rebuild', ('mode', -4105),
                        ('mode', -4135), ('edit', 100), ('force', False), 'full_rebuild', ('mode', -4105)]
    assert result['full_rebuild_calls'] == 2 and s.calls.count('save') == 1 and s.calls.count('close') == 2
    assert result['samples'][1]['values'] == expected(True)
    assert result['samples'][2]['values'] == result['samples'][3]['values'] == expected()
    assert len(result['samples'][3]['values']) == 24
    assert [e['force_full'] for e in result['functions24']['evidence']] == [False] * 11


@pytest.mark.parametrize('initial', [True, 0, None])
def test_opened_must_be_strict_bool_false_before_any_edit(tmp_path, initial):
    directory, manifest = prepared(tmp_path)
    s = ASession(initial=initial)
    with pytest.raises(common.GateError): execute_worker(directory, manifest, s)
    assert not any(isinstance(c, tuple) and c[0] == 'edit' for c in s.calls)
    assert s.rebuilds == 0 and 'save' not in s.calls


def test_initial_manual_mode_rejected_before_edit(tmp_path):
    directory, manifest = prepared(tmp_path)
    s = ASession(); s.excel.mode = -4135
    with pytest.raises(common.GateError): execute_worker(directory, manifest, s)
    assert ('edit', 0) not in s.calls and s.rebuilds == 0


def test_opened_pending_is_allowed_but_completion_requires_zero(tmp_path):
    directory, manifest = prepared(tmp_path)
    s = ASession()
    s.excel.__class__.CalculationState = property(lambda self: 2 if not s.rebuilds else 0)
    result, _ = execute_worker(directory, manifest, s)
    assert result['functions24']['evidence'][0]['sample']['state'] == 2
    assert result['samples'][-1]['state'] == 0 and result['status'] == 'SAVED_REOPENED'


@pytest.mark.parametrize('state', ['pending', 'calculating'])
def test_matching_values_with_nonzero_state_timeout_at_180_without_restore(tmp_path, monkeypatch, state):
    directory, manifest = prepared(tmp_path)
    now = [0.0]
    monkeypatch.setattr(worker.time, 'monotonic', lambda: now[0])
    monkeypatch.setattr(worker.time, 'sleep', lambda n: now.__setitem__(0, now[0] + 60))
    monkeypatch.setattr(worker, 'pump_messages', lambda: None)
    s = ASession(state)
    with pytest.raises(common.GateError) as caught: execute_worker(directory, manifest, s)
    assert caught.value.status == 'TIMEOUT' and now[0] == 180
    assert s.values == expected(True) and s.rebuilds == 1
    assert ('edit', 100) not in s.calls and 'save' not in s.calls


@pytest.mark.parametrize('diagnostic', [None, F24_DIAGNOSTIC, E1_DIAGNOSTIC])
def test_old_paths_reject_a_input_before_any_com(tmp_path, diagnostic):
    directory, manifest = prepared(tmp_path)
    s = ASession()
    with pytest.raises(ValueError):
        worker.execute_case(s, manifest, directory, common.Progress(directory, 'test', 'functions24'), [],
                            diagnostic=diagnostic, evidence=[])
    assert not s.calls


def test_a_rejects_old_original_before_any_com(tmp_path):
    manifest = cases.prepare('functions24', tmp_path, 'test')
    s = ASession()
    with pytest.raises(ValueError): execute_worker(tmp_path, manifest, s)
    assert not s.calls


@pytest.mark.parametrize('fault', ['source', 'request_sha', 'run', 'protocol', 'diagnostic', 'variant', 'case',
                                 'proof', 'work', 'old_f24', 'old_smoke', 'old_e1', 'case_control_only', 'stripped_e1'])
def test_worker_rejects_broken_binding_before_connect(tmp_path, monkeypatch, fault):
    directory, manifest = prepared(tmp_path, source=run.source_identity()['source_sha256'])
    request = dict(run_id='test', case_id='functions24', diagnostic=A_DIAGNOSTIC, protocol=A_PROTOCOL,
                   source_sha256=run.source_identity()['source_sha256'], manifest_sha256=common.sha(directory / 'case.json'),
                   control=copy.deepcopy(manifest['control']))
    if fault == 'source': request['source_sha256'] = 'wrong'
    if fault == 'request_sha': request['manifest_sha256'] = 'wrong'
    if fault == 'run': request['run_id'] = 'wrong'
    if fault == 'protocol': request['protocol'] = 'wrong'
    if fault == 'diagnostic': request['diagnostic'] = 'E2'
    if fault == 'variant': request['control']['variant'] = 'T0'
    if fault == 'case': request['case_id'] = 'tiny3'
    if fault == 'proof': (directory / 'transform.json').write_bytes(b'{}')
    if fault == 'work': (directory / 'work' / manifest['file']).write_bytes(builder.ORIGINAL.read_bytes())
    if fault in ('old_f24', 'old_smoke', 'old_e1', 'case_control_only'):
        request['diagnostic'] = {'old_f24': F24_DIAGNOSTIC, 'old_smoke': None,
                                 'old_e1': E1_DIAGNOSTIC, 'case_control_only': E1_DIAGNOSTIC}[fault]
        if fault in ('old_e1', 'case_control_only'): request['case_id'] = 'tiny3'
        if fault == 'case_control_only': request.pop('control')
    if fault == 'stripped_e1':
        request.update(diagnostic=E1_DIAGNOSTIC, case_id='tiny3')
        request.pop('control'); manifest.pop('control'); manifest['case_id'] = 'tiny3'
        common.write_json(directory / 'case.json', manifest)
        request['manifest_sha256'] = common.sha(directory / 'case.json')
    class NoConnection:
        owned, facts = False, None
        def __init__(self, *a): pass
        def connect(self): pytest.fail('invalid binding must never connect')
        def close(self): pass
        def release(self): pass
    monkeypatch.setattr(worker, 'Session', NoConnection)
    common.write_json(directory / 'request.json', request)
    assert worker.main(directory / 'request.json') == 1
    result = common.read_json(directory / 'worker.json')
    assert result['status'] == 'FAIL' and result['first_error'] is not None


@pytest.mark.parametrize('fault', ['worker_control', 'worker_pid', 'excel_pid', 'cleanup', 'progress', 'primary',
                                 'missing_phase', 'flag', 'state', 'saved_sha', 'mutated_proof', 'cache'])
def test_parent_does_not_accept_mock_success_without_all_evidence(tmp_path, monkeypatch, fault):
    def monitor(directory, request):
        r = runtime(directory, request)
        if fault == 'worker_control': r['preopen_control']['verified'] = False
        if fault == 'worker_pid': r['python_process']['ppid'] = 999
        if fault == 'excel_pid': r['owned_pid'] = 999
        if fault == 'cleanup': r['cleanup_errors'] = [dict(message='Quit failed')]
        if fault == 'progress': r['progress_log_errors'] = [dict(message='log failed')]
        if fault == 'primary': r['first_error'] = dict(message='COM failed')
        if fault == 'missing_phase': r['functions24']['evidence'].pop()
        if fault == 'flag': r['functions24']['evidence'][0]['force_full'] = True
        if fault == 'state': r['samples'][-1]['state'] = 2
        if fault == 'saved_sha': r['saved_sha256'] = 'wrong'
        if fault == 'mutated_proof': (directory / 'transform.json').write_bytes(b'{}')
        if fault == 'cache':
            monkeypatch.setattr(run, 'verify_saved', lambda *a: (_ for _ in ()).throw(ValueError('no cache')))
        return r
    value, _ = execute(tmp_path, monkeypatch, monitor)
    assert value['status'] == 'FAIL' and value['exit_code'] == 1
    assert value['diagnostic_recalculation_verified'] is False
    assert value['stages'][0]['status'] == value['stages'][2]['status'] == 'NOT_RUN'


@pytest.mark.parametrize('fault', ['residual', 'residual_unknown', 'access', 'interrupted'])
def test_owned_process_exit_mandatory_and_no_force_termination(tmp_path, monkeypatch, fault):
    def residual(r):
        if fault == 'access': raise PermissionError('cannot observe exit')
        if fault == 'interrupted': raise KeyboardInterrupt()
        return None if fault == 'residual_unknown' else True
    value, _ = execute(tmp_path, monkeypatch, residual=residual)
    assert value['status'] in ('FAIL', 'INTERRUPTED')
    assert value['force_termination'] is False and value['diagnostic_recalculation_verified'] is False


@pytest.mark.parametrize('status', ['TIMEOUT', 'BLOCKED', 'INTERRUPTED'])
def test_original_failure_kept_with_secondary_cleanup_and_residual(tmp_path, monkeypatch, status):
    def monitor(directory, request):
        r = runtime(directory, request)
        r.update(status=status, first_error=dict(message='original failure'),
                 cleanup_errors=[dict(message='secondary Close failure')])
        return r
    value, _ = execute(tmp_path, monkeypatch, monitor, residual=lambda r: True)
    stage = value['stages'][1]
    assert value['status'] == status and stage['runtime']['first_error']['message'] == 'original failure'
    assert stage['runtime']['cleanup_errors'][0]['message'] == 'secondary Close failure'
    assert 'residual_error' in stage and value['diagnostic_recalculation_verified'] is False


def test_a_supervisor_seals_request_once_and_records_own_worker_only(tmp_path, monkeypatch):
    proof = python_proof()
    seen = []
    def observe(pid):
        seen.append(pid)
        return dict(pid=pid, ppid=0 if pid == os.getpid() else os.getpid(), created=1, path='MOCK')
    monkeypatch.setattr(f24, 'process_identity', observe)
    monkeypatch.setattr(f24, 'worker_exit_observation', lambda observed: dict(outcome='PROCESS_EXITED'))
    common.Progress(tmp_path, 'r', 'functions24').emit('python.worker_started',
                                                    process=dict(proof['python_process'], pid=os.getpid()))
    common.write_json(tmp_path / 'worker.json', dict(status='SAVED_REOPENED'))
    popens = []
    def popen(*a, **kw):
        popens.append(a)
        return SimpleNamespace(pid=2, returncode=0, poll=lambda: 0)
    request = dict(run_id='r', case_id='functions24', diagnostic=A_DIAGNOSTIC)
    result = run.supervise(tmp_path, request, popen=popen)
    assert result['status'] == 'SAVED_REOPENED' and seen == [os.getpid(), 2]
    assert result['python_processes']['worker_exit']['outcome'] == 'PROCESS_EXITED'
    with pytest.raises(FileExistsError): run.supervise(tmp_path, request, popen=popen)
    assert len(popens) == 1  # no retry worker when a request was already sealed


def test_cancel_after_full_rebuild_does_not_restore_save_or_continue(tmp_path):
    directory, manifest = prepared(tmp_path)
    s = ASession()
    original = s.excel.CalculateFullRebuild
    def stop():
        original(); (directory / 'STOP').touch()
    s.excel.CalculateFullRebuild = stop
    with pytest.raises(common.GateError) as caught: execute_worker(directory, manifest, s)
    assert caught.value.status == 'INTERRUPTED'
    assert s.rebuilds == 1 and ('edit', 100) not in s.calls and 'save' not in s.calls


@pytest.mark.parametrize('fault', ['preflight', 'approval', 'profile', 'units', 'transform'])
def test_precondition_failures_have_no_worker_launch(tmp_path, monkeypatch, fault):
    config = tmp_path / 'profile.json'
    if fault != 'profile': common.write_json(config, profile())
    if fault == 'transform':
        monkeypatch.setattr(builder, 'convert', lambda raw: raw)
    value, _ = run.execute('functions24-preopen-force0', config, tmp_path / 'run',
                           check=lambda: 'not Windows' if fault == 'preflight' else None,
                           monitor=lambda *a: pytest.fail('invalid preparation must not start worker'),
                           units=lambda r: dict(returncode=1 if fault == 'units' else 0),
                           approval_id=None if fault == 'approval' else APPROVAL)
    assert value['status'] in ('FAIL', 'BLOCKED')
    assert value['stages'][1]['status'] == value['status']


def test_real_cache_checker_rejects_uncalculated_a_copy(tmp_path):
    directory, manifest = prepared(tmp_path)
    (directory / 'results' / manifest['file']).write_bytes((directory / 'inputs' / manifest['file']).read_bytes())
    with pytest.raises((AssertionError, ValueError)): cases.verify_saved(directory, manifest)


def test_worker_keeps_primary_and_secondary_cleanup_errors(tmp_path, monkeypatch):
    directory, manifest = prepared(tmp_path, source=run.source_identity()['source_sha256'])
    calls = []
    class Session(ASession):
        def __init__(self, *a):
            super().__init__('stale_challenge'); self.owned, self.facts = False, None
        def connect(self):
            self.owned = True; self.facts = profile()['facts']; self.owner = self.facts
        def close(self): calls.append('quit'); raise RuntimeError('secondary Quit failure')
        def release(self): calls.append('release')
    monkeypatch.setattr(worker, 'Session', Session)
    monkeypatch.setattr(f24, 'worker_process', lambda: python_proof()['python_process'])
    request = dict(run_id='test', case_id='functions24', diagnostic=A_DIAGNOSTIC, protocol=A_PROTOCOL,
                   profile=profile(), source_sha256=run.source_identity()['source_sha256'],
                   manifest_sha256=common.sha(directory / 'case.json'), control=manifest['control'])
    common.write_json(directory / 'request.json', request)
    assert worker.main(directory / 'request.json') == 1
    result = common.read_json(directory / 'worker.json')
    assert result['first_error']['type'] == 'GateError'
    assert result['cleanup_errors'][0]['message'] == 'secondary Quit failure'
    assert calls == ['quit', 'release'] and not (directory / 'results' / manifest['file']).exists()


def test_cli_closed_a_mode_requires_id_and_keeps_false_acceptance(tmp_path, monkeypatch, capsys):
    value, root = execute(tmp_path, monkeypatch)
    seen = []
    def fake(*args, **kwargs): seen.append((args, kwargs)); return value, root
    monkeypatch.setattr(run, 'execute', fake)
    monkeypatch.setattr(run.sys, 'argv', ['run', 'functions24-preopen-force0', '--approval-id', APPROVAL])
    assert run.main() == 0
    assert seen[0][1] == dict(approval_id=APPROVAL)
    summary = json.loads(capsys.readouterr().out)
    assert summary['gate2_acceptance'] is False and summary['diagnostic_only'] is True
