"""E1 mocks only: no Windows/Excel, no fabricated XLSX formula caches."""
from pathlib import Path
from types import SimpleNamespace
import sys

import pytest

from tools.excel2024_validation import SCHEMA, E1_DIAGNOSTIC, E1_PROTOCOL
from tools.excel2024_validation import cases, common, e1, run, worker


def profile():
    facts = dict(pid=123, created=1000, path=r'C:\Office\EXCEL.EXE', version='16.0',
                 exe_sha256='MOCK-NOT-REAL-EXCEL', file_version='16.0.1.2', build='1',
                 installation=dict(product_release_ids=['HomeBusiness2024Retail']))
    return dict(schema=SCHEMA, product='Microsoft Office Home and Business 2024', licensed=True,
                facts=facts, ui_confirmed_on_owned_pid=True)


def runtime(directory, request):
    common.write_json(directory / 'request.json', request)
    manifest = common.read_json(directory / 'case.json')
    values = {s + '!' + c: v for s, c, v in manifest['probes']}
    samples = [dict(phase=p, state_before=0, state=0, mode=-4105,
                    differences=[], values=values.copy()) for p in ('before_save', 'reopened')]
    evidence = []
    for phase in ('opened', 'before_flag', 'after_flag', 'after_full_rebuild', 'after_automatic', 'reopened'):
        sample = dict(phase='e1.' + phase, state_before=0, state=0,
                      mode=-4135 if phase in ('before_flag', 'after_flag', 'after_full_rebuild') else -4105,
                      differences=[], values=values.copy())
        evidence.append(dict(phase=phase, utc=common.utc(), force_full=phase in ('opened', 'before_flag'),
                             workbooks=[dict(name=manifest['file'], path=str(directory / 'work' / manifest['file']),
                                             read_only=False)], sample=sample))
    # Opaque mock save, not an XLSX; real cache validation is tested separately.
    saved = directory / 'results' / manifest['file']
    saved.write_bytes(b'MOCK SAVE ONLY, NOT AN XLSX, NO FORMULA CACHES')
    return dict(schema=SCHEMA, run_id=request['run_id'], case_id='tiny3', status='SAVED_REOPENED',
                identity=profile()['facts'], owned_pid=123, first_error=None, cleanup_errors=[],
                progress_log_errors=[], worker_still_running=False, force_termination=False,
                request_sha256=common.sha(directory / 'request.json'), manifest_sha256=common.sha(directory / 'case.json'),
                saved_sha256=common.sha(saved), full_rebuild_calls=1, samples=samples,
                e1=dict(diagnostic=E1_DIAGNOSTIC, protocol=E1_PROTOCOL, evidence=evidence))


def execute(tmp_path, monkeypatch, monitor=runtime, residual=lambda r: False, units=lambda root: dict(returncode=0)):
    config = tmp_path / 'profile.json'
    common.write_json(config, profile())
    def fake_verify(directory, manifest):
        return dict(saved_sha256=common.sha(directory / 'results' / manifest['file']), numeric_comparison='MOCK-ONLY')
    monkeypatch.setattr(run, 'verify_saved', fake_verify)
    return run.execute('e1', config, tmp_path / 'run', check=lambda: None,
                       monitor=monitor, residual=residual, units=units)


def test_e1_only_tiny3_and_diagnostic_success_is_not_gate2(tmp_path, monkeypatch):
    requests = []
    def monitor(directory, request):
        requests.append(request)
        assert request['diagnostic'] == E1_DIAGNOSTIC
        return runtime(directory, request)
    value, root = execute(tmp_path, monkeypatch, monitor)
    assert [r['case_id'] for r in requests] == ['tiny3']
    assert value['status'] == 'DIAGNOSTIC_PASS' and value['exit_code'] == 0
    assert value['diagnostic_recalculation_verified'] is True
    assert value['diagnostic_only'] is True and value['gate2_acceptance'] is False
    assert value['product_acceptance'] is False and value['excel_recalculation_verified'] is False
    assert [s['status'] for s in value['stages']] == ['PASS', 'NOT_RUN', 'NOT_RUN']
    assert not (root / 'functions24').exists() and not (root / 'annual1').exists()
    original = cases.REPO / 'output/s1_tiny_diagnostic/tiny_calc.xlsx'
    assert common.sha(root / 'tiny3/inputs/tiny_calc.xlsx') == common.sha(original)
    assert common.sha(root / 'tiny3/work/tiny_calc.xlsx') == common.sha(original)
    assert common.read_json(root / 'run.json') == value


@pytest.mark.parametrize('problem', ['state', 'mode', 'value', 'cleanup', 'progress', 'primary', 'worker_alive',
                                     'request', 'manifest', 'rebuild', 'false_flag', 'missing_phase', 'phase_order',
                                     'snapshot_phase', 'manual_rebuild', 'missing_probe', 'protocol', 'diagnostic',
                                     'sku', 'license'])
def test_e1_rejects_incomplete_or_incorrect_proof(tmp_path, monkeypatch, problem):
    def monitor(directory, request):
        result = runtime(directory, request)
        if problem == 'state': result['samples'][0]['state'] = 2
        if problem == 'mode': result['samples'][1]['mode'] = -4135
        if problem == 'value': result['samples'][0]['values']['診断!B1'] = 9
        if problem == 'cleanup': result['cleanup_errors'] = [dict(message='Quit error')]
        if problem == 'progress': result['progress_log_errors'] = [dict(message='log error')]
        if problem == 'primary': result['first_error'] = dict(message='RPC error')
        if problem == 'worker_alive': result['worker_still_running'] = True
        if problem in ('request', 'manifest'): result[problem + '_sha256'] = 'wrong'
        if problem == 'rebuild': result['full_rebuild_calls'] = 0
        if problem == 'false_flag': result['e1']['evidence'][2]['force_full'] = True
        if problem == 'missing_phase': result['e1']['evidence'].pop()
        if problem == 'phase_order': result['e1']['evidence'].reverse()
        if problem == 'snapshot_phase': result['e1']['evidence'][2]['sample']['phase'] = 'other'
        if problem == 'manual_rebuild': result['e1']['evidence'][3]['sample']['mode'] = -4105
        if problem == 'missing_probe': result['e1']['evidence'][2]['sample']['values'].pop('診断!B1')
        if problem == 'protocol': result['e1']['protocol'] = 'E0'
        if problem == 'diagnostic': result['e1']['diagnostic'] = 'E2'
        if problem == 'sku': result['identity']['installation']['product_release_ids'] = ['O365ProPlusRetail']
        if problem == 'license': request['profile']['licensed'] = False
        return result
    value, _ = execute(tmp_path, monkeypatch, monitor)
    assert value['status'] in ('FAIL', 'BLOCKED') and value['exit_code'] != 0
    assert value['diagnostic_recalculation_verified'] is False
    assert [s['status'] for s in value['stages']][1:] == ['NOT_RUN', 'NOT_RUN']


@pytest.mark.parametrize('status', ['TIMEOUT', 'FAIL', 'BLOCKED', 'INTERRUPTED'])
def test_e1_worker_failure_keeps_original_error_and_checks_residual(tmp_path, monkeypatch, status):
    observed = []
    def monitor(directory, request):
        result = runtime(directory, request)
        result.update(status=status, first_error=dict(message='original calculation error'))
        result['cleanup_errors'] = [dict(message='secondary close error')]
        return result
    def residual(result):
        observed.append(result['identity']['pid'])
        return True
    value, _ = execute(tmp_path, monkeypatch, monitor, residual)
    stage = value['stages'][0]
    assert value['status'] == status
    assert observed == [123]
    assert stage['runtime']['first_error']['message'] == 'original calculation error'
    assert stage['runtime']['cleanup_errors'][0]['message'] == 'secondary close error'
    assert 'residual_error' in stage


@pytest.mark.parametrize('failure', ['remains', 'access', 'invalid', 'cache'])
def test_e1_never_succeeds_without_cache_and_process_exit(tmp_path, monkeypatch, failure):
    def residual(result):
        if failure == 'access': raise PermissionError('process lookup denied')
        return None if failure == 'invalid' else failure == 'remains'
    def monitor(directory, request):
        result = runtime(directory, request)
        if failure == 'cache':
            def reject(*args): raise ValueError('cached values mismatch')
            monkeypatch.setattr(run, 'verify_saved', reject)
        return result
    value, _ = execute(tmp_path, monkeypatch, monitor, residual)
    assert value['status'] == 'FAIL' and value['exit_code'] == 1


def test_real_verifier_rejects_uncalculated_input_for_e1(tmp_path):
    manifest = cases.prepare('tiny3', tmp_path, 'test')
    (tmp_path / 'results' / manifest['file']).write_bytes((tmp_path / 'inputs' / manifest['file']).read_bytes())
    with pytest.raises((ValueError, AssertionError)):
        cases.verify_saved(tmp_path, manifest)


@pytest.mark.parametrize('reason', ['preflight', 'profile', 'units'])
def test_e1_preconditions_prevent_com_launch(tmp_path, reason):
    config = tmp_path / 'profile.json'
    if reason != 'profile': common.write_json(config, profile())
    def forbidden(*args): pytest.fail('must not launch Excel')
    value, _ = run.execute('e1', config, tmp_path / 'run', check=lambda: 'no Windows' if reason == 'preflight' else None,
                           monitor=forbidden, units=lambda root: dict(returncode=1))
    assert value['status'] in ('FAIL', 'BLOCKED')
    assert value['gate2_acceptance'] is False


class FakeSession:
    """In-memory protocol mock; writes opaque non-XLSX saves only."""
    def __init__(self, *, pending=False, refuse_flag=False, close_error=False):
        self.calls = []
        self.values = dict(A1=2, A2=3, B1=5, B2=6, B3=11)
        self.current = None
        self.saved_flag = False
        self.pending = pending
        self.refuse_flag = refuse_flag
        self.close_error = close_error
        owner = self
        class Cell:
            def __init__(self, cell): self.cell = cell
            @property
            def Value2(self): return owner.values[self.cell]
            @Value2.setter
            def Value2(self, value): owner.values[self.cell] = value
        class Book:
            Date1904 = False
            def __init__(self, path, readonly):
                self.FullName, self.Name, self.ReadOnly = path, Path(path).name, readonly
                self.flag = owner.saved_flag if readonly else True
                self.Worksheets = SimpleNamespace(Item=lambda sheet: SimpleNamespace(Range=Cell))
            @property
            def ForceFullCalculation(self): return self.flag
            @ForceFullCalculation.setter
            def ForceFullCalculation(self, value):
                owner.calls.append(('force', value))
                if not owner.refuse_flag: self.flag = value
            def SaveAs(self, path, FileFormat):
                assert FileFormat == 51
                owner.calls.append('save')
                owner.saved_flag = self.flag
                Path(path).write_bytes(b'MOCK SAVE ONLY, NOT XLSX, NO CACHES')
            def Close(self, save):
                assert save is False
                owner.calls.append('close')
                if owner.close_error: raise RuntimeError('secondary close error')
                owner.current = None
        class BookCollection:
            @property
            def Count(self): return int(owner.current is not None)
            def Item(self, index):
                assert index == 1
                return owner.current
            def Open(self, path, **kwargs):
                owner.calls.append(('open', kwargs['ReadOnly']))
                owner.current = Book(path, kwargs['ReadOnly'])
                return owner.current
        class Excel:
            CalculationState = 2 if pending else 0
            Workbooks = BookCollection()
            def __init__(self): self.mode = -4105
            @property
            def Calculation(self): return self.mode
            @Calculation.setter
            def Calculation(self, value):
                owner.calls.append(('mode', value))
                self.mode = value
            def CalculateFullRebuild(self):
                owner.calls.append('full_rebuild')
                a, b = owner.values['A1'], owner.values['A2']
                owner.values.update(B1=a+b, B2=a*b, B3=a+b+a*b)
        self.excel = Excel()
    def confirm_owner(self): self.calls.append('identity')


def test_e1_worker_uses_false_but_keeps_e0_calculation_order_and_saves(tmp_path):
    manifest = cases.prepare('tiny3', tmp_path, 'test')
    s = FakeSession()
    evidence = []
    result = worker.execute_case(s, manifest, tmp_path, common.Progress(tmp_path, 'test', 'tiny3'), [],
                                 diagnostic=E1_DIAGNOSTIC, evidence=evidence)
    assert result['status'] == 'SAVED_REOPENED' and result['full_rebuild_calls'] == 1
    protocol = [c for c in s.calls if c == 'full_rebuild' or isinstance(c, tuple) and c[0] in ('mode', 'force')]
    assert protocol == [('mode', -4135), ('force', False), 'full_rebuild', ('mode', -4105)]
    assert s.calls.count('save') == 1 and s.calls.count('close') == 2
    assert [e['phase'] for e in evidence] == ['opened', 'before_flag', 'after_flag', 'after_full_rebuild', 'after_automatic', 'reopened']
    e1.validate_evidence(dict(e1=dict(diagnostic=E1_DIAGNOSTIC, protocol=E1_PROTOCOL, evidence=evidence)), manifest)
    assert evidence[0]['force_full'] is True and evidence[2]['force_full'] is False


def test_e1_false_readback_rejection_stops_before_rebuild(tmp_path):
    manifest = cases.prepare('tiny3', tmp_path, 'test')
    s = FakeSession(refuse_flag=True)
    with pytest.raises(common.GateError, match='読み戻せません'):
        worker.execute_case(s, manifest, tmp_path, common.Progress(tmp_path, 'test', 'tiny3'), [],
                            diagnostic=E1_DIAGNOSTIC, evidence=[])
    assert 'full_rebuild' not in s.calls and 'save' not in s.calls


def test_e1_correct_values_with_pending_never_save(tmp_path, monkeypatch):
    manifest = cases.prepare('tiny3', tmp_path, 'test')
    s = FakeSession(pending=True)
    now = [0]
    monkeypatch.setattr(worker.time, 'monotonic', lambda: now[0])
    monkeypatch.setattr(worker.time, 'sleep', lambda seconds: now.__setitem__(0, now[0] + 60))
    monkeypatch.setattr(worker, 'pump_messages', lambda: None)
    with pytest.raises(common.GateError) as error:
        worker.execute_case(s, manifest, tmp_path, common.Progress(tmp_path, 'test', 'tiny3'), [],
                            diagnostic=E1_DIAGNOSTIC, evidence=[])
    assert error.value.status == 'TIMEOUT' and now[0] == 180
    assert [s.values[c] for c in ('B1', 'B2', 'B3')] == [10, 21, 31]
    assert 'save' not in s.calls


@pytest.mark.parametrize('case_id,diagnostic', [('annual1', E1_DIAGNOSTIC), ('functions24', E1_DIAGNOSTIC),
                                               ('doctor', E1_DIAGNOSTIC), ('tiny3', 'E2')])
def test_worker_rejects_non_tiny3_or_unapproved_diagnostic_before_connect(tmp_path, monkeypatch, case_id, diagnostic):
    class Session:
        owned = False
        facts = None
        def __init__(self, *args): pass
        def connect(self): pytest.fail('invalid request must not start Excel')
        def close(self): pass
        def release(self): pass
    monkeypatch.setattr(worker, 'Session', Session)
    request = dict(run_id='test', case_id=case_id, diagnostic=diagnostic, source_sha256=run.source_identity()['source_sha256'])
    common.write_json(tmp_path / 'request.json', request)
    assert worker.main(tmp_path / 'request.json') == 1
    assert '未承認' in common.read_json(tmp_path / 'worker.json')['first_error']['message']


def test_e1_cli_exposes_diagnostic_result(tmp_path, monkeypatch, capsys):
    value, root = execute(tmp_path, monkeypatch)
    monkeypatch.setattr(run, 'execute', lambda *args: (value, root))
    monkeypatch.setattr(sys, 'argv', ['run', 'e1'])
    assert run.main() == 0
    import json
    summary = json.loads(capsys.readouterr().out)
    assert summary['status'] == 'DIAGNOSTIC_PASS' and summary['gate2_acceptance'] is False
    assert summary['diagnostic_only'] is True


@pytest.mark.parametrize('flag', [None, '', 0, []])
def test_e1_unknown_flag_type_is_not_coerced_to_false(tmp_path, flag):
    log = common.Progress(tmp_path, 'test', 'tiny3')
    with pytest.raises(common.GateError, match='真偽値ではありません'):
        worker.e1_context(None, SimpleNamespace(ForceFullCalculation=flag), None, log, 'after_flag')


def test_e1_flag_com_exception_keeps_hresult(tmp_path):
    class Rejected(RuntimeError): hresult = -2147418111
    class Book:
        @property
        def ForceFullCalculation(self): raise Rejected('RPC rejected')
    log = common.Progress(tmp_path, 'test', 'tiny3')
    with pytest.raises(Rejected):
        worker.e1_context(None, Book(), None, log, 'before_flag')
    assert run.events(log.path)[-1]['error']['hresult'] == -2147418111


def test_e1_unexpected_owned_workbook_is_blocked_without_closing_it(tmp_path):
    manifest = cases.prepare('tiny3', tmp_path, 'test')
    s = FakeSession()
    open_book = s.excel.Workbooks.Open
    s.excel.Workbooks = SimpleNamespace(Count=2, Item=lambda index: s.current, Open=open_book)
    with pytest.raises(common.GateError) as error:
        worker.execute_case(s, manifest, tmp_path, common.Progress(tmp_path, 'test', 'tiny3'), [],
                            diagnostic=E1_DIAGNOSTIC, evidence=[])
    assert error.value.status == 'BLOCKED'
    assert 'full_rebuild' not in s.calls and 'close' not in s.calls


def test_e1_worker_preserves_timeout_and_secondary_cleanup_errors(tmp_path, monkeypatch):
    manifest = cases.prepare('tiny3', tmp_path, 'test')
    calls = []
    class Session(FakeSession):
        def __init__(self, *args):
            super().__init__(pending=True, close_error=True)
            self.owned = False
            self.facts = None
        def connect(self):
            self.owned = True
            self.facts = profile()['facts']
            self.owner = self.facts
        def close(self): calls.append('quit'); raise RuntimeError('secondary Quit failure')
        def release(self): calls.append('release')
    monkeypatch.setattr(worker, 'Session', Session)
    now = [0]
    monkeypatch.setattr(worker.time, 'monotonic', lambda: now[0])
    monkeypatch.setattr(worker.time, 'sleep', lambda seconds: now.__setitem__(0, now[0] + 60))
    monkeypatch.setattr(worker, 'pump_messages', lambda: None)
    request = dict(run_id='test', case_id='tiny3', diagnostic=E1_DIAGNOSTIC, profile=profile(),
                   source_sha256=run.source_identity()['source_sha256'],
                   manifest_sha256=common.sha(tmp_path / 'case.json'))
    common.write_json(tmp_path / 'request.json', request)
    assert worker.main(tmp_path / 'request.json') == 1
    result = common.read_json(tmp_path / 'worker.json')
    assert result['status'] == 'TIMEOUT' and result['first_error']['type'] == 'GateError'
    assert [e['message'] for e in result['cleanup_errors']] == ['secondary close error', 'secondary Quit failure']
    assert result['e1']['diagnostic'] == E1_DIAGNOSTIC and len(result['e1']['evidence']) == 5
    assert calls == ['quit', 'release']
    assert not (tmp_path / 'results' / manifest['file']).exists()


def test_normal_worker_still_uses_true_without_e1_observation_calls(tmp_path):
    manifest = cases.prepare('tiny3', tmp_path, 'test')
    s = FakeSession()
    log = common.Progress(tmp_path, 'test', 'tiny3')
    result = worker.execute_case(s, manifest, tmp_path, log, [])
    assert result['status'] == 'SAVED_REOPENED'
    protocol = [c for c in s.calls if c == 'full_rebuild' or isinstance(c, tuple) and c[0] in ('mode', 'force')]
    assert protocol == [('mode', -4135), ('force', True), 'full_rebuild', ('mode', -4105)]
    assert not any(e['event'].startswith('e1.') for e in run.events(log.path))


def test_normal_smoke_keeps_all_three_stages_and_no_e1_request(tmp_path, monkeypatch):
    config = tmp_path / 'profile.json'
    common.write_json(config, profile())
    launched = []
    def prepare(stage, directory, run_id):
        common.write_json(directory / 'case.json', dict(case_id=stage))
        return dict(case_id=stage)
    def monitor(directory, request):
        assert 'diagnostic' not in request
        launched.append(request['case_id'])
        return dict(status='SAVED_REOPENED', saved_sha256='MOCK-ONLY')
    monkeypatch.setattr(run, 'prepare', prepare)
    monkeypatch.setattr(run, 'validate_runtime', lambda *args: None)
    monkeypatch.setattr(run, 'verify_saved', lambda *args: dict(saved_sha256='MOCK-ONLY'))
    value, _ = run.execute('smoke', config, tmp_path / 'run', check=lambda: None, monitor=monitor,
                           units=lambda root: dict(returncode=0), residual=lambda result: False)
    assert launched == ['tiny3', 'functions24', 'annual1']
    assert value['status'] == 'PASS' and value['excel_recalculation_verified'] is True
    assert [s['status'] for s in value['stages']] == ['PASS'] * 3
    assert 'diagnostic_only' not in value and 'gate2_acceptance' not in value
