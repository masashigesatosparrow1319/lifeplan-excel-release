"""Protocol tests use mocks only. No Excel/cached results are fabricated."""
from types import SimpleNamespace

import pytest

from tools.excel2024_validation import SCHEMA, STAGES
from tools.excel2024_validation import cases, common, run, session, worker


def facts():
    return dict(pid=123, created=1000, path=r'C:\Office\EXCEL.EXE', version='16.0',
                exe_sha256='EXAMPLE-ONLY', file_version='16.0.1.2', build='1',
                installation=dict(product_release_ids=['Excel2024Retail']))


def profile():
    return dict(schema=SCHEMA, product='Microsoft Excel 2024', licensed=True,
                facts=facts(), ui_confirmed_on_owned_pid=True)


@pytest.mark.parametrize('product,licensed,ids', [
    ('Microsoft 365', True, ['O365ProPlusRetail']),
    ('Microsoft Excel 2021', True, ['Excel2021Retail']),
    ('Microsoft Excel 2024', False, ['Excel2024Retail']),
    ('unknown 2024', True, ['Excel2024Retail']),
    ('Microsoft Excel 2024', True, []),
    ('Microsoft Excel 2024', True, ['Excel2024Retail', 'O365ProPlusRetail']),
    ('Microsoft Excel 2024', True, ['Excel2024Retail', 'UNKNOWN']),
    ('Microsoft Excel 2024 / Microsoft 365', True, ['Excel2024Retail']),
])
def test_ui_and_installation_both_required(product, licensed, ids):
    with pytest.raises(common.GateError, match='BLOCKED_IDENTITY'):
        session.check_product(product, licensed, dict(product_release_ids=ids))


@pytest.mark.parametrize('product,ids', [
    ('Microsoft Excel 2024', ['Excel2024Retail']),
    ('Microsoft Office Home 2024', ['Home2024Retail']),
    ('Microsoft Office LTSC Professional Plus 2024', ['ProPlus2024Volume']),
])
def test_explicit_2024_ui_and_unambiguous_metadata(product, ids):
    session.check_product(product, True, dict(product_release_ids=ids))


@pytest.mark.parametrize('field,value', [('pid', 5), ('created', 900), ('path', r'C:\other\EXCEL.EXE')])
def test_existing_pid_wrong_path_or_old_creation_never_owned(field, value):
    f = facts(); f[field] = value
    with pytest.raises(common.GateError):
        session.check_ownership(f, {5}, 1000, r'C:\Office\EXCEL.EXE')


def test_case_insensitive_windows_path_ownership():
    session.check_ownership(facts(), {5}, 1000, r'c:\office\excel.exe')


@pytest.mark.parametrize('field,value', [('version', '14.0'), ('exe_sha256', 'new binary'),
                                       ('build', 'updated'), ('file_version', 'new version'),
                                       ('path', r'C:\other\EXCEL.EXE')])
def test_profile_invalidated_when_process_environment_changes(field, value):
    f = facts(); f[field] = value
    with pytest.raises(common.GateError):
        session.check_profile(profile(), f)


def test_same_version_16_cannot_replace_ui_evidence():
    p = profile(); p['ui_confirmed_on_owned_pid'] = False
    with pytest.raises(common.GateError):
        session.check_profile(p, facts())


@pytest.fixture(scope='module')
def prepared(tmp_path_factory):
    root = tmp_path_factory.mktemp('gate2-prepared')
    result = {}
    for stage in STAGES:
        directory = root / stage; directory.mkdir()
        result[stage] = (directory, cases.prepare(stage, directory, 'test-run'))
    return result


def test_only_three_scoped_inputs_and_no_formula_caches(prepared):
    assert set(prepared) == set(STAGES)
    assert prepared['tiny3'][1]['formula_count'] == 3
    assert prepared['functions24'][1]['formula_count'] == 24
    annual, manifest = prepared['annual1']
    packet = common.read_json(annual / 'inputs/manifest.json')
    assert len(packet['entries']) == 1 and packet['entries'][0]['file'] == 'slot200_add.xlsx'
    assert (packet['entries'][0]['expected_cash'], packet['entries'][0]['expected_total']) == (1100000, 1600000)
    for directory, manifest in prepared.values():
        assert cases.static_input(directory / 'inputs' / manifest['file']) > 0
        assert common.sha(directory / 'inputs' / manifest['file']) == manifest['sha256']


def test_24_original_expectations_unchanged_and_update_challenge_separate(prepared):
    directory, manifest = prepared['functions24']
    assert common.sha(directory / 'inputs' / manifest['file']) == common.sha(cases.REPO / 'output/s1_compatibility/01_function_gate.xlsx')
    expected = {cell: amount for _, cell, amount in manifest['probes']}
    assert len(expected) == 24
    assert [expected[f'I{r}'] for r in range(8, 14)] == [0, 1, 0, 1, 0, 1]
    assert [expected[f'L{r}'] for r in range(8, 14)] == [2, 2, 0, 2, 1, 4]
    assert manifest['challenges'] == [['確認', 'B8', 0]]
    assert manifest['edits'] == [['確認', 'B8', 100]]


@pytest.mark.parametrize('stage', STAGES)
def test_uncalculated_original_is_never_a_saved_pass(prepared, stage):
    directory, manifest = prepared[stage]
    import shutil
    target = directory / 'results' / manifest['file']
    shutil.copyfile(directory / 'inputs' / manifest['file'], target)
    try:
        with pytest.raises((ValueError, AssertionError)):
            cases.verify_saved(directory, manifest)
    finally:
        target.unlink()


def test_zero_expected_value_and_missing_value_are_different():
    manifest = dict(probes=[['s', 'A1', 0]])
    assert cases.compare_probes(manifest, lambda s, c: 0) == []
    assert cases.compare_probes(manifest, lambda s, c: None)[0]['actual'] is None
    assert cases.compare_probes(manifest, lambda s, c: False)


def test_run_directory_cannot_be_repository_or_existing_original(tmp_path):
    with pytest.raises(ValueError):
        common.create_run(cases.REPO / 'output/forbidden')
    with pytest.raises(FileExistsError):
        common.create_run(tmp_path)
    path = tmp_path / 'new'
    assert common.create_run(path) == path.resolve()


def test_call_logs_before_block_and_late_return_cannot_continue(tmp_path):
    log = common.Progress(tmp_path, 'test', 'tiny3')
    def action():
        log.stop.touch()
        return 'late value'
    with pytest.raises(common.GateError, match='停止要求'):
        log.call('save', action)
    records = run.events(log.path)
    assert [e['event'] for e in records] == ['save.before_call', 'save.after_call']
    called = []
    with pytest.raises(common.GateError):
        log.call('must_not_save', lambda: called.append(True))
    assert not called
    log.call('cleanup', lambda: called.append('normal cleanup only'), cleanup=True)
    assert called == ['normal cleanup only']


def test_partial_json_line_does_not_hide_prior_flushed_progress(tmp_path):
    path = tmp_path / 'progress.jsonl'
    path.write_text('{"event":"one"}\n{"event":', encoding='utf-8')
    assert run.events(path) == [dict(event='one')]


@pytest.mark.parametrize('interrupt', [False, True])
def test_supervisor_deadline_or_ctrl_c_stops_without_termination(tmp_path, interrupt):
    class Process:
        pid = 99
        returncode = None
        def poll(self): return None
        def kill(self): pytest.fail('強制終了禁止')
        def terminate(self): pytest.fail('強制終了禁止')
    now = [0.0]
    def pause(n):
        if interrupt:
            raise KeyboardInterrupt()
        now[0] += 601
    request = dict(run_id='test', case_id='tiny3')
    result = run.supervise(tmp_path, request, popen=lambda *a, **k: Process(),
                           clock=lambda: now[0], pause=pause)
    assert result['status'] == ('INTERRUPTED' if interrupt else 'TIMEOUT')
    assert result['worker_still_running'] is True
    assert result['force_termination'] is False and (tmp_path / 'STOP').exists()


def test_sync_call_timeout_reports_its_exact_stage(tmp_path):
    log = common.Progress(tmp_path, 'test', 'tiny3')
    log.emit('final.full_rebuild.before_call', timeout_seconds=0)
    class Process:
        pid = 99; returncode = None
        def poll(self): return None
    result = run.supervise(tmp_path, dict(run_id='test', case_id='tiny3'), popen=lambda *a, **k: Process())
    assert result['status'] == 'TIMEOUT'
    assert result['pending_calls'] == ['final.full_rebuild']


def test_supervisor_missing_worker_log_is_failure(tmp_path):
    class Process:
        pid = 99; returncode = 0
        def poll(self): return 0
    result = run.supervise(tmp_path, dict(run_id='test', case_id='tiny3'), popen=lambda *a, **k: Process())
    assert result['status'] == 'FAIL' and '最終ログなし' in result['reason']


def test_linux_preflight_blocks_first_stage_and_never_launches(tmp_path):
    def forbidden(*a): pytest.fail('Excel不在時は生成/起動しない')
    result, root = run.execute('smoke', tmp_path / 'profile.json', tmp_path / 'run',
                               check=lambda: 'BLOCKED_EXCEL_UNAVAILABLE', monitor=forbidden, units=forbidden)
    assert result['status'] == 'BLOCKED' and result['excel_recalculation_verified'] is False
    assert [s['status'] for s in result['stages']] == ['BLOCKED', 'NOT_RUN', 'NOT_RUN']
    assert common.read_json(root / 'run.json')['status'] == 'BLOCKED'


@pytest.mark.parametrize('failed_stage', STAGES)
def test_failure_never_starts_the_next_stage(tmp_path, monkeypatch, failed_stage):
    config = tmp_path / 'profile.json'; common.write_json(config, profile())
    launched = []
    def prepare(stage, directory, run_id):
        common.write_json(directory / 'case.json', dict(case_id=stage))
        return dict(case_id=stage)
    def monitor(directory, request):
        launched.append(request['case_id'])
        return dict(status='FAIL' if request['case_id'] == failed_stage else 'SAVED_REOPENED', saved_sha256='test')
    monkeypatch.setattr(run, 'prepare', prepare)
    monkeypatch.setattr(run, 'validate_runtime', lambda *a: None)
    monkeypatch.setattr(run, 'verify_saved', lambda *a: dict(saved_sha256='test'))
    result, _ = run.execute('smoke', config, tmp_path / 'run', check=lambda: None, monitor=monitor,
                            units=lambda root: dict(returncode=0), residual=lambda result: False)
    index = STAGES.index(failed_stage)
    assert launched == list(STAGES[:index + 1])
    assert [s['status'] for s in result['stages']] == ['PASS'] * index + ['FAIL'] + ['NOT_RUN'] * (2 - index)
    assert result['status'] == 'FAIL' and result['excel_recalculation_verified'] is False


def test_python_failure_never_opens_excel(tmp_path):
    config = tmp_path / 'profile.json'; common.write_json(config, profile())
    def forbidden(*a): pytest.fail('単体試験不合格ならExcelへ進まない')
    result, _ = run.execute('smoke', config, tmp_path / 'run', check=lambda: None, monitor=forbidden,
                            units=lambda root: dict(returncode=1))
    assert result['status'] == 'FAIL'
    assert all(s['status'] == 'NOT_RUN' for s in result['stages'])


@pytest.mark.parametrize('state,mode,value', [(2, -4105, 10), (0, -4135, 10), (0, -4105, 5)])
def test_correct_values_pending_and_done_old_values_both_fail(tmp_path, state, mode, value):
    excel = SimpleNamespace(CalculationState=state, Calculation=mode)
    book = SimpleNamespace(Worksheets=SimpleNamespace(Item=lambda sheet: SimpleNamespace(Range=lambda c: SimpleNamespace(Value2=value))))
    manifest = dict(probes=[['s', 'B1', 10]])
    with pytest.raises(common.GateError):
        worker.snapshot(excel, book, manifest, common.Progress(tmp_path, 'test', 'tiny3'), 'before_save')


def test_full_rebuild_automatic_restore_then_state_wait_in_that_order(tmp_path, monkeypatch):
    calls = []
    class Excel:
        def CalculateFullRebuild(self): calls.append('full_rebuild')
        def __setattr__(self, name, value):
            if name == 'Calculation': calls.append(('mode', value))
            object.__setattr__(self, name, value)
    excel = Excel()
    monkeypatch.setattr(worker, 'await_done', lambda e, log: calls.append('wait_done'))
    worker.recalculate(excel, SimpleNamespace(), common.Progress(tmp_path, 'test', 'tiny3'), 'final')
    assert calls == ['full_rebuild', ('mode', -4105), 'wait_done']


def test_pending_is_observed_and_message_pumped_before_done(tmp_path, monkeypatch):
    states = iter((2, 0))
    class Excel:
        @property
        def CalculationState(self): return next(states)
    pumped = []
    monkeypatch.setattr(worker, 'pump_messages', lambda: pumped.append(True))
    monkeypatch.setattr(worker.time, 'sleep', lambda n: None)
    log = common.Progress(tmp_path, 'test', 'tiny3')
    worker.await_done(Excel(), log)
    assert pumped == [True]
    assert [e['state'] for e in run.events(log.path) if e['event'] == 'calculation.observed'] == [2, 0]


def test_stopped_worker_cleans_up_owned_com_without_hiding_primary_error(tmp_path, monkeypatch):
    called = []
    class Session:
        def __init__(self, *args): self.owned = False; self.facts = None
        def connect(self):
            self.owned = True; self.facts = facts()
            raise RuntimeError('primary RPC_E_DISCONNECTED')
        def close(self): called.append('normal quit'); raise RuntimeError('secondary Quit error')
        def release(self): called.append('release')
    monkeypatch.setattr(worker, 'Session', Session)
    common.write_json(tmp_path / 'request.json', dict(run_id='test', case_id='tiny3', source_sha256=run.source_identity()['source_sha256']))
    assert worker.main(tmp_path / 'request.json') == 1
    result = common.read_json(tmp_path / 'worker.json')
    assert 'primary RPC_E_DISCONNECTED' in result['first_error']['message']
    assert result['cleanup_errors'][0]['message'] == 'secondary Quit error'
    assert called == ['normal quit', 'release']


def test_unowned_session_cannot_quit_even_when_application_exists(tmp_path):
    s = session.Session(common.Progress(tmp_path, 'test', 'tiny3'))
    s.excel = SimpleNamespace(Quit=lambda: pytest.fail('所有不明のExcelを終了してはいけません'))
    s.close()


@pytest.mark.parametrize('problem', ['state', 'mode', 'values', 'live', 'cleanup', 'request', 'manifest', 'count', 'run', 'identity'])
def test_runtime_proof_cannot_be_replaced_by_numeric_cache_only(tmp_path, problem):
    request = dict(run_id='test', case_id='tiny3')
    common.write_json(tmp_path / 'request.json', request)
    common.write_json(tmp_path / 'case.json', dict(case_id='tiny3', probes=[['s', 'B1', 10]]))
    sample = dict(state=0, state_before=0, mode=-4105, differences=[], values={'s!B1': 10})
    result = dict(schema=SCHEMA, run_id='test', case_id='tiny3', status='SAVED_REOPENED', identity=facts(),
                  request_sha256=common.sha(tmp_path / 'request.json'), manifest_sha256=common.sha(tmp_path / 'case.json'),
                  full_rebuild_calls=1, first_error=None, cleanup_errors=[],
                  samples=[dict(phase=p, **sample) for p in ('before_save', 'reopened')])
    if problem in ('state', 'mode'): result['samples'][1][problem] = 2
    if problem == 'values': result['samples'][1]['differences'] = [dict(actual=None)]
    if problem == 'live': result['samples'][1]['values'] = {'s!B1': 5}  # no trusted 'matched' flag
    if problem == 'cleanup': result['cleanup_errors'] = ['error']
    if problem == 'request': result['request_sha256'] = 'other'
    if problem == 'manifest': result['manifest_sha256'] = 'other'
    if problem == 'count': result['full_rebuild_calls'] = 0
    if problem == 'run': result['run_id'] = 'old'
    if problem == 'identity': result['identity']['installation'] = dict(product_release_ids=['O365ProPlusRetail'])
    with pytest.raises((ValueError, common.GateError)):
        run.validate_runtime(result, tmp_path, request, profile())


def test_doctor_never_overwrites_environment_profile(tmp_path):
    config = tmp_path / 'profile.json'; config.write_text('PRESERVE', encoding='utf-8')
    result, _ = run.execute('doctor', config, tmp_path / 'run', check=lambda: None)
    assert result['status'] == 'FAIL' and config.read_text(encoding='utf-8') == 'PRESERVE'


def test_live_edit_keeps_zero_and_uses_ole_date_not_numberformat():
    targets = {}
    def cell(address): return targets.setdefault(address, SimpleNamespace(Value2=None))
    book = SimpleNamespace(Date1904=False, Worksheets=SimpleNamespace(Item=lambda sheet: SimpleNamespace(Range=cell)))
    worker.set_edits(book, [['s', 'A1', 0], ['s', 'A2', {'date': '2026-08-01'}]])
    assert targets['A1'].Value2 == 0 and targets['A2'].Value2 == 46235


@pytest.mark.parametrize('stage', ('tiny3', 'functions24'))
def test_mock_com_full_protocol_updates_restores_saves_and_reopens(prepared, stage):
    # A stub, NOT Excel and NOT a cache producer. Its saved file is opaque text
    # and is never passed to the real saved-XLSX verifier or reported as real PASS.
    directory, manifest = prepared[stage]
    values = {(s, c): v for s, c, v in manifest['probes']}
    values[('確認', 'B8')] = 100
    calls = []
    class Cell:
        def __init__(self, s, c): self.key = (s, c)
        @property
        def Value2(self): return values[self.key]
        @Value2.setter
        def Value2(self, value): values[self.key] = value; calls.append(('edit', self.key, value))
    class Book:
        ReadOnly = False
        Date1904 = False
        Worksheets = SimpleNamespace(Item=lambda s: SimpleNamespace(Range=lambda c: Cell(s, c)))
        def SaveAs(self, path, FileFormat):
            calls.append('save'); assert FileFormat == 51
            from pathlib import Path
            Path(path).write_bytes(b'MOCK COM SAVE, NOT AN XLSX, NO CACHES')
        def Close(self, save): calls.append('close'); assert save is False
    class Excel:
        Calculation = -4105
        CalculationState = 0
        Workbooks = SimpleNamespace(Open=lambda path, **kwargs: opened(kwargs))
        def CalculateFullRebuild(self):
            calls.append('full_rebuild')
            if stage == 'tiny3':
                a, b = values[('診断', 'A1')], values[('診断', 'A2')]
                values.update({('診断', 'B1'): a+b, ('診断', 'B2'): a*b, ('診断', 'B3'): a+b+a*b})
            else:
                challenge = values[('確認', 'B8')] == 0
                values[('確認', 'L8')] = 1 if challenge else 2
                values[('確認', 'M8')] = '不一致' if challenge else '一致'
    def opened(kwargs):
        calls.append(('open', kwargs['ReadOnly']))
        assert kwargs['UpdateLinks'] == 0 and kwargs['CorruptLoad'] == 0
        return Book()
    if stage == 'tiny3':
        values.update({('診断', 'A1'): 2, ('診断', 'A2'): 3,
                       ('診断', 'B1'): 5, ('診断', 'B2'): 6, ('診断', 'B3'): 11})
    fake = SimpleNamespace(excel=Excel(), confirm_owner=lambda: calls.append('identity'))
    log = common.Progress(directory, 'test-run', stage)
    target = directory / 'results' / manifest['file']
    try:
        result = worker.execute_case(fake, manifest, directory, log, [])
        assert result['status'] == 'SAVED_REOPENED'
        assert calls.count('full_rebuild') == (2 if stage == 'functions24' else 1)
        assert calls.count('save') == 1 and calls.count('close') == 2
        assert [c for c in calls if isinstance(c, tuple) and c[0] == 'open'] == [('open', False), ('open', True)]
        assert all(not s['differences'] and s['state'] == 0 for s in result['samples'][1:])
        if stage == 'functions24':
            challenge = next(s for s in result['samples'] if s['phase'] == 'challenge_updated')
            assert challenge['values']['確認!L8'] == 1
            assert challenge['values']['確認!M8'] == '不一致'
            assert result['samples'][-1]['values']['確認!L8'] == 2
            assert values[('確認', 'B8')] == 100
        else:
            assert list(result['samples'][-1]['values'].values()) == [10, 21, 31]
    finally:
        target.unlink(missing_ok=True)


def test_supervisor_launch_error_retains_failure_without_killing(tmp_path):
    def failing(*args, **kwargs): raise OSError('worker launch failure')
    result = run.supervise(tmp_path, dict(run_id='test', case_id='tiny3'), popen=failing)
    assert result['status'] == 'FAIL' and result['worker_pid'] is None
    assert (tmp_path / 'STOP').exists()


def test_ui_wait_does_not_block_supervisors_deadline(tmp_path):
    import threading
    log = common.Progress(tmp_path, 'test', 'doctor')
    log.emit('doctor.ui_request', identity=facts())
    log.emit('doctor.ui_wait.before_call', timeout_seconds=0)
    release = threading.Event()
    class Process:
        pid = 99; returncode = None
        def poll(self): return None
    def blocked_ui(event):
        release.wait(2)
        return dict(pid=123)
    try:
        result = run.supervise(tmp_path, dict(run_id='test', case_id='doctor'),
                               popen=lambda *a, **k: Process(), ui=blocked_ui)
        assert result['status'] == 'TIMEOUT' and not (tmp_path / 'ui.json').exists()
    finally:
        release.set()


def test_missing_initial_product_profile_is_blocked_before_unit_or_com(tmp_path):
    def forbidden(*a): pytest.fail('環境未確認なら先へ進めない')
    result, _ = run.execute('smoke', tmp_path / 'missing.json', tmp_path / 'run', check=lambda: None,
                            monitor=forbidden, units=forbidden)
    assert result['status'] == 'BLOCKED'
    assert [s['status'] for s in result['stages']] == ['BLOCKED', 'NOT_RUN', 'NOT_RUN']


def test_doctor_profile_is_only_configuration_not_a_recalculation_pass(tmp_path):
    def monitor(directory, request):
        common.write_json(directory / 'request.json', request)
        return dict(schema=SCHEMA, run_id=request['run_id'], case_id='doctor', status='CONFIGURED',
                    profile=profile(), identity=facts(), first_error=None, cleanup_errors=[], progress_log_errors=[],
                    request_sha256=common.sha(directory / 'request.json'))
    config = tmp_path / 'profile.json'
    result, _ = run.execute('doctor', config, tmp_path / 'run', check=lambda: None, monitor=monitor, residual=lambda r: False)
    assert result['status'] == 'CONFIGURED' and result['excel_recalculation_verified'] is False
    assert all(s['status'] == 'NOT_RUN' for s in result['stages'])
    assert common.read_json(config)['ui_confirmed_on_owned_pid'] is True


def test_recycled_pid_cannot_be_quit(tmp_path):
    s = session.Session(common.Progress(tmp_path, 'test', 'tiny3'))
    s.owned = True; s.owner = facts()
    s.excel = SimpleNamespace(Hwnd=9, Quit=lambda: pytest.fail('再利用されたPIDへQuitしてはいけない'))
    s.psutil = SimpleNamespace(Process=lambda pid: SimpleNamespace(create_time=lambda: 9999))
    s.win32process = SimpleNamespace(GetWindowThreadProcessId=lambda hwnd: (1, 123))
    with pytest.raises(common.GateError):
        s.close()
    assert s.owned is False


@pytest.mark.parametrize('problem', ('none', 'wrong_path', 'ambiguous'))
def test_read_only_registry_is_bound_to_process_path(monkeypatch, problem):
    import sys
    class Key:
        def __init__(self, view): self.view = view
        def __enter__(self): return self
        def __exit__(self, *args): pass
    def query(key, name):
        values = dict(InstallationPath=r'C:\Other' if problem == 'wrong_path' else r'C:\Office',
                      ProductReleaseIds='Excel2021Retail' if problem == 'ambiguous' and key.view == 2 else 'Excel2024Retail',
                      VersionToReport='16.0.1.2', Platform='x64')
        return values[name], 1
    fake = SimpleNamespace(KEY_WOW64_64KEY=1, KEY_WOW64_32KEY=2, HKEY_LOCAL_MACHINE=3, KEY_READ=4,
                           OpenKey=lambda h, p, reserved, access: Key(access & 3), QueryValueEx=query)
    monkeypatch.setitem(sys.modules, 'winreg', fake)
    if problem == 'none':
        metadata = session.installation(r'C:\Office\root\Office16\EXCEL.EXE')
        assert metadata['product_release_ids'] == ['Excel2024Retail'] and metadata['platform'] == 'x64'
    else:
        with pytest.raises(common.GateError, match='Click-to-Run'):
            session.installation(r'C:\Office\root\Office16\EXCEL.EXE')


def test_no_process_termination_implementation_or_legacy_modification():
    source = '\n'.join(p.read_text(encoding='utf-8') for p in (cases.REPO / 'tools/excel2024_validation').glob('*.py'))
    for prohibited in ('.kill(', '.terminate(', 'taskkill', 'Stop-Process', 'GetActiveObject'):
        assert prohibited not in source
    assert 'DisplayAlerts = True' in source
    assert 'regserver' not in source and 'winreg.SetValue' not in source


@pytest.mark.parametrize('raw,expected', [
    ('HomeBusiness2024Retail', ['HomeBusiness2024Retail']),
    ('HomeBusiness2024Retail,OneNoteFreeRetail,', ['HomeBusiness2024Retail', 'OneNoteFreeRetail']),
    ('; HOMEBUSINESS2024RETAIL ;; homebusiness2024retail ,', ['HomeBusiness2024Retail']),
    (['HomeBusiness2024Retail', 'OneNoteFreeRetail', ''], ['HomeBusiness2024Retail', 'OneNoteFreeRetail']),
    (['HomeBusiness2024Retail,OneNoteFreeRetail', 'ProofingTools'], ['HomeBusiness2024Retail', 'OneNoteFreeRetail', 'ProofingTools']),
    ('', []), (', ; ', []), ([], []),
])
def test_registry_id_parsing_empty_tokens_case_and_multi_string(raw, expected):
    assert session.parse_product_ids(raw) == expected


@pytest.mark.parametrize('raw', [None, 123, True, b'HomeBusiness2024Retail', ['HomeBusiness2024Retail', None], {'id': 'HomeBusiness2024Retail'}])
def test_invalid_product_id_value_types_are_blocked_with_evidence(raw):
    with pytest.raises(common.GateError) as caught:
        session.parse_product_ids(raw)
    assert caught.value.status == 'BLOCKED'
    assert caught.value.details['code'] == 'INVALID_PRODUCT_IDS_TYPE'
    assert caught.value.details['detected_value'] == raw


@pytest.mark.parametrize('raw,expected_code', [
    ('HomeBusiness2024Retail', None),
    ('HomeBusiness2024Retail,OneNoteFreeRetail', None),
    ('HomeBusiness2024Retail,LanguagePack,ProofingTools', None),
    ('homebusiness2024retail,onenotefreeretail,', None),
    ('HomeBusiness2024Retail,O365HomePremRetail', 'SUBSCRIPTION_ID_PRESENT'),
    ('HomeBusiness2024Retail,M365AppsRetail', 'SUBSCRIPTION_ID_PRESENT'),
    ('HomeBusiness2024Retail,Excel2021Retail', 'OLDER_EXCEL_ID_PRESENT'),
    ('HomeBusiness2024Retail,HomeBusiness2019Retail', 'OLDER_EXCEL_ID_PRESENT'),
    ('HomeBusiness2024Retail,Unknown2024Retail', 'UNKNOWN_PRODUCT_ID_PRESENT'),
    ('HomeBusiness2024Retail.16', 'UNKNOWN_PRODUCT_ID_PRESENT'),
    ('HomeBusiness2024Retail,VisioPro2024Retail', 'UNKNOWN_PRODUCT_ID_PRESENT'),
    ('OneNoteFreeRetail', 'NO_SUPPORTED_EXCEL_2024_ID'),
    ('LanguagePack,ProofingTools', 'NO_SUPPORTED_EXCEL_2024_ID'),
    ('', 'NO_SUPPORTED_EXCEL_2024_ID'),
])
def test_home_business_2024_companions_are_not_excel_but_ambiguity_still_blocks(raw, expected_code):
    metadata = dict(product_release_ids=session.parse_product_ids(raw))
    product = 'Microsoft Office Home and Business 2024'
    if expected_code:
        with pytest.raises(common.GateError) as caught:
            session.check_product(product, True, metadata)
        assert expected_code in caught.value.details['rejection_reasons']
        assert caught.value.details['accepted'] is False
    else:
        decision = session.check_product(product, True, metadata)
        assert decision['accepted'] is True
        assert decision['excel_2024_ids'] == ['HomeBusiness2024Retail']


def test_rejected_product_ids_and_reasons_survive_both_logs(tmp_path):
    log = common.Progress(tmp_path, 'test', 'doctor')
    metadata = dict(product_release_ids=['HomeBusiness2024Retail', 'O365HomePremRetail'])
    with pytest.raises(common.GateError) as caught:
        session.check_installation_ids(metadata, progress=log)
    decision = run.events(log.path)[-1]['diagnostics']
    assert decision['subscription_ids'] == ['O365HomePremRetail']
    assert decision['rejection_reasons'] == ['SUBSCRIPTION_ID_PRESENT']
    assert 'ui_license_attested' not in decision  # pre-UI checks must not fake UI confirmation
    common.write_json(tmp_path / 'error.json', common.error_info(caught.value))
    assert common.read_json(tmp_path / 'error.json')['details'] == decision


def test_companion_allowlist_never_replaces_ui_license_or_profile_check():
    metadata = dict(product_release_ids=['HomeBusiness2024Retail', 'OneNoteFreeRetail'])
    for product, licensed in [('Microsoft Office Home and Business 2024', False), ('Microsoft 365', True)]:
        with pytest.raises(common.GateError) as caught:
            session.check_product(product, licensed, metadata)
        assert 'UI_PRODUCT_OR_LICENSE_NOT_CONFIRMED' in caught.value.details['rejection_reasons']
    p = profile()
    f = facts(); f['installation'] = metadata
    with pytest.raises(common.GateError, match='環境が変化'):
        session.check_profile(p, f)


def install_registry_mock(monkeypatch, raw64, raw32=None, *, error32=None, root32=r'C:\Office'):
    import sys
    class Key:
        def __init__(self, view): self.view = view
        def __enter__(self): return self
        def __exit__(self, *args): pass
    def open_key(hive, path, reserved, access):
        assert path == session.REGISTRY_KEY and access & 4
        view = access & 3
        if view == 2 and error32 == 'missing_key':
            raise FileNotFoundError('alternate key is absent')
        if view == 2 and error32 == 'access_denied':
            raise PermissionError('cannot inspect this view')
        return Key(view)
    def query(key, name):
        values = dict(InstallationPath=root32 if key.view == 2 else r'C:\Office',
                      ProductReleaseIds=raw64 if key.view == 1 or raw32 is None else raw32,
                      VersionToReport='16.0.20430.20182', Platform='x64')
        if key.view == 2 and error32 == 'missing_value' and name == 'ProductReleaseIds':
            raise FileNotFoundError('ProductReleaseIds is missing')
        value = values[name]
        return value, 7 if isinstance(value, list) else 1
    fake = SimpleNamespace(KEY_WOW64_64KEY=1, KEY_WOW64_32KEY=2, HKEY_LOCAL_MACHINE=3, KEY_READ=4,
                           OpenKey=open_key, QueryValueEx=query)
    monkeypatch.setitem(sys.modules, 'winreg', fake)


def test_registry_diagnostics_preserve_raw_values_and_merge_equivalent_views(tmp_path, monkeypatch):
    raw64 = 'HomeBusiness2024Retail,OneNoteFreeRetail,'
    raw32 = ['onenotefreeretail', 'homebusiness2024retail']
    install_registry_mock(monkeypatch, raw64, raw32)
    log = common.Progress(tmp_path, 'test', 'doctor')
    metadata = session.installation(r'C:\Office\root\Office16\EXCEL.EXE', progress=log)
    assert metadata['product_release_ids'] == ['HomeBusiness2024Retail', 'OneNoteFreeRetail']
    record = run.events(log.path)[-1]['diagnostics']
    assert record['matched_candidate_count'] == 1
    assert record['registry_views'][0]['values']['ProductReleaseIds']['value'] == raw64
    assert record['registry_views'][1]['values']['ProductReleaseIds']['value'] == raw32
    assert [o['view'] for o in record['registry_views']] == ['64bit', '32bit']


@pytest.mark.parametrize('problem', ['missing_key', 'access_denied', 'missing_value', 'bad_type', 'conflicting_ids', 'other_path'])
def test_registry_views_missing_unreadable_conflicting_and_other_installation(tmp_path, monkeypatch, problem):
    raw = 'HomeBusiness2024Retail,OneNoteFreeRetail'
    raw32 = {'unexpected': 'value'} if problem == 'bad_type' else 'Excel2021Retail' if problem in ('conflicting_ids', 'other_path') else None
    install_registry_mock(monkeypatch, raw, raw32, error32=problem,
                          root32=r'C:\Other' if problem == 'other_path' else r'C:\Office')
    log = common.Progress(tmp_path, 'test', 'doctor')
    if problem in ('missing_key', 'other_path'):
        metadata = session.installation(r'C:\Office\root\Office16\EXCEL.EXE', progress=log)
        assert metadata['product_release_ids'] == ['HomeBusiness2024Retail', 'OneNoteFreeRetail']
    else:
        with pytest.raises(common.GateError) as caught:
            session.installation(r'C:\Office\root\Office16\EXCEL.EXE', progress=log)
        assert caught.value.details['registry_views'][1]['status'] in ('READ_FAILED', 'PATH_MATCHED')
        assert caught.value.details['rejection_reasons']
    assert run.events(log.path)[-1]['event'] == 'installation.observed'


def test_worker_preserves_observed_identity_and_rejection_before_ui(tmp_path, monkeypatch):
    log_holder = []
    class RejectedSession:
        def __init__(self, log, profile=None):
            log_holder.append(log)
            self.owned = False; self.facts = None; self.observed_facts = None
        def connect(self):
            self.owner = dict(pid=123, path=r'C:\Office\EXCEL.EXE', created=1000)
            self.owned = True
            self.observed_facts = facts()
            self.observed_facts['installation'] = dict(product_release_ids=['HomeBusiness2024Retail', 'O365HomePremRetail'])
            session.check_installation_ids(self.observed_facts['installation'], progress=log_holder[0])
        def close(self): pass
        def release(self): pass
    monkeypatch.setattr(worker, 'Session', RejectedSession)
    common.write_json(tmp_path / 'request.json', dict(run_id='test', case_id='doctor', source_sha256=run.source_identity()['source_sha256']))
    assert worker.main(tmp_path / 'request.json') == 1
    result = common.read_json(tmp_path / 'worker.json')
    assert result['status'] == 'BLOCKED' and result['owned_pid'] == 123
    assert result['observed_identity']['installation']['product_release_ids'] == ['HomeBusiness2024Retail', 'O365HomePremRetail']
    assert result['first_error']['details']['subscription_ids'] == ['O365HomePremRetail']
    assert 'profile' not in result and not (tmp_path / 'ui.json').exists()
