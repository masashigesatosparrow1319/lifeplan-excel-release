"""Owned-session cleanup and bounded observation; no Excel/calculation runs."""
from types import SimpleNamespace

import pytest

from tools.excel2024_validation import common, run, session


IDENTITY = dict(pid=123, created=100, path=r'C:\Office\EXCEL.EXE')


def observe(process, *, timeout=10):
    now = [0.0]
    result = dict(identity=IDENTITY.copy())
    def pause(seconds):
        now[0] += seconds
    remains = run.residual_check(result, process=lambda pid: process(now[0]),
                                 clock=lambda: now[0], pause=pause, timeout=timeout)
    return remains, result['residual_observation'], now[0]


def owned():
    return SimpleNamespace(create_time=lambda: 100, exe=lambda: IDENTITY['path'])


@pytest.mark.parametrize('exit_after', [0, 0.5, 3, 9])
def test_exit_is_observed_including_after_old_two_second_window(exit_after):
    def process(elapsed):
        if elapsed >= exit_after:
            raise ProcessLookupError()
        return owned()
    remains, detail, elapsed = observe(process)
    assert remains is False
    assert detail['outcome'] == 'PROCESS_EXITED'
    assert exit_after <= elapsed <= exit_after + .11


def test_real_residual_at_deadline_is_not_success():
    remains, detail, elapsed = observe(lambda elapsed: owned())
    assert remains is True
    assert detail['outcome'] == 'OWNED_PROCESS_REMAINS'
    assert elapsed == 10
    assert detail['samples'][-1]['path'] == IDENTITY['path']


def test_recycled_pid_is_not_the_owned_process():
    remains, detail, _ = observe(lambda elapsed: SimpleNamespace(create_time=lambda: 101))
    assert remains is False and detail['outcome'] == 'PID_REUSED'


@pytest.mark.parametrize('problem', ['path', 'access'])
def test_inability_to_verify_identity_never_becomes_success(problem):
    def process(elapsed):
        if problem == 'access':
            raise PermissionError('access denied')
        return SimpleNamespace(create_time=lambda: 100, exe=lambda: r'C:\Other\EXCEL.EXE')
    with pytest.raises((ValueError, PermissionError)):
        observe(process)


def test_missing_identity_is_not_success():
    with pytest.raises(ValueError):
        run.residual_check({})


def test_release_drops_wrapper_and_collects_before_sta_uninitialize(tmp_path, monkeypatch):
    s = session.Session(common.Progress(tmp_path, 'test', 'doctor'))
    s.excel = object()
    s.initialized = True
    calls = []
    def collect():
        assert s.excel is None
        calls.append('collect')
        return 2
    monkeypatch.setattr(session.gc, 'collect', collect)
    s.pythoncom = SimpleNamespace(CoUninitialize=lambda: calls.append('uninitialize'))
    s.release()
    assert calls == ['collect', 'uninitialize']
    assert s.initialized is False
    assert run.events(s.log.path)[0]['collected'] == 2


def test_release_log_failure_still_uninitializes_sta(tmp_path):
    s = session.Session(common.Progress(tmp_path, 'test', 'doctor'))
    s.initialized = True
    calls = []
    s.pythoncom = SimpleNamespace(CoUninitialize=lambda: calls.append('uninitialize'))
    def fail(*args, **kwargs):
        raise OSError('log failure')
    s.log.emit = fail
    with pytest.raises(OSError):
        s.release()
    assert calls == ['uninitialize']


@pytest.mark.parametrize('quit_error', [False, True])
def test_owned_quit_diagnostics_do_not_suppress_rpc_exception(tmp_path, quit_error):
    s = session.Session(common.Progress(tmp_path, 'test', 'doctor'))
    s.owned = True
    calls = []
    s.confirm_owner = lambda: calls.append('identity')
    def quit_excel():
        calls.append('quit')
        if quit_error:
            raise RuntimeError('RPC disconnected')
    s.excel = SimpleNamespace(Ready=True, Interactive=True, DisplayAlerts=True, Quit=quit_excel)
    if quit_error:
        with pytest.raises(RuntimeError, match='RPC disconnected'):
            s.close()
    else:
        s.close()
    assert calls == ['identity', 'identity', 'quit']
    observations = [e for e in run.events(s.log.path) if e['event'] == 'cleanup.observed']
    assert [e['property'] for e in observations] == ['Ready', 'Interactive', 'DisplayAlerts']


def test_read_only_diagnostic_exception_does_not_skip_owned_quit(tmp_path):
    s = session.Session(common.Progress(tmp_path, 'test', 'doctor'))
    s.owned = True
    s.confirm_owner = lambda: None
    calls = []
    class Excel:
        @property
        def Ready(self):
            raise RuntimeError('property unavailable')
        Interactive = True
        DisplayAlerts = True
        def Quit(self):
            calls.append('quit')
    s.excel = Excel()
    s.close()
    assert calls == ['quit']
    assert any(e['event'] == 'cleanup.observation_unavailable' for e in run.events(s.log.path))
