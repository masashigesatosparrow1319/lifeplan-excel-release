"""Bounded state observation only; never a substitute for Windows Excel."""
from types import SimpleNamespace

import pytest

from tools.excel2024_validation import common, run, worker


@pytest.fixture
def clock(monkeypatch):
    now = [0.0]
    monkeypatch.setattr(worker.time, 'monotonic', lambda: now[0])
    # Simulate a descheduled worker; avoid any real 180-second waits.
    monkeypatch.setattr(worker.time, 'sleep', lambda seconds: now.__setitem__(0, now[0] + 60))
    monkeypatch.setattr(worker, 'pump_messages', lambda: None)
    return now


def state_excel(states):
    values = iter(states)
    class Excel:
        @property
        def CalculationState(self):
            return next(values)
    return Excel()


def test_pending_and_calculating_must_reach_done_before_deadline(tmp_path, clock):
    log = common.Progress(tmp_path, 'test', 'tiny3')
    worker.await_done(state_excel([1, 2, 0]), log)
    records = run.events(log.path)
    observations = [e for e in records if e['event'] == 'calculation.observed']
    assert [(e['state_name'], e['elapsed_seconds']) for e in observations] == [
        ('xlCalculating', 0), ('xlPending', 60), ('xlDone', 120)]
    assert records[-1]['event'] == 'calculation.wait_done'


@pytest.mark.parametrize('state', [1, 2])
def test_continuous_non_done_times_out_without_an_extra_com_read(tmp_path, clock, state):
    reads = []
    class Excel:
        @property
        def CalculationState(self):
            reads.append(clock[0])
            return state
    log = common.Progress(tmp_path, 'test', 'tiny3')
    with pytest.raises(common.GateError) as error:
        worker.await_done(Excel(), log)
    assert error.value.status == 'TIMEOUT'
    assert error.value.details['last_state'] == state
    assert error.value.details['state_counts'] == {str(state): 3}
    assert reads == [0, 60, 120]
    assert clock[0] == 180
    assert run.events(log.path)[-1]['event'] == 'calculation.wait_timeout'


def test_state_read_deadline_uses_remaining_budget(tmp_path, clock):
    clock[0] = 0
    reads = []
    class Excel:
        @property
        def CalculationState(self):
            reads.append(clock[0])
            return 2
    log = common.Progress(tmp_path, 'test', 'tiny3')
    # First scheduling delay consumes 170 of the 180 seconds.
    worker.time.sleep = lambda seconds: clock.__setitem__(0, clock[0] + 170)
    with pytest.raises(common.GateError):
        worker.await_done(Excel(), log)
    before = [e for e in run.events(log.path) if e['event'] == 'calculation_state.before_call']
    assert [e['timeout_seconds'] for e in before] == [60, 10]
    assert reads == [0, 170]


def test_done_returned_after_deadline_is_never_accepted(tmp_path, clock):
    class Excel:
        @property
        def CalculationState(self):
            clock[0] = 181
            return 0
    log = common.Progress(tmp_path, 'test', 'tiny3')
    with pytest.raises(common.GateError) as error:
        worker.await_done(Excel(), log)
    assert error.value.status == 'TIMEOUT'
    assert not any(e['event'] == 'calculation.wait_done' for e in run.events(log.path))


@pytest.mark.parametrize('state', [-1, 3])
def test_unknown_state_is_failure(tmp_path, clock, state):
    log = common.Progress(tmp_path, 'test', 'tiny3')
    with pytest.raises(common.GateError) as error:
        worker.await_done(state_excel([state]), log)
    assert error.value.status == 'FAIL'


def test_call_rejected_is_logged_without_retry_or_pending_success(tmp_path, clock):
    class Rejected(RuntimeError):
        hresult = -2147418111  # RPC_E_CALL_REJECTED
    class Excel:
        @property
        def CalculationState(self):
            raise Rejected('Call was rejected by callee')
    log = common.Progress(tmp_path, 'test', 'tiny3')
    with pytest.raises(Rejected):
        worker.await_done(Excel(), log)
    records = run.events(log.path)
    assert records[-1]['event'] == 'calculation_state.failed'
    assert records[-1]['error']['hresult'] == -2147418111
    assert not any(e['event'] == 'calculation.wait_done' for e in records)


def test_stop_is_respected_before_state_read(tmp_path, clock):
    log = common.Progress(tmp_path, 'test', 'tiny3')
    log.stop.touch()
    with pytest.raises(common.GateError) as error:
        worker.await_done(state_excel([]), log)
    assert error.value.status == 'INTERRUPTED'
    assert not any(e['event'] == 'calculation_state.before_call' for e in run.events(log.path))


def test_updated_values_with_pending_state_are_observation_only(tmp_path, clock):
    calls = []
    class Excel:
        CalculationState = 2
        Calculation = -4135
        def CalculateFullRebuild(self):
            calls.append('full_rebuild')
    values = {'B1': 10, 'B2': 21, 'B3': 31}
    manifest = dict(probes=[['診断', cell, value] for cell, value in values.items()])
    book = SimpleNamespace(Worksheets=SimpleNamespace(Item=lambda sheet: SimpleNamespace(
        Range=lambda cell: SimpleNamespace(Value2=values[cell]))))
    log = common.Progress(tmp_path, 'test', 'tiny3')
    excel = Excel()
    with pytest.raises(common.GateError) as error:
        worker.recalculate(excel, book, log, 'final', manifest=manifest)
    assert error.value.status == 'TIMEOUT'
    assert calls == ['full_rebuild']
    assert excel.Calculation == -4105 and book.ForceFullCalculation is True
    records = run.events(log.path)
    sample = next(e['sample'] for e in records if e['event'] == 'snapshot')
    assert sample['phase'] == 'final.before_wait'
    assert sample['state'] == 2 and sample['mode'] == -4105
    assert sample['differences'] == []
    assert sample['values'] == {'診断!' + cell: value for cell, value in values.items()}
    assert not any(e['event'] == 'calculation.wait_done' for e in records)


@pytest.mark.parametrize('observed_mode', [-4135, 2])
def test_automatic_read_mismatch_stops_before_state_wait(tmp_path, monkeypatch, observed_mode):
    class Excel:
        @property
        def Calculation(self):
            return observed_mode
        @Calculation.setter
        def Calculation(self, value):
            assert value == -4105
        def CalculateFullRebuild(self):
            pass
    monkeypatch.setattr(worker, 'await_done', lambda *args: pytest.fail('mode mismatch must not wait'))
    log = common.Progress(tmp_path, 'test', 'tiny3')
    with pytest.raises(common.GateError) as error:
        worker.recalculate(Excel(), SimpleNamespace(), log, 'final')
    assert error.value.status == 'FAIL'
    mode = next(e for e in run.events(log.path) if e['event'] == 'calculation.automatic_observed')
    assert mode['mode'] == observed_mode
