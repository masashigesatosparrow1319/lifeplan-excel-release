"""Doctor aggregation only: no COM, workbook generation, or calculation tests."""
import copy
import json
from pathlib import Path
import sys

import pytest

from tools.excel2024_validation import SCHEMA
from tools.excel2024_validation import common, run


@pytest.fixture(autouse=True)
def forbid_calculation_stages(monkeypatch):
    def forbidden(*args, **kwargs):
        pytest.fail('doctor集約試験では3段階の生成/計算を実行しません')
    monkeypatch.setattr(run, 'prepare', forbidden)


def configured_result(directory, request):
    common.write_json(directory / 'request.json', request)
    identity = dict(pid=123, created=1000, path=r'C:\Office\EXCEL.EXE', version='16.0',
                    exe_sha256='MOCK-NOT-REAL-EXCEL', file_version='16.0.20430.20182', build='20430',
                    installation=dict(product_release_ids=['HomeBusiness2024Retail']))
    profile = dict(schema=SCHEMA, product='Microsoft Office Home and Business 2024', licensed=True,
                   facts=copy.deepcopy(identity), ui_confirmed_on_owned_pid=True)
    return dict(schema=SCHEMA, run_id=request['run_id'], case_id='doctor', status='CONFIGURED',
                profile=profile, identity=identity, first_error=None, cleanup_errors=[], progress_log_errors=[],
                force_termination=False, worker_still_running=False,
                request_sha256=common.sha(directory / 'request.json'))


def execute_doctor(tmp_path, *, monitor=configured_result, residual=lambda result: False):
    def no_units(*args):
        pytest.fail('doctorでは単体試験起動や計算試験へ進みません')
    return run.execute('doctor', tmp_path / 'profile.json', tmp_path / 'run', check=lambda: None,
                       monitor=monitor, residual=residual, units=no_units)


def test_doctor_success_requires_every_parent_check_and_profile_saved(tmp_path):
    value, root = execute_doctor(tmp_path)
    assert value['status'] == value['doctor']['status'] == 'CONFIGURED'
    assert value['exit_code'] == 0 and value['excel_recalculation_verified'] is False
    assert value['doctor_aggregation']['phase'] == 'complete'
    assert value['doctor_aggregation']['checks'] == {
        'worker_evidence': 'PASS', 'product_and_profile': 'PASS', 'worker_cleanup': 'PASS',
        'residual': 'PASS', 'profile_save': 'PASS',
    }
    assert value['doctor_aggregation']['residual_check']['same_owned_process_remains'] is False
    assert common.read_json(tmp_path / 'profile.json') == value['doctor']['profile']
    assert common.read_json(root / 'run.json') == value
    assert all(s['status'] == 'NOT_RUN' for s in value['stages'])


def test_doctor_configured_child_with_owned_excel_remaining_is_still_fail(tmp_path):
    value, root = execute_doctor(tmp_path, residual=lambda result: True)
    child = value['doctor']
    # Reproduce all fields in the owner's report, without assuming its cause.
    assert child['status'] == 'CONFIGURED' and child['first_error'] is None
    assert child['cleanup_errors'] == child['progress_log_errors'] == []
    assert child['worker_still_running'] is False and child['force_termination'] is False
    assert value['status'] == 'FAIL' and value['exit_code'] == 1
    assert value['reason_code'] == 'OWNED_EXCEL_REMAINS'
    aggregation = value['doctor_aggregation']
    assert aggregation['failed_at'] == 'residual'
    assert aggregation['residual_check']['pid'] == 123
    assert aggregation['residual_check']['same_owned_process_remains'] is True
    assert aggregation['checks']['residual'] == 'FAIL'
    assert not (tmp_path / 'profile.json').exists()
    assert common.read_json(root / 'run.json')['doctor']['status'] == 'CONFIGURED'


@pytest.mark.parametrize('problem,phase,status', [
    ('request', 'worker_evidence', 'FAIL'),
    ('run', 'worker_evidence', 'FAIL'),
    ('schema', 'worker_evidence', 'FAIL'),
    ('identity', 'product_and_profile', 'FAIL'),
    ('license', 'product_and_profile', 'BLOCKED'),
    ('pid_ui', 'product_and_profile', 'BLOCKED'),
    ('primary', 'worker_cleanup', 'FAIL'),
    ('cleanup', 'worker_cleanup', 'FAIL'),
    ('progress', 'worker_cleanup', 'FAIL'),
    ('worker_alive', 'worker_cleanup', 'FAIL'),
    ('cleanup_pending', 'worker_cleanup', 'FAIL'),
])
def test_doctor_parent_rejects_incomplete_evidence_without_status_rewriting(tmp_path, problem, phase, status):
    def monitor(directory, request):
        result = configured_result(directory, request)
        if problem == 'request': result['request_sha256'] = 'different'
        if problem == 'run': result['run_id'] = 'old-run'
        if problem == 'schema': result['schema'] = 'old-schema'
        if problem == 'identity': result['profile']['facts']['pid'] = 999
        if problem == 'license': result['profile']['licensed'] = False
        if problem == 'pid_ui': result['profile']['ui_confirmed_on_owned_pid'] = False
        if problem == 'primary': result['first_error'] = dict(message='child primary error')
        if problem == 'cleanup': result['cleanup_errors'] = [dict(message='Quit error')]
        if problem == 'progress': result['progress_log_errors'] = [dict(message='log write error')]
        if problem == 'worker_alive': result['worker_still_running'] = True
        if problem == 'cleanup_pending': result['cleanup_pending'] = True
        return result
    def residual(result):
        pytest.fail('前段の検証失敗時は残留チェック/保存へ進みません')
    value, root = execute_doctor(tmp_path, monitor=monitor, residual=residual)
    assert value['status'] == status
    assert value['doctor']['status'] == 'CONFIGURED'  # preserve the worker's evidence
    assert value['doctor_aggregation']['failed_at'] == phase
    assert value['first_error']['message']
    assert value['exit_code'] == (2 if status == 'BLOCKED' else 1)
    assert not (tmp_path / 'profile.json').exists()
    assert common.read_json(root / 'run.json')['first_error'] == value['first_error']


@pytest.mark.parametrize('worker_status,expected', [
    ('FAIL', 'FAIL'), ('BLOCKED', 'BLOCKED'), ('TIMEOUT', 'TIMEOUT'),
    ('INTERRUPTED', 'INTERRUPTED'), ('NOT_RUN', 'FAIL'), ('PASS', 'FAIL'), (None, 'FAIL'),
])
def test_doctor_worker_failure_timeout_and_unexpected_status_stay_distinct(tmp_path, worker_status, expected):
    def monitor(directory, request):
        return dict(status=worker_status, first_error=None, cleanup_errors=[])
    def residual(result):
        pytest.fail('子がCONFIGUREDでなければ残留チェック/保存へ進みません')
    value, _ = execute_doctor(tmp_path, monitor=monitor, residual=residual)
    assert value['status'] == expected and value['doctor']['status'] == worker_status
    assert value['reason'] and value['reason_code'] == 'DOCTOR_WORKER_NOT_CONFIGURED'
    assert value['doctor_aggregation']['failed_at'] == 'worker_status'
    assert not (tmp_path / 'profile.json').exists()
    assert all(s['status'] == 'NOT_RUN' for s in value['stages'])


@pytest.mark.parametrize('failure', ['permission', 'invalid_result', 'interrupted'])
def test_doctor_residual_observer_error_is_not_misreported_as_product_error(tmp_path, failure):
    def residual(result):
        if failure == 'permission': raise PermissionError('cannot read process')
        if failure == 'interrupted': raise KeyboardInterrupt()
        return None
    value, _ = execute_doctor(tmp_path, residual=residual)
    assert value['status'] == ('INTERRUPTED' if failure == 'interrupted' else 'FAIL')
    assert value['doctor']['status'] == 'CONFIGURED' and value['doctor']['first_error'] is None
    assert value['first_error']['type'] == ('KeyboardInterrupt' if failure == 'interrupted' else 'PermissionError' if failure == 'permission' else 'ValueError')
    assert value['doctor_aggregation']['failed_at'] == 'residual'
    assert 'same_owned_process_remains' not in value['doctor_aggregation']['residual_check']
    assert not (tmp_path / 'profile.json').exists()


def test_doctor_profile_write_permission_error_preserves_parent_exception(tmp_path, monkeypatch):
    original_open = Path.open
    def denied(path, mode='r', *args, **kwargs):
        if path == tmp_path / 'profile.json' and mode == 'x':
            raise PermissionError('profile write denied')
        return original_open(path, mode, *args, **kwargs)
    monkeypatch.setattr(Path, 'open', denied)
    value, root = execute_doctor(tmp_path)
    assert value['status'] == 'FAIL' and value['doctor']['status'] == 'CONFIGURED'
    assert value['doctor']['first_error'] is None
    assert value['first_error']['type'] == 'PermissionError'
    assert value['reason'] == 'profile write denied'
    assert value['doctor_aggregation']['failed_at'] == 'profile_save'
    assert value['exit_code'] == 1
    assert common.read_json(root / 'run.json')['first_error'] == value['first_error']


def test_doctor_profile_created_by_another_run_is_not_overwritten(tmp_path):
    def monitor(directory, request):
        (tmp_path / 'profile.json').write_text('OTHER-RUN-PROFILE', encoding='utf-8')
        return configured_result(directory, request)
    value, _ = execute_doctor(tmp_path, monitor=monitor)
    assert value['status'] == 'FAIL' and value['first_error']['type'] == 'FileExistsError'
    assert value['doctor_aggregation']['failed_at'] == 'profile_save'
    assert (tmp_path / 'profile.json').read_text(encoding='utf-8') == 'OTHER-RUN-PROFILE'


def test_doctor_timeout_cannot_be_overridden_by_late_worker_completion(tmp_path):
    def monitor(directory, request):
        common.write_json(directory / 'worker.json', configured_result(directory, request))
        return dict(status='TIMEOUT', reason='call deadline', worker_still_running=True)
    value, root = execute_doctor(tmp_path, monitor=monitor)
    assert common.read_json(root / 'doctor/worker.json')['status'] == 'CONFIGURED'
    assert value['doctor']['status'] == value['status'] == 'TIMEOUT'
    assert value['exit_code'] == 1 and not (tmp_path / 'profile.json').exists()


@pytest.mark.parametrize('scenario,expected_code', [('normal', 0), ('residual', 1), ('exception', 1), ('blocked', 2), ('timeout', 1)])
def test_doctor_console_run_json_and_exit_code_report_both_parent_and_child(tmp_path, monkeypatch, capsys, scenario, expected_code):
    def residual(result):
        if scenario == 'exception': raise PermissionError('parent process query denied')
        return scenario == 'residual'
    def monitor(directory, request):
        if scenario in ('blocked', 'timeout'):
            return dict(status='BLOCKED' if scenario == 'blocked' else 'TIMEOUT', first_error=None)
        return configured_result(directory, request)
    value, root = execute_doctor(tmp_path, monitor=monitor, residual=residual)
    monkeypatch.setattr(run, 'execute', lambda *args: (value, root))
    monkeypatch.setattr(sys, 'argv', ['gate2', 'doctor'])
    assert run.main() == expected_code
    summary = json.loads(capsys.readouterr().out)
    stored = common.read_json(root / 'run.json')
    assert summary['status'] == stored['status']
    assert summary['exit_code'] == stored['exit_code'] == expected_code
    assert summary['doctor_status'] == stored['doctor']['status']
    assert summary['first_error'] == stored.get('first_error')
    assert summary['identity_error'] == stored['doctor'].get('first_error')
    assert summary['doctor_aggregation'] == stored['doctor_aggregation']
    if scenario == 'exception':
        assert summary['identity_error'] is None and summary['first_error']['type'] == 'PermissionError'
    if scenario == 'residual':
        assert summary['identity_error'] is None and 'PID=123' in summary['reason']
