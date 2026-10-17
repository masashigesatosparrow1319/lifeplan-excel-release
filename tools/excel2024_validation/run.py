"""Gate 2 supervisor. Three stages only; no forced termination or auto retries."""
import argparse
from datetime import datetime
import html
import importlib.util
import json
import ntpath
import os
import platform
from pathlib import Path
import subprocess
import sys
import time
import threading
import uuid
import xml.etree.ElementTree as ET

from . import SCHEMA, STAGES, E1_DIAGNOSTIC, E1_PROTOCOL, F24_DIAGNOSTIC, F24_PROTOCOL, A_DIAGNOSTIC, A_PROTOCOL
from .cases import REPO, prepare, verify_saved
from .common import GateError, create_run, error_info, read_json, sha, utc, write_json
from .session import check_profile

DEFAULT_PROFILE = REPO.parent / 'lifeplan-excel2024-profile.json'


def preflight():
    if not __debug__:
        return 'python -O/PYTHONOPTIMIZEで検証assertを無効化しないでください'
    if sys.platform != 'win32':
        return 'BLOCKED_EXCEL_UNAVAILABLE: この環境はWindowsではありません'
    missing = [name for name in ('psutil', 'pythoncom', 'win32com', 'win32api', 'win32process')
               if importlib.util.find_spec(name) is None]
    return '追加依存がありません: ' + ', '.join(missing) if missing else None


def events(path):
    if not path.exists():
        return []
    result = []
    for line in path.read_text(encoding='utf-8').splitlines():
        try:
            result.append(json.loads(line))
        except ValueError:
            continue  # incomplete trailing line during a write
    return result


def ask_ui(event):
    facts = event['identity']
    print(f"専用Excel PID={facts['pid']} / {facts['path']}", flush=True)
    print('この専用ウィンドウの「ファイル→アカウント」で確認してください（60秒以内）。', flush=True)
    product = input('表示された製品名（Excel/Office 2024）: ')
    licensed = input('同じ専用ウィンドウでライセンス認証済みを確認しましたか [yes]: ') == 'yes'
    return dict(pid=facts['pid'], product=product, licensed=licensed)


def supervise(directory, request, *, popen=subprocess.Popen, clock=time.monotonic,
              pause=time.sleep, ui=ask_ui):
    """Never wait indefinitely or send termination signals, even to our worker."""
    directory = Path(directory)
    if request.get('diagnostic') == A_DIAGNOSTIC:
        from .preopen_input import publish_json
        publish_json(directory / 'request.json', request)
    else:
        write_json(directory / 'request.json', request)
    started = clock()
    seen, in_flight, identity, ui_done = 0, {}, None, False
    proc = None
    outcome = None
    records = []
    ui_errors = []
    python_processes = None
    try:
        with (directory / 'worker-console.txt').open('w', encoding='utf-8') as console:
            proc = popen([sys.executable, '-m', 'tools.excel2024_validation.worker',
                          '--request', str(directory / 'request.json')], cwd=REPO,
                         stdout=console, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL)
        if request.get('diagnostic') in (F24_DIAGNOSTIC, A_DIAGNOSTIC):
            from .functions24 import process_identity
            python_processes = dict(supervisor=process_identity(os.getpid()), launcher=process_identity(proc.pid))
        while True:
            records = events(directory / 'progress.jsonl')
            for event in records[seen:]:
                if event.get('run_id') != request['run_id'] or event.get('case_id') != request['case_id']:
                    raise ValueError('別run/caseの進行ログです')
                name = event['event']
                if name == 'python.worker_started' and python_processes is not None:
                    if 'worker' in python_processes or event['worker_pid'] != event['process']['pid']:
                        raise ValueError('Python worker起動証拠の重複/PID不一致')
                    python_processes['worker'] = event['process']
                if 'identity' in event:
                    identity = event['identity']
                if name.endswith('.before_call'):
                    # Timestamp-based deadline catches a call even when the
                    # supervisor did not see its first few milliseconds.
                    elapsed = max(0, time.time() - datetime.fromisoformat(event['utc']).timestamp())
                    in_flight[name.removesuffix('.before_call')] = clock() + event['timeout_seconds'] - elapsed
                elif name.endswith(('.after_call', '.failed')):
                    in_flight.pop(name.rsplit('.', 1)[0], None)
                if name == 'doctor.ui_request' and not ui_done:
                    def confirm(event=event):
                        try:
                            response = ui(event)
                            if not (directory / 'STOP').exists():
                                write_json(directory / 'ui.json', response)
                        except BaseException as error:
                            ui_errors.append(error_info(error))
                    # Input waits must not block the supervisor's COM deadlines.
                    threading.Thread(target=confirm, daemon=True).start()
                    ui_done = True
            seen = len(records)
            if ui_errors:
                outcome = dict(status='INTERRUPTED', reason='UI確認が完了しませんでした', ui_errors=ui_errors)
                break
            expired = [name for name, deadline in in_flight.items() if clock() >= deadline]
            if expired or clock() - started > 600:
                outcome = dict(status='TIMEOUT', reason='同期COM/操作期限超過', pending_calls=expired)
                break
            if proc.poll() is not None:
                result = directory / 'worker.json'
                if not result.is_file():
                    outcome = dict(status='FAIL', reason='子プロセス最終ログなし', worker_exit=proc.returncode)
                else:
                    outcome = read_json(result)
                    if proc.returncode != 0 and outcome.get('status') in ('SAVED_REOPENED', 'CONFIGURED'):
                        outcome.update(status='FAIL', reason='子プロセス終了コード不一致')
                break
            pause(.1)
    except KeyboardInterrupt:
        outcome = dict(status='INTERRUPTED', reason='所有者がCtrl+Cで中断')
    except BaseException as error:
        outcome = dict(status='FAIL', first_error=error_info(error))
    if outcome['status'] in ('TIMEOUT', 'INTERRUPTED', 'FAIL'):
        # A blocked synchronous COM call may return later. It must then stop
        # before any edit/rebuild/save; only owned-instance cleanup is allowed.
        (directory / 'STOP').touch(exist_ok=True)
    outcome.update(worker_pid=proc.pid if proc else None, last_progress=records[-1] if records else None,
                   observed_identity=identity,
                   force_termination=False, worker_still_running=proc is not None and proc.poll() is None)
    if python_processes is not None:
        outcome['python_processes'] = python_processes
        if not outcome['worker_still_running'] and 'worker' in python_processes:
            try:
                from .functions24 import worker_exit_observation
                observed_exit = worker_exit_observation(python_processes['worker'])
                python_processes['worker_exit'] = observed_exit
                if observed_exit['outcome'] == 'WORKER_STILL_RUNNING':
                    outcome['worker_still_running'] = True
            except BaseException as error:
                python_processes['worker_exit_error'] = error_info(error)
                if outcome['status'] == 'SAVED_REOPENED':
                    outcome.update(status='INTERRUPTED' if isinstance(error, KeyboardInterrupt) else 'FAIL',
                                   reason='Python worker終了の確認不能')
        if outcome['status'] == 'SAVED_REOPENED' and outcome['worker_still_running']:
            outcome.update(status='FAIL', reason='Python worker終了処理が未完了')
        if outcome['status'] in ('TIMEOUT', 'INTERRUPTED', 'FAIL'):
            (directory / 'STOP').touch(exist_ok=True)
    if outcome['worker_still_running']:
        outcome['cleanup_pending'] = True
        outcome['action'] = '専用PIDをログで確認し、必要なら所有者がExcelを通常操作で閉じてください。既存Excelは閉じません'
    return outcome


def validate_runtime(result, directory, request, profile):
    if result.get('schema') != SCHEMA or result.get('run_id') != request['run_id'] or result.get('case_id') != request['case_id']:
        raise ValueError('別実行の子プロセス結果です')
    if result.get('request_sha256') != sha(directory / 'request.json'):
        raise ValueError('実行要求が変更されています')
    if result.get('manifest_sha256') != sha(directory / 'case.json'):
        raise ValueError('入力manifestが変更されています')
    if (result.get('status') != 'SAVED_REOPENED' or result.get('cleanup_errors')
            or result.get('first_error') or result.get('progress_log_errors')):
        raise ValueError('実Excelの完了/終了処理の証拠がありません')
    check_profile(profile, result['identity'])
    expected_rebuilds = 2 if request['case_id'] == 'functions24' else 1
    if result.get('full_rebuild_calls') != expected_rebuilds:
        raise ValueError('全再構築の呼出回数不一致')
    by_phase = {s['phase']: s for s in result['samples']}
    manifest = read_json(directory / 'case.json')
    phases = ('challenge_updated', 'before_save', 'reopened') if expected_rebuilds == 2 else ('before_save', 'reopened')
    for phase in phases:
        sample = by_phase.get(phase)
        if not sample or sample['state_before'] != 0 or sample['state'] != 0 or sample['mode'] != -4105 or sample['differences']:
            raise ValueError('状態/モード/数値の実測証拠不一致: ' + phase)
        wanted = manifest
        if phase == 'challenge_updated':
            wanted = dict(manifest, probes=[[s, c, 1 if c == 'L8' else '不一致' if c == 'M8' else v]
                                           for s, c, v in manifest['probes']])
        from .cases import compare_probes
        if compare_probes(wanted, lambda s, c: sample.get('values', {}).get(s + '!' + c)):
            raise ValueError('親プロセスのライブ値再照合が不一致: ' + phase)


def residual_check(result, *, process=None, clock=time.monotonic, pause=time.sleep, timeout=10):
    identity = result.get('identity')
    if not identity:
        raise ValueError('専用Excelの残留確認に必要なidentityがありません')
    if process is None:
        import psutil
        process = psutil.Process
        missing = psutil.NoSuchProcess
    else:
        missing = ProcessLookupError
    started = clock()
    observation = dict(timeout_seconds=timeout, samples=[], outcome='CHECKING')
    result['residual_observation'] = observation
    while True:
        try:
            proc = process(identity['pid'])
            created = proc.create_time()
            if created != identity['created']:
                observation['outcome'] = 'PID_REUSED'
                return False  # recycled PID is not ours
            path = proc.exe()
            if ntpath.normcase(path) != ntpath.normcase(identity['path']):
                raise ValueError('専用Excel PIDの実行ファイルパスが変化しています')
            observation['samples'].append(dict(elapsed_seconds=clock() - started,
                                               pid=identity['pid'], created=created, path=path))
        except missing:
            observation['outcome'] = 'PROCESS_EXITED'
            return False
        if clock() - started >= timeout:
            observation['outcome'] = 'OWNED_PROCESS_REMAINS'
            return True
        pause(min(.1, max(0, timeout - (clock() - started))))


def report(root, value):
    write_json(root / 'run.json', value)
    (root / 'summary.html').write_text('<!doctype html><meta charset="utf-8"><title>Excel 2024 Gate 2</title>'
                                    '<pre>' + html.escape(json.dumps(value, ensure_ascii=False, indent=2, default=str)) + '</pre>', encoding='utf-8')


def exit_code_for(status):
    return 0 if status in ('PASS', 'CONFIGURED', 'DIAGNOSTIC_PASS') else 2 if status == 'BLOCKED' else 1


def unit_tests(root):
    env = os.environ.copy()
    for name in list(env):
        if 'RECALCULATED' in name or name.startswith(('P2A_S1_', 'S1_RECALCULATED', 'P2B_MEXT_')):
            env.pop(name)
    command = [sys.executable, '-m', 'pytest', '-q', 'tests/test_excel2024_validation.py',
               'tests/test_p1.py', 'tests/test_p2a.py', '--junitxml=' + str(root / 'unit-tests.xml')]
    with (root / 'unit-tests.txt').open('w', encoding='utf-8') as console:
        proc = subprocess.Popen(command, cwd=REPO, env=env, stdout=console, stderr=subprocess.STDOUT)
    started = time.monotonic()
    while proc.poll() is None:
        if time.monotonic() - started > 300:
            raise GateError('TIMEOUT', f'Python試験期限超過。PID={proc.pid}。強制終了せずExcelへ進みません')
        time.sleep(.1)
    counts = dict(passed=0, failed=0, errors=0, skipped=0)
    if (root / 'unit-tests.xml').is_file():
        suites = ET.parse(root / 'unit-tests.xml').getroot().iter('testsuite')
        for suite in suites:
            n = {k: int(suite.get(k, 0)) for k in ('tests', 'failures', 'errors', 'skipped')}
            counts['passed'] += n['tests'] - n['failures'] - n['errors'] - n['skipped']
            for key, other in [('failed', 'failures'), ('errors', 'errors'), ('skipped', 'skipped')]:
                counts[key] += n[other]
    else:
        raise ValueError('Python試験のJUnit結果がありません')
    return dict(returncode=proc.returncode, command=command, junit='unit-tests.xml', log='unit-tests.txt', **counts)


def source_identity():
    import hashlib
    paths = sorted((REPO / 'tools/excel2024_validation').glob('*.py'))
    paths += [REPO / p for p in ('src/lifeplan.py', 'src/p2a.py', 'src/p2a_excel.py', 'src/p2a_s1.py',
                                  'tools/s1_tiny_diagnostic.py', 'tools/s1_compatibility.py',
                                  'tools/p2a_s1_cases.py', 'tools/verify_p2a_s1.py', 'tests/p2a_cases.py')]
    hashes = {str(p.relative_to(REPO)): sha(p) for p in paths}
    digest = hashlib.sha256(json.dumps(hashes, sort_keys=True).encode()).hexdigest()
    try:
        commit = subprocess.run(['git', 'rev-parse', 'HEAD'], cwd=REPO, capture_output=True, text=True)
        commit_id = commit.stdout.strip() if commit.returncode == 0 else None
    except OSError:
        commit_id = None
    from importlib.metadata import version, PackageNotFoundError
    dependencies = {}
    for name in ('openpyxl', 'pytest', 'psutil', 'pywin32'):
        try:
            dependencies[name] = version(name)
        except PackageNotFoundError:
            dependencies[name] = 'not installed'
    return dict(source_sha256=digest, files=hashes, commit=commit_id, dependencies=dependencies)


def execute(mode, config, output=None, *, check=preflight, monitor=supervise,
            units=unit_tests, residual=residual_check, approval_id=None):
    run_id = str(uuid.uuid4())
    root = create_run(output or (REPO.parent / 'lifeplan-excel2024-runs' / run_id))
    value = dict(schema=SCHEMA, run_id=run_id, started_utc=utc(), status='NOT_RUN',
                 mode=mode, excel_recalculation_verified=False,
                 stages=[dict(case_id=s, status='NOT_RUN') for s in STAGES],
                 warning_check='初回実機で所有者による警告/修復/操作性確認が必要',
                 force_termination=False, scope='automatic_numeric_gate_only', product_acceptance=False,
                 python=sys.version, python_bits=64 if sys.maxsize > 2**32 else 32,
                 os=platform.platform(), source=source_identity())
    active_stage = None
    if mode == 'e1':
        value.update(scope='tiny3_e1_diagnostic_only', diagnostic_only=True, gate2_acceptance=False,
                     diagnostic=E1_DIAGNOSTIC, protocol=E1_PROTOCOL,
                     diagnostic_recalculation_verified=False)
    if mode == 'functions24-false':
        value.update(scope='functions24_false_diagnostic_only', diagnostic_only=True, gate2_acceptance=False,
                     diagnostic=F24_DIAGNOSTIC, protocol=F24_PROTOCOL,
                     diagnostic_recalculation_verified=False)
    if mode == 'functions24-preopen-force0':
        value.update(scope='functions24_preopen_force0_diagnostic_only', diagnostic_only=True,
                     gate2_acceptance=False, diagnostic=A_DIAGNOSTIC, protocol=A_PROTOCOL,
                     approval_id=approval_id, diagnostic_recalculation_verified=False)
    if mode == 'doctor':
        value['doctor_aggregation'] = dict(phase='preflight', status='NOT_RUN', checks={})
    try:
        reason = check()
        if reason:
            value.update(status='BLOCKED', reason=reason)
            if mode in ('smoke', 'e1'):
                value['stages'][0].update(status='BLOCKED', reason=reason)
            if mode in ('functions24-false', 'functions24-preopen-force0'):
                value['stages'][1].update(status='BLOCKED', reason=reason)
            return value, root
        config = Path(config).resolve()
        if mode == 'doctor':
            aggregation = value['doctor_aggregation']
            aggregation.update(phase='profile_precondition', profile_path=str(config))
        if config.is_relative_to(REPO):
            raise ValueError('環境profileもリポジトリ外へ保存してください')
        if approval_id is not None and mode != 'functions24-preopen-force0':
            raise ValueError('承認IDはA専用modeだけに使用できます')
        if mode == 'doctor':
            if config.exists():
                raise FileExistsError('既存profileは上書きしません。新しいファイル名を指定してください')
            directory = root / 'doctor'
            directory.mkdir()
            request = dict(run_id=run_id, case_id='doctor', source_sha256=value['source']['source_sha256'])
            aggregation['phase'] = 'worker'
            result = monitor(directory, request)
            value['doctor'] = result
            aggregation.update(phase='worker_status', worker_status=result.get('status'))
            if result.get('status') == 'CONFIGURED':
                aggregation['phase'] = 'worker_evidence'
                if (result.get('schema') != SCHEMA or result.get('run_id') != run_id or result.get('case_id') != 'doctor'
                        or result.get('request_sha256') != sha(directory / 'request.json')):
                    raise ValueError('別実行のdoctor記録です')
                aggregation['checks']['worker_evidence'] = 'PASS'
                aggregation['phase'] = 'product_and_profile'
                profile = result['profile']
                if profile['facts'] != result['identity']:
                    raise ValueError('UI確認とCOM接続先の結び付けが不一致です')
                check_profile(profile, profile['facts'])
                aggregation['checks']['product_and_profile'] = 'PASS'
                aggregation['phase'] = 'worker_cleanup'
                if (result.get('first_error') or result.get('cleanup_errors') or result.get('progress_log_errors')
                        or result.get('worker_still_running') or result.get('cleanup_pending')):
                    raise ValueError('doctor終了処理に失敗')
                aggregation['checks']['worker_cleanup'] = 'PASS'
                aggregation['phase'] = 'residual'
                identity = result['identity']
                observation = dict(pid=identity['pid'], created=identity['created'], path=identity['path'],
                                   started_utc=utc())
                aggregation['residual_check'] = observation
                started = time.monotonic()
                try:
                    remains = residual(result)
                finally:
                    observation['elapsed_seconds'] = time.monotonic() - started
                if type(remains) is not bool:
                    raise ValueError('専用Excel残留確認が真偽値を返しませんでした')
                observation['same_owned_process_remains'] = remains
                aggregation['checks']['residual'] = 'FAIL' if remains else 'PASS'
                if remains:
                    value.update(status='FAIL', reason=f"通常Quit後も専用Excel PID={identity['pid']}が残留しています",
                                 reason_code='OWNED_EXCEL_REMAINS')
                    aggregation.update(status='FAIL', failed_at='residual')
                    return value, root
                aggregation['phase'] = 'profile_save'
                config.parent.mkdir(parents=True, exist_ok=True)
                # Exclusive create prevents races and accidental overwrite.
                with config.open('x', encoding='utf-8') as f:
                    json.dump(profile, f, ensure_ascii=False, indent=2)
                aggregation['checks']['profile_save'] = 'PASS'
                aggregation.update(phase='complete', status='CONFIGURED')
                value.update(status='CONFIGURED', config=str(config))
            else:
                status = result.get('status')
                value.update(status=status if status in ('FAIL', 'BLOCKED', 'TIMEOUT', 'INTERRUPTED') else 'FAIL',
                             reason='doctor子プロセスが正常完了していません: ' + str(status),
                             reason_code='DOCTOR_WORKER_NOT_CONFIGURED')
                aggregation.update(status=value['status'], failed_at='worker_status')
            return value, root
        if mode == 'e1':
            from .e1 import execute as execute_e1
            return execute_e1(value, root, config, monitor=monitor, units=units, residual=residual)
        if mode == 'functions24-false':
            from .functions24 import execute as execute_functions24
            return execute_functions24(value, root, config, monitor=monitor, units=units, residual=residual)
        if mode == 'functions24-preopen-force0':
            from .preopen_input import validate_approval_id
            from .functions24 import execute as execute_functions24
            validate_approval_id(approval_id)
            return execute_functions24(value, root, config, monitor=monitor, units=units,
                                       residual=residual, preopen_approval=approval_id)
        if mode != 'smoke':
            raise ValueError('承認対象はdoctor/smokeのみ')
        if not config.is_file():
            value.update(status='BLOCKED', reason='最初にdoctorで同一専用Excelの製品/認証を確認してください')
            value['stages'][0].update(status='BLOCKED', reason=value['reason'])
            return value, root
        profile = read_json(config)
        if profile.get('schema') != SCHEMA:
            raise ValueError('旧版/別profileです')
        value['unit_tests'] = units(root)
        if value['unit_tests']['returncode'] != 0:
            value.update(status='FAIL', reason='Pythonテスト不合格。Excelへ進みません')
            return value, root
        for stage in value['stages']:
            active_stage = stage
            directory = root / stage['case_id']
            directory.mkdir()
            manifest = prepare(stage['case_id'], directory, run_id)
            request = dict(run_id=run_id, case_id=stage['case_id'], profile=profile,
                           manifest_sha256=sha(directory / 'case.json'), source_sha256=value['source']['source_sha256'])
            result = monitor(directory, request)
            stage['runtime'] = result
            if result.get('status') != 'SAVED_REOPENED':
                stage['status'] = result.get('status', 'FAIL')
                break
            try:
                validate_runtime(result, directory, request, profile)
                verified = verify_saved(directory, manifest)
                if verified['saved_sha256'] != result['saved_sha256']:
                    raise ValueError('保存後に結果が変更されています')
                if residual(result):
                    raise ValueError('通常Quit後も専用Excelが残留しています。強制終了しません')
            except BaseException as error:
                stage.update(status='FAIL', first_error=error_info(error))
                break
            stage.update(status='PASS', verification=verified,
                         product_basis='同一専用PIDの初回UI確認＋現在の製品ID/プロセス指紋一致')
            report(root, value)
        else:
            value.update(status='PASS', excel_recalculation_verified=True)
            return value, root
        value['status'] = next((s['status'] for s in value['stages'] if s['status'] not in ('PASS', 'NOT_RUN')), 'FAIL')
        return value, root
    except BaseException as error:
        value.update(status=error.status if isinstance(error, GateError) else 'INTERRUPTED' if isinstance(error, KeyboardInterrupt) else 'FAIL',
                     first_error=error_info(error))
        if mode == 'doctor':
            aggregation = value['doctor_aggregation']
            aggregation.update(status=value['status'], failed_at=aggregation['phase'])
            aggregation['checks'][aggregation['phase']] = 'FAIL'
            value.update(reason=str(error), reason_code='DOCTOR_PARENT_ERROR')
        if active_stage is not None and active_stage['status'] == 'NOT_RUN':
            active_stage.update(status=value['status'], reason='この段階で処理中断')
        if mode in ('functions24-false', 'functions24-preopen-force0') and value['stages'][1]['status'] == 'NOT_RUN':
            value['stages'][1].update(status=value['status'], reason='専用診断の事前条件で処理中断')
        return value, root
    finally:
        if mode == 'doctor' and value['doctor_aggregation']['status'] == 'NOT_RUN':
            value['doctor_aggregation'].update(status=value['status'], failed_at=value['doctor_aggregation']['phase'])
        value['completed_utc'] = utc()
        value['exit_code'] = exit_code_for(value['status'])
        report(root, value)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('mode', choices=('doctor', 'smoke', 'e1', 'functions24-false', 'functions24-preopen-force0'))
    parser.add_argument('--approval-id', help='A実機実行の別承認ID。今回Windows実行は未承認')
    parser.add_argument('--config', type=Path, default=DEFAULT_PROFILE)
    parser.add_argument('--output', type=Path, help='リポジトリ外の新規フォルダー（省略推奨）')
    args = parser.parse_args()
    if args.approval_id is not None or args.mode == 'functions24-preopen-force0':
        value, root = execute(args.mode, args.config, args.output, approval_id=args.approval_id)
    else:
        value, root = execute(args.mode, args.config, args.output)
    doctor = value.get('doctor', {})
    summary = dict(status=value['status'], stages=value['stages'],
                          doctor_status=doctor.get('status'), identity_error=doctor.get('first_error'),
                          first_error=value.get('first_error'), reason=value.get('reason'),
                          doctor_aggregation=value.get('doctor_aggregation'),
                   exit_code=value['exit_code'], report=str(root / 'run.json'))
    if args.mode in ('e1', 'functions24-false', 'functions24-preopen-force0'):
        summary.update(diagnostic_only=True, gate2_acceptance=False,
                       diagnostic_recalculation_verified=value['diagnostic_recalculation_verified'])
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return value['exit_code']


if __name__ == '__main__':
    raise SystemExit(main())
