"""Approved functions24-only diagnostic; no Gate 2/product acceptance."""
import ntpath
import os
from pathlib import Path
import sys

from . import SCHEMA, F24_DIAGNOSTIC, F24_PROTOCOL, A_DIAGNOSTIC, A_PROTOCOL
from .cases import EXPECTED24, compare_probes
from .common import GateError, error_info, read_json, sha, utc


def protocol_for(diagnostic):
    if diagnostic not in (F24_DIAGNOSTIC, A_DIAGNOSTIC):
        raise ValueError('未承認のfunctions24診断です')
    return A_PROTOCOL if diagnostic == A_DIAGNOSTIC else F24_PROTOCOL


def validate_manifest(manifest, *, diagnostic=F24_DIAGNOSTIC):
    protocol_for(diagnostic)
    if (manifest.get('case_id') != 'functions24' or manifest.get('file') != '01_function_gate.xlsx'
            or manifest.get('formula_count') != 24
            or manifest.get('probes') != [['確認', c, v] for c, v in EXPECTED24.items()]
            or manifest.get('challenges') != [['確認', 'B8', 0]]
            or manifest.get('edits') != [['確認', 'B8', 100]]):
        raise ValueError('functions24専用の原本・24固定値・challenge/restoreに限定されています')
    from .preopen_verify import ORIGINAL_SHA
    if diagnostic == F24_DIAGNOSTIC and (manifest.get('control') is not None or manifest.get('sha256') != ORIGINAL_SHA):
        raise ValueError('旧functions24へA対照入力を持ち込めません')
    if diagnostic == A_DIAGNOSTIC and (not manifest.get('control') or manifest.get('sha256') == ORIGINAL_SHA):
        raise ValueError('A専用対照入力とcontrolが必要です')


def challenge_manifest(manifest):
    return dict(manifest, probes=[[s, c, 1 if c == 'L8' else '不一致' if c == 'M8' else v]
                                 for s, c, v in manifest['probes']])


def process_identity(pid):
    """Observe only this supervisor/launcher/worker, never enumerate other trees."""
    import psutil
    process = psutil.Process(pid)
    return dict(pid=process.pid, ppid=process.ppid(), created=process.create_time(), path=process.exe())


def worker_process():
    return dict(**process_identity(os.getpid()), sys_executable=sys.executable)


def worker_exit_observation(observed, *, process=None):
    """Read-only leaf check after the launcher exits; no new wait/termination."""
    if process is None:
        import psutil
        process, missing = psutil.Process, psutil.NoSuchProcess
    else:
        missing = ProcessLookupError
    try:
        child = process(observed['pid'])
        if child.create_time() != observed['created']:
            return dict(outcome='PID_REUSED')
        if ntpath.normcase(child.exe()) != ntpath.normcase(observed['path']):
            raise ValueError('Python worker PIDの実行パスが変化しています')
    except missing:
        return dict(outcome='PROCESS_EXITED')
    return dict(outcome='WORKER_STILL_RUNNING')


def validate_processes(result):
    proof = result.get('python_processes', {})
    parent, launcher, observed = (proof.get(k, {}) for k in ('supervisor', 'launcher', 'worker'))
    child = result.get('python_process', {})
    for record in (parent, launcher, observed, child):
        if (type(record.get('pid')) is not int or record['pid'] <= 0
                or type(record.get('ppid')) is not int or not record.get('path')
                or not isinstance(record.get('created'), (int, float))):
            raise ValueError('Python監視PID対応の観測不足')
    if (child != observed or not child.get('sys_executable')
            or launcher['pid'] != result.get('worker_pid') or launcher['ppid'] != parent['pid']):
        raise ValueError('Python監視PIDとworker証拠が対応しません')
    if proof.get('worker_exit', {}).get('outcome') not in ('PROCESS_EXITED', 'PID_REUSED'):
        raise ValueError('Python workerの終了を確認できません')
    if child['pid'] == launcher['pid']:
        if any(child[k] != launcher[k] for k in ('created', 'path', 'ppid')):
            raise ValueError('同一Python PIDの指紋が一致しません')
        return 'DIRECT_WORKER'
    if child['ppid'] != launcher['pid'] or child['created'] < launcher['created']:
        raise ValueError('Pythonランチャーとworkerの直接親子関係を確認できません')
    return 'LAUNCHER_CHILD_WORKER'


def context(excel, book, manifest, log, phase, *, diagnostic=F24_DIAGNOSTIC):
    protocol_for(diagnostic)
    from .worker import snapshot
    prefix = 'functions24.' + phase
    flag = log.call(prefix + '.force_full', lambda: book.ForceFullCalculation)
    if type(flag) is not bool:
        raise GateError('FAIL', 'functions24のForceFullCalculationが真偽値ではありません')

    def inventory():
        collection = excel.Workbooks
        items = []
        for index in range(1, int(collection.Count) + 1):
            item = collection.Item(index)
            items.append(dict(name=str(item.Name), path=str(item.FullName), read_only=bool(item.ReadOnly)))
            item = None
        collection = None
        return items

    books = log.call(prefix + '.workbooks', inventory)
    log.emit('functions24.inventory', phase=phase, force_full=flag, workbooks=books)
    expected_path = log.directory / ('results' if phase == 'reopened' else 'work') / manifest['file']
    if (len(books) != 1 or ntpath.normcase(books[0]['path']) != ntpath.normcase(str(expected_path.resolve()))
            or books[0]['name'] != manifest['file'] or books[0]['read_only'] != (phase == 'reopened')):
        raise GateError('BLOCKED', 'functions24専用インスタンスのブック/パス/ReadOnlyが想定外です',
                        details=dict(phase=phase, expected_path=str(expected_path.resolve()), workbooks=books))
    if (diagnostic == A_DIAGNOSTIC or phase not in ('opened', 'challenge.before_flag')) and flag is not False:
        raise GateError('FAIL', 'functions24のForceFullCalculation=Falseが維持されていません')
    sample = snapshot(excel, book, manifest, log, prefix, required=phase in ('before_save', 'reopened'))
    if diagnostic == A_DIAGNOSTIC and phase == 'opened' and sample['mode'] != -4105:
        raise GateError('BLOCKED', 'A対照の初期計算modeが自動ではありません')
    if sample['state_before'] not in (0, 1, 2) or sample['state'] not in (0, 1, 2):
        raise GateError('FAIL', 'functions24の不明なCalculationState')
    if phase.endswith(('.before_flag', '.after_flag', '.after_full_rebuild')) and sample['mode'] != -4135:
        raise GateError('FAIL', 'functions24の手動計算設定の不一致')
    item = dict(phase=phase, utc=utc(), force_full=flag, workbooks=books, sample=sample)
    log.emit('functions24.context', evidence=item)
    return item


def recalculate(excel, book, log, label, manifest, evidence, *, diagnostic=F24_DIAGNOSTIC):
    """Same manual-edit/full-rebuild/auto/wait order; only the flag is False."""
    from .worker import await_done
    protocol = protocol_for(diagnostic)
    observe = lambda phase: context(excel, book, manifest, log, phase, diagnostic=diagnostic)
    evidence.append(observe(label + '.before_flag'))
    log.call(label + '.force_full', lambda: setattr(book, 'ForceFullCalculation', False))
    flag = log.call(label + '.force_full_read', lambda: book.ForceFullCalculation)
    log.emit('calculation.protocol', label=label, force_full_observed=flag, protocol=protocol)
    if flag is not False:
        raise GateError('FAIL', 'functions24のForceFullCalculation=Falseを読み戻せませんでした')
    evidence.append(observe(label + '.after_flag'))
    log.call(label + '.full_rebuild', excel.CalculateFullRebuild, 180)
    evidence.append(observe(label + '.after_full_rebuild'))
    log.call(label + '.automatic_restore', lambda: setattr(excel, 'Calculation', -4105))
    mode = log.call(label + '.automatic_read', lambda: int(excel.Calculation))
    log.emit('calculation.automatic_observed', label=label, mode=mode)
    if mode != -4105:
        raise GateError('FAIL', '自動計算への復帰不一致')
    evidence.append(observe(label + '.after_automatic'))
    await_done(excel, log)


def validate_evidence(result, manifest, directory, *, diagnostic=F24_DIAGNOSTIC):
    validate_manifest(manifest, diagnostic=diagnostic)
    record = result.get('functions24', {})
    if record.get('diagnostic') != diagnostic or record.get('protocol') != protocol_for(diagnostic):
        raise ValueError('functions24専用プロトコルの実測証拠がありません')
    phases = ['opened'] + [label + '.' + p for label in ('challenge', 'final')
                           for p in ('before_flag', 'after_flag', 'after_full_rebuild', 'after_automatic')] + ['before_save', 'reopened']
    evidence = record.get('evidence', [])
    if [e.get('phase') for e in evidence] != phases:
        raise ValueError('functions24観測段階の不足/順序不一致')
    for item in evidence:
        phase = item['phase']
        books = item.get('workbooks', [])
        if not item.get('utc') or type(item.get('force_full')) is not bool or len(books) != 1:
            raise ValueError('functions24の時刻/フラグ/ブック一覧証拠不足')
        expected_path = Path(directory) / ('results' if phase == 'reopened' else 'work') / manifest['file']
        if (ntpath.normcase(books[0].get('path', '')) != ntpath.normcase(str(expected_path.resolve()))
                or books[0].get('name') != manifest['file'] or books[0].get('read_only') is not (phase == 'reopened')):
            raise ValueError('functions24のブック/パス/ReadOnly証拠不一致')
        if (diagnostic == A_DIAGNOSTIC or phase not in ('opened', 'challenge.before_flag')) and item['force_full'] is not False:
            raise ValueError('functions24のFalseが両フェーズで維持されていません')
        sample = item['sample']
        if diagnostic == A_DIAGNOSTIC and phase == 'opened' and sample.get('mode') != -4105:
            raise ValueError('A初期modeの証拠不一致')
        if sample.get('state_before') not in (0, 1, 2) or sample.get('state') not in (0, 1, 2):
            raise ValueError('functions24の計算状態証拠不明')
        if sample.get('phase') != 'functions24.' + phase:
            raise ValueError('functions24観測とsnapshot段階が不一致です')
        if any(phase.endswith('.' + p) for p in ('before_flag', 'after_flag', 'after_full_rebuild')):
            if sample.get('mode') != -4135:
                raise ValueError('functions24のFullRebuild前後が手動計算ではありません')
        if phase.endswith('.after_automatic') or phase in ('before_save', 'reopened'):
            if sample.get('mode') != -4105:
                raise ValueError('functions24の自動復帰証拠不一致')
        if phase in ('before_save', 'reopened') and (sample['state_before'] != 0 or sample['state'] != 0 or sample['differences']):
            raise ValueError('functions24の保存/再オープンの状態0・値証拠不一致')
        wanted = challenge_manifest(manifest) if phase.startswith('challenge.') else manifest
        if set(sample.get('values', {})) != {s + '!' + c for s, c, _ in wanted['probes']}:
            raise ValueError('functions24の全24セル観測が不足しています')
        if sample.get('differences') != compare_probes(wanted, lambda s, c: sample['values'][s + '!' + c]):
            raise ValueError('functions24の観測値と差異記録が不一致です')
    if (result.get('owned_pid') != result['identity']['pid'] or result.get('force_termination') is not False):
        raise ValueError('functions24の専用Excel所有/正常終了証拠不足')
    result['python_process_relation'] = validate_processes(result)


def execute(value, root, config, *, monitor, units, residual, preopen_approval=None):
    from . import run
    stage = next(s for s in value['stages'] if s['case_id'] == 'functions24')
    result, ready = None, False
    diagnostic = F24_DIAGNOSTIC if preopen_approval is None else A_DIAGNOSTIC
    try:
        if diagnostic == A_DIAGNOSTIC:
            from . import preopen_input
            if value.get('diagnostic') != A_DIAGNOSTIC or value.get('mode') != 'functions24-preopen-force0':
                raise ValueError('A入口以外からのprepareは禁止')
        if not config.is_file():
            raise GateError('BLOCKED', '最初にdoctorで同一専用Excelの製品/認証を確認してください')
        profile = read_json(config)
        if profile.get('schema') != SCHEMA:
            raise ValueError('旧版/別profileです')
        value['unit_tests'] = units(root)
        if value['unit_tests']['returncode'] != 0:
            raise ValueError('Pythonテスト不合格。Excelへ進みません')
        directory = root / 'functions24'
        directory.mkdir()
        if diagnostic == A_DIAGNOSTIC:
            manifest = preopen_input.prepare(directory, value['run_id'], preopen_approval, config.parent,
                                             value['source']['source_sha256'])
        else:
            manifest = run.prepare('functions24', directory, value['run_id'])
        validate_manifest(manifest, diagnostic=diagnostic)
        request = dict(run_id=value['run_id'], case_id='functions24', diagnostic=diagnostic,
                       profile=profile, manifest_sha256=sha(directory / 'case.json'),
                       source_sha256=value['source']['source_sha256'])
        if diagnostic == A_DIAGNOSTIC:
            request.update(protocol=A_PROTOCOL, control=manifest['control'])
            value['control'] = manifest['control']
            preopen_input.validate_bundle(directory, manifest, request)
        result = monitor(directory, request)
        stage['runtime'] = result
        if result.get('status') != 'SAVED_REOPENED':
            status = result.get('status')
            stage['status'] = status if status in ('FAIL', 'BLOCKED', 'TIMEOUT', 'INTERRUPTED') else 'FAIL'
            value['status'] = stage['status']
        else:
            if result.get('worker_still_running') or result.get('cleanup_pending'):
                raise ValueError('functions24子プロセス終了処理が未完了です')
            run.validate_runtime(result, directory, request, profile)
            validate_evidence(result, manifest, directory, diagnostic=diagnostic)
            if diagnostic == A_DIAGNOSTIC:
                verified_control = preopen_input.validate_bundle(directory, manifest, request)
                if result.get('preopen_control') != verified_control:
                    raise ValueError('A workerと親の入力再検査証拠が不一致')
                stage['control_verification'] = verified_control
            verified = run.verify_saved(directory, manifest)
            if verified['saved_sha256'] != result['saved_sha256']:
                raise ValueError('保存後にfunctions24結果が変更されています')
            stage['verification'] = verified
            ready = True
    except BaseException as error:
        status = error.status if isinstance(error, GateError) else 'INTERRUPTED' if isinstance(error, KeyboardInterrupt) else 'FAIL'
        stage.update(status=status, first_error=error_info(error))
        value.update(status=status, first_error=error_info(error))
    if result and result.get('identity') and not result.get('worker_still_running'):
        try:
            remains = residual(result)
            if type(remains) is not bool:
                raise ValueError('functions24残留確認が真偽値を返しませんでした')
            stage['residual'] = dict(same_owned_process_remains=remains)
            if remains:
                raise ValueError('通常Quit後もfunctions24専用Excelが残留しています。強制終了しません')
        except BaseException as error:
            stage['residual_error'] = error_info(error)
            if ready:
                status = 'INTERRUPTED' if isinstance(error, KeyboardInterrupt) else 'FAIL'
                stage.update(status=status, first_error=error_info(error))
                value.update(status=status, first_error=error_info(error))
            ready = False
    elif ready:
        stage.update(status='FAIL', first_error=dict(message='functions24専用PID終了の証拠不足'))
        value['status'], ready = 'FAIL', False
    if ready:
        stage.update(status='PASS', product_basis='doctor profile + owned PID + functions24-only evidence')
        value.update(status='DIAGNOSTIC_PASS', diagnostic_recalculation_verified=True)
    return value, root
