"""Explicit tiny3-only E1 coordinator; never a Gate 2 acceptance result."""
from . import SCHEMA, E1_DIAGNOSTIC, E1_PROTOCOL
from .common import GateError, error_info, read_json, sha


def validate_evidence(result, manifest):
    from .cases import compare_probes
    record = result.get('e1', {})
    if record.get('diagnostic') != E1_DIAGNOSTIC or record.get('protocol') != E1_PROTOCOL:
        raise ValueError('E1専用プロトコルの実測証拠がありません')
    evidence = record.get('evidence', [])
    phases = ['opened', 'before_flag', 'after_flag', 'after_full_rebuild', 'after_automatic', 'reopened']
    if [e.get('phase') for e in evidence] != phases:
        raise ValueError('E1観測段階の不足/順序不一致')
    for item in evidence:
        if not item.get('utc') or type(item.get('force_full')) is not bool or len(item.get('workbooks', [])) != 1:
            raise ValueError('E1フラグ/時刻/ブック一覧の証拠不足')
        if item['phase'] not in ('opened', 'before_flag') and item['force_full'] is not False:
            raise ValueError('E1のForceFullCalculation=Falseが維持されていません')
        sample = item['sample']
        if sample.get('phase') != 'e1.' + item['phase']:
            raise ValueError('E1観測段階とsnapshotの対応が一致しません')
        if item['phase'] in ('before_flag', 'after_flag', 'after_full_rebuild') and sample['mode'] != -4135:
            raise ValueError('E1のFullRebuild前後が手動計算ではありません')
        if item['phase'] in ('after_automatic', 'reopened'):
            if sample['mode'] != -4105:
                raise ValueError('E1の自動計算復帰が確認できません')
        # Before waiting can legitimately be pending, even with updated values.
        # Required state0/value evidence is checked by validate_runtime instead.
        if set(sample.get('values', {})) != {s + '!' + c for s, c, _ in manifest['probes']}:
            raise ValueError('E1の3セル観測が不足しています')
        if sample.get('differences') != compare_probes(manifest, lambda s, c: sample['values'][s + '!' + c]):
            raise ValueError('E1観測値と差異記録が一致しません')


def execute(value, root, config, *, monitor, units, residual):
    from . import run
    stage = value['stages'][0]
    result = None
    ready = False
    try:
        if not config.is_file():
            raise GateError('BLOCKED', '最初にdoctorで同一専用Excelの製品/認証を確認してください')
        profile = read_json(config)
        if profile.get('schema') != SCHEMA:
            raise ValueError('旧版/別profileです')
        value['unit_tests'] = units(root)
        if value['unit_tests']['returncode'] != 0:
            raise ValueError('Pythonテスト不合格。Excelへ進みません')
        directory = root / 'tiny3'
        directory.mkdir()
        manifest = run.prepare('tiny3', directory, value['run_id'])
        request = dict(run_id=value['run_id'], case_id='tiny3', diagnostic=E1_DIAGNOSTIC,
                       profile=profile, manifest_sha256=sha(directory / 'case.json'),
                       source_sha256=value['source']['source_sha256'])
        result = monitor(directory, request)
        stage['runtime'] = result
        if result.get('status') != 'SAVED_REOPENED':
            status = result.get('status')
            stage['status'] = status if status in ('FAIL', 'BLOCKED', 'TIMEOUT', 'INTERRUPTED') else 'FAIL'
            value['status'] = stage['status']
        else:
            if result.get('worker_still_running') or result.get('cleanup_pending'):
                raise ValueError('E1子プロセスの終了処理が未完了です')
            run.validate_runtime(result, directory, request, profile)
            validate_evidence(result, manifest)
            verified = run.verify_saved(directory, manifest)
            if verified['saved_sha256'] != result['saved_sha256']:
                raise ValueError('保存後にE1結果が変更されています')
            stage['verification'] = verified
            ready = True
    except BaseException as error:
        status = error.status if isinstance(error, GateError) else 'INTERRUPTED' if isinstance(error, KeyboardInterrupt) else 'FAIL'
        stage.update(status=status, first_error=error_info(error))
        value.update(status=status, first_error=error_info(error))
    # Observe normal process exit on failed diagnostics too, without replacing
    # the original timeout/COM/cleanup error. Never touch a still-running worker.
    if result and result.get('identity') and not result.get('worker_still_running'):
        try:
            remains = residual(result)
            if type(remains) is not bool:
                raise ValueError('E1残留確認が真偽値を返しませんでした')
            stage['residual'] = dict(same_owned_process_remains=remains)
            if remains:
                raise ValueError('通常Quit後もE1専用Excelが残留しています。強制終了しません')
        except BaseException as error:
            stage['residual_error'] = error_info(error)
            if ready:
                stage.update(status='FAIL', first_error=error_info(error))
                value.update(status='FAIL', first_error=error_info(error))
            ready = False
    elif ready:
        stage.update(status='FAIL', first_error=dict(message='E1専用PID終了の証拠不足'))
        value['status'] = 'FAIL'
        ready = False
    if ready:
        stage.update(status='PASS', product_basis='doctor profile + owned PID + E1-only evidence')
        value.update(status='DIAGNOSTIC_PASS', diagnostic_recalculation_verified=True)
    return value, root
