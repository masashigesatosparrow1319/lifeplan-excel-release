"""Child process owns COM. A late return obeys STOP; no kill/terminate paths."""
import argparse
from datetime import date
from pathlib import Path
import sys
import time
import traceback

from . import SCHEMA, E1_DIAGNOSTIC, E1_PROTOCOL, F24_DIAGNOSTIC, A_DIAGNOSTIC
from .cases import compare_probes
from .common import GateError, Progress, error_info, read_json, sha, utc, write_json
from .session import Session, LIMITATION, check_product


def pump_messages():
    if sys.platform == 'win32':
        import pythoncom
        if pythoncom.PumpWaitingMessages():
            raise GateError('INTERRUPTED', 'COM STAに終了メッセージを受信')


def await_done(excel, log):
    started = time.monotonic()
    deadline = started + 180
    observations = 0
    last_state = None
    counts = {}
    log.emit('calculation.wait_started', timeout_seconds=180)

    def expired():
        details = dict(elapsed_seconds=time.monotonic() - started, timeout_seconds=180,
                       last_state=last_state, observations=observations, state_counts=counts)
        log.emit('calculation.wait_timeout', **details)
        raise GateError('TIMEOUT', 'CalculationState待機180秒超過', details=details)

    while True:
        log.check_stop()
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            expired()
        # A COM read may block. Its parent-visible deadline cannot exceed
        # the remaining wait budget; synchronous COM is never force-aborted.
        state = log.call('calculation_state', lambda: int(excel.CalculationState), min(60, remaining))
        last_state = state
        observations += 1
        counts[str(state)] = counts.get(str(state), 0) + 1
        elapsed = time.monotonic() - started
        log.emit('calculation.observed', state=state,
                 state_name={0: 'xlDone', 1: 'xlCalculating', 2: 'xlPending'}.get(state, 'UNKNOWN'),
                 elapsed_seconds=elapsed, observation=observations)
        if time.monotonic() >= deadline:
            expired()
        if state == 0:
            log.emit('calculation.wait_done', elapsed_seconds=time.monotonic() - started,
                     observations=observations, state_counts=counts)
            return
        if state not in (1, 2):
            raise GateError('FAIL', '不明なCalculationState: ' + str(state))
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            expired()
        log.call('com.pump_messages', pump_messages, min(60, remaining))
        time.sleep(min(.1, max(0, deadline - time.monotonic())))


def set_edits(book, edits):
    for sheet, cell, value in edits:
        target = book.Worksheets.Item(sheet).Range(cell)
        if value is None:
            target.ClearContents()
        elif isinstance(value, dict) and set(value) == {'date'}:
            # Excel fixtures use the 1900 date system. No NumberFormat mutation.
            if book.Date1904:
                raise ValueError('1904日付系のfixtureは対象外です')
            day = date.fromisoformat(value['date'])
            target.Value2 = (day - date(1899, 12, 30)).days
        elif isinstance(value, (int, float, str)) and not isinstance(value, bool):
            target.Value2 = value
        else:
            raise ValueError('未承認の編集型')
        target = None


def snapshot(excel, book, manifest, log, name, *, required=True):
    state_before = log.call(name + '.state_before', lambda: int(excel.CalculationState))
    values = log.call(name + '.values', lambda: {
        sheet + '!' + cell: book.Worksheets.Item(sheet).Range(cell).Value2
        for sheet, cell, _ in manifest['probes']})
    mode = log.call(name + '.mode', lambda: int(excel.Calculation))
    state = log.call(name + '.state_after', lambda: int(excel.CalculationState))
    differences = compare_probes(manifest, lambda s, c: values[s + '!' + c])
    sample = dict(phase=name, state_before=state_before, state=state, mode=mode,
                  values=values, differences=differences)
    log.emit('snapshot', sample=sample)
    if required and (state_before != 0 or state != 0 or mode != -4105 or differences):
        raise GateError('FAIL', '状態0・自動計算・固定値の条件未達: ' + str(sample))
    return sample


def e1_context(excel, book, manifest, log, phase):
    """Read only the owned application's objects; retain primitives only."""
    flag = log.call('e1.' + phase + '.force_full', lambda: book.ForceFullCalculation)
    if type(flag) is not bool:
        raise GateError('FAIL', 'E1のForceFullCalculationが真偽値ではありません')
    def inventory():
        collection = excel.Workbooks
        items = []
        for index in range(1, int(collection.Count) + 1):
            item = collection.Item(index)
            items.append(dict(name=str(item.Name), path=str(item.FullName), read_only=bool(item.ReadOnly)))
            item = None
        collection = None
        return items
    books = log.call('e1.' + phase + '.workbooks', inventory)
    sample = snapshot(excel, book, manifest, log, 'e1.' + phase, required=False)
    evidence = dict(phase=phase, utc=utc(), force_full=flag, workbooks=books, sample=sample)
    log.emit('e1.context', evidence=evidence)
    if len(books) != 1:
        raise GateError('BLOCKED', 'E1専用インスタンスに予期しないブックがあります')
    if phase not in ('opened', 'before_flag') and flag is not False:
        raise GateError('FAIL', 'E1観測でForceFullCalculation=Falseが維持されていません')
    return evidence


def recalculate(excel, book, log, label, *, manifest=None, diagnostic=None, evidence=None):
    if diagnostic is not None and diagnostic != E1_DIAGNOSTIC:
        raise ValueError('未承認の診断です')
    is_e1 = diagnostic == E1_DIAGNOSTIC
    if is_e1 and (manifest is None or evidence is None):
        raise ValueError('E1観測のmanifestと証拠リストが必要です')
    if is_e1:
        evidence.append(e1_context(excel, book, manifest, log, 'before_flag'))
    log.call(label + '.force_full', lambda: setattr(book, 'ForceFullCalculation', not is_e1))
    force_full = log.call(label + '.force_full_read',
                          lambda: book.ForceFullCalculation if is_e1 else bool(book.ForceFullCalculation))
    log.emit('calculation.protocol', label=label, force_full_observed=force_full,
             protocol=E1_PROTOCOL if is_e1 else 'full_rebuild_auto_wait_v1')
    if is_e1:
        if force_full is not False:
            raise GateError('FAIL', 'E1のForceFullCalculation=Falseを読み戻せませんでした')
        evidence.append(e1_context(excel, book, manifest, log, 'after_flag'))
    log.call(label + '.full_rebuild', excel.CalculateFullRebuild, 180)
    if is_e1:
        evidence.append(e1_context(excel, book, manifest, log, 'after_full_rebuild'))
    log.call(label + '.automatic_restore', lambda: setattr(excel, 'Calculation', -4105))
    mode = log.call(label + '.automatic_read', lambda: int(excel.Calculation))
    log.emit('calculation.automatic_observed', label=label, mode=mode)
    if mode != -4105:
        raise GateError('FAIL', '自動計算への復帰不一致')
    if manifest is not None:
        # Observation only: correct values with xlPending are NEVER success.
        snapshot(excel, book, manifest, log, label + '.before_wait', required=False)
    if is_e1:
        evidence.append(e1_context(excel, book, manifest, log, 'after_automatic'))
    await_done(excel, log)


def execute_case(session, manifest, directory, log, books, *, diagnostic=None, evidence=None):
    is_f24 = diagnostic in (F24_DIAGNOSTIC, A_DIAGNOSTIC)
    if is_f24:
        from . import functions24
        functions24.validate_manifest(manifest, diagnostic=diagnostic)
        if diagnostic == A_DIAGNOSTIC:
            from .preopen_input import validate_bundle
            validate_bundle(directory, manifest)
        if evidence is None:
            raise ValueError('functions24の証拠リストが必要です')
    elif diagnostic is not None and (diagnostic != E1_DIAGNOSTIC or manifest['case_id'] != 'tiny3'):
        raise ValueError('E1はtiny3だけに限定されています')
    if not is_f24 and (manifest.get('control') is not None or manifest.get('case_id') == 'functions24'):
        # Ordinary smoke uses the same fixed original, never a derived A input.
        from .functions24 import validate_manifest
        validate_manifest(manifest)
    is_e1 = diagnostic == E1_DIAGNOSTIC
    if is_e1 and (evidence is None or manifest.get('challenges')):
        raise ValueError('E1は証拠リスト付きの通常tiny3だけに限定されています')
    excel = session.excel
    directory = Path(directory)
    work = directory / 'work' / manifest['file']
    saved = directory / 'results' / manifest['file']
    if saved.exists() or work.resolve() == saved.resolve() or sha(work) != manifest['sha256']:
        raise ValueError('操作コピーの不一致/結果上書きは禁止')
    log.call('open.identity', session.confirm_owner)
    book = log.call('open', lambda: excel.Workbooks.Open(str(work.resolve()), UpdateLinks=0, ReadOnly=False,
                     IgnoreReadOnlyRecommended=True, AddToMru=False, CorruptLoad=0))
    books.append(book)
    if log.call('read_only', lambda: bool(book.ReadOnly)):
        raise GateError('BLOCKED', '操作コピーが読み取り専用です')
    samples = [snapshot(excel, book, manifest, log, 'opened', required=False)]
    if is_e1:
        evidence.append(e1_context(excel, book, manifest, log, 'opened'))
    if is_f24:
        evidence.append(functions24.context(excel, book, manifest, log, 'opened', diagnostic=diagnostic))
    rebuilds = 0
    if manifest['challenges']:
        if is_f24:
            log.call('challenge.identity', session.confirm_owner)
        log.call('challenge.manual', lambda: setattr(excel, 'Calculation', -4135))
        if is_f24 and log.call('challenge.manual.read', lambda: int(excel.Calculation)) != -4135:
            raise GateError('FAIL', 'challengeの手動計算設定の不一致')
        log.call('challenge.edit', lambda: set_edits(book, manifest['challenges']))
        challenge = dict(manifest, probes=[
            [s, c, 1 if c == 'L8' else '不一致' if c == 'M8' else v]
            for s, c, v in manifest['probes']])
        if is_f24:
            functions24.recalculate(excel, book, log, 'challenge', challenge, evidence, diagnostic=diagnostic)
        else:
            recalculate(excel, book, log, 'challenge', manifest=challenge)
        rebuilds += 1
        samples.append(snapshot(excel, book, challenge, log, 'challenge_updated'))
    if is_f24:
        log.call('restore.identity', session.confirm_owner)
    log.call('manual', lambda: setattr(excel, 'Calculation', -4135))
    if log.call('manual.read', lambda: int(excel.Calculation)) != -4135:
        raise GateError('FAIL', '手動計算設定の不一致')
    log.call('edit', lambda: set_edits(book, manifest['edits']))
    if is_f24:
        functions24.recalculate(excel, book, log, 'final', manifest, evidence, diagnostic=diagnostic)
    elif is_e1:
        recalculate(excel, book, log, 'final', manifest=manifest, diagnostic=diagnostic, evidence=evidence)
    else:
        recalculate(excel, book, log, 'final', manifest=manifest)
    rebuilds += 1
    samples.append(snapshot(excel, book, manifest, log, 'before_save'))
    if is_f24:
        evidence.append(functions24.context(excel, book, manifest, log, 'before_save', diagnostic=diagnostic))
    log.call('save.identity', session.confirm_owner)
    log.call('save_as', lambda: book.SaveAs(str(saved.resolve()), FileFormat=51))
    log.call('close', lambda: book.Close(False))
    books.clear()
    book = None
    log.call('reopen.identity', session.confirm_owner)
    book = log.call('reopen', lambda: excel.Workbooks.Open(str(saved.resolve()), UpdateLinks=0, ReadOnly=True,
                     IgnoreReadOnlyRecommended=True, AddToMru=False, CorruptLoad=0))
    books.append(book)
    if is_e1:
        evidence.append(e1_context(excel, book, manifest, log, 'reopened'))
    await_done(excel, log)
    if is_f24:
        evidence.append(functions24.context(excel, book, manifest, log, 'reopened', diagnostic=diagnostic))
    samples.append(snapshot(excel, book, manifest, log, 'reopened'))
    log.call('reopened.close', lambda: book.Close(False))
    books.clear()
    book = None
    if sha(directory / 'inputs' / manifest['file']) != manifest['sha256']:
        raise ValueError('入力原本が変更されました')
    return dict(status='SAVED_REOPENED', samples=samples, full_rebuild_calls=rebuilds,
                saved_sha256=sha(saved))


def doctor(session, directory, log, books):
    log.call('doctor.caption', lambda: setattr(session.excel, 'Caption', f"Gate 2 product confirmation PID {session.facts['pid']}"))
    books.append(log.call('doctor.blank', session.excel.Workbooks.Add))
    def wait_for_user():
        limit = time.monotonic() + 60
        while not (directory / 'ui.json').exists():
            log.check_stop()
            if time.monotonic() >= limit:
                raise GateError('BLOCKED', '製品UI確認期限60秒超過。doctorを再実行してください')
            time.sleep(.1)
        return read_json(directory / 'ui.json')
    log.emit('doctor.ui_request', identity=session.facts,
             instruction='この専用PIDのExcel「ファイル→アカウント」で製品名と認証済み状態を確認')
    response = log.call('doctor.ui_wait', wait_for_user)
    if response.get('pid') != session.facts['pid']:
        raise GateError('BLOCKED', 'UI確認対象PIDが一致しません')
    check_product(response['product'], response['licensed'], session.facts['installation'], progress=log)
    return dict(status='CONFIGURED', profile=dict(schema=SCHEMA, product=response['product'], licensed=True,
                facts=session.facts, confirmed_utc=utc(), ui_confirmed_on_owned_pid=True, limitation=LIMITATION))


def main(request_path):
    request_path = Path(request_path).resolve()
    request = read_json(request_path)
    directory = request_path.parent
    log = Progress(directory, request['run_id'], request['case_id'])
    session = Session(log, request.get('profile'))
    books = []
    result = dict(schema=SCHEMA, run_id=request['run_id'], case_id=request['case_id'],
                  request_sha256=sha(request_path), status='FAIL', first_error=None, cleanup_errors=[])
    try:
        from .run import source_identity
        if request.get('source_sha256') != source_identity()['source_sha256']:
            raise ValueError('親子プロセスのコード識別が不一致です')
        diagnostic = request.get('diagnostic')
        if diagnostic is not None:
            if diagnostic in (F24_DIAGNOSTIC, A_DIAGNOSTIC) and request['case_id'] == 'functions24':
                from . import functions24
                manifest = read_json(directory / 'case.json')
                functions24.validate_manifest(manifest, diagnostic=diagnostic)
                if (manifest['run_id'] != request['run_id']
                        or request.get('manifest_sha256') != sha(directory / 'case.json')):
                    raise ValueError('functions24のrun/manifest不一致です')
                if diagnostic == A_DIAGNOSTIC:
                    from .preopen_input import validate_bundle
                    result['preopen_control'] = validate_bundle(directory, manifest, request)
                result['functions24'] = dict(diagnostic=diagnostic, protocol=functions24.protocol_for(diagnostic), evidence=[])
                result['python_process'] = functions24.worker_process()
                log.emit('python.worker_started', process=result['python_process'])
            elif diagnostic != E1_DIAGNOSTIC or request['case_id'] != 'tiny3':
                raise ValueError('未承認の診断又はtiny3以外のE1要求です')
            else:
                result['e1'] = dict(diagnostic=diagnostic, protocol=E1_PROTOCOL, evidence=[])
        if diagnostic is None and request['case_id'] == 'functions24':
            from .functions24 import validate_manifest
            validate_manifest(read_json(directory / 'case.json'))
        if diagnostic in (None, E1_DIAGNOSTIC) and request.get('control') is not None:
            raise ValueError('旧経路へA controlを持ち込めません')
        if diagnostic in (None, E1_DIAGNOSTIC) and (directory / 'case.json').is_file():
            legacy_manifest = read_json(directory / 'case.json')
            if legacy_manifest.get('control') is not None:
                raise ValueError('旧caseへA controlを持ち込めません')
            if legacy_manifest.get('file') == '01_function_gate.xlsx':
                from .functions24 import validate_manifest
                validate_manifest(legacy_manifest)  # also reject A when its control was stripped
        session.connect()
        result['identity'] = session.facts
        if request['case_id'] == 'doctor':
            result.update(doctor(session, directory, log, books))
        else:
            manifest = read_json(directory / 'case.json')
            if manifest['run_id'] != request['run_id'] or manifest['case_id'] != request['case_id']:
                raise ValueError('別run/caseの入力です')
            if request.get('manifest_sha256') != sha(directory / 'case.json'):
                raise ValueError('親が固定したmanifestと一致しません')
            result['manifest_sha256'] = sha(directory / 'case.json')
            if diagnostic in (F24_DIAGNOSTIC, A_DIAGNOSTIC):
                result.update(execute_case(session, manifest, directory, log, books,
                                          diagnostic=diagnostic, evidence=result['functions24']['evidence']))
            elif diagnostic == E1_DIAGNOSTIC:
                result.update(execute_case(session, manifest, directory, log, books,
                                          diagnostic=diagnostic, evidence=result['e1']['evidence']))
            else:
                result.update(execute_case(session, manifest, directory, log, books))
    except BaseException as error:
        result.update(status=error.status if isinstance(error, GateError) else 'INTERRUPTED' if isinstance(error, KeyboardInterrupt) else 'FAIL',
                      first_error=error_info(error), traceback=traceback.format_exc())
    finally:
        # Cleanup is allowed after STOP but only for an owned instance.
        if session.owned:
            for book in books:
                try:
                    log.call('cleanup.identity', session.confirm_owner, cleanup=True)
                    log.call('cleanup.close', lambda: book.Close(False), cleanup=True)
                except BaseException as error:
                    result['cleanup_errors'].append(error_info(error))
            books.clear()
            book = None
        try:
            session.close()
        except BaseException as error:
            result['cleanup_errors'].append(error_info(error))
        try:
            log.call('cleanup.release', session.release, cleanup=True)
        except BaseException as error:
            result['cleanup_errors'].append(error_info(error))
        if result['cleanup_errors'] and result['status'] in ('SAVED_REOPENED', 'CONFIGURED'):
            result['status'] = 'FAIL'
        result['progress_log_errors'] = log.errors
        if log.errors and result['status'] in ('SAVED_REOPENED', 'CONFIGURED'):
            result['status'] = 'FAIL'
        owned = getattr(session, 'owner', None) or session.facts
        result['owned_pid'] = owned['pid'] if session.owned and owned else None
        result['observed_process'] = getattr(session, 'owner', None)
        result['observed_identity'] = getattr(session, 'observed_facts', None)
        result['completed_utc'] = utc()
        write_json(directory / 'worker.json', result)
    return 0 if result['status'] in ('SAVED_REOPENED', 'CONFIGURED') else 1


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--request', required=True)
    raise SystemExit(main(parser.parse_args().request))
