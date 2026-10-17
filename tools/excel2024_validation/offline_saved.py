"""G2-010 evidence ONLY: no Excel, process probing, recalculation, or source writes.

The CLI is prepared for later PM approval; do not run it on Windows yet.
"""
import argparse
import copy
import hashlib
import json
import os
from pathlib import Path
import re
import sys

from . import A_DIAGNOSTIC, A_PROTOCOL
from . import cases, functions24, preopen_input, windows_preflight as pre
from .common import GateError, sha, utc

OLD_COMMIT = 'e41dfb7c73e4ab26a6980a1ff10237f38659ccb7'
OLD_PORTABLE = '3674eb3eaf8e17e19eb0b0d9ed22b60f65a8bde19d15ba5e7e476c5eec1035c6'
OLD_CASES = 'e278a9ddfcf971cab4d8194392c80df79eb9fa1390239dab7f9a646210df26b4'
SAVED_SHA = '09830e47b67b77485fb31bd9063e05ddc54cc298bb8021f94c31a82a8e41f8ab'
INPUT_SHA = '8ee488ad64e3255113ab6a64fb3a2e23246f00efe808b1950315d860b453fbb0'
MAX_FILE, MAX_TOTAL, MAX_FILES = 16 * 1024**2, 64 * 1024**2, 64


class Blocked(Exception):
    pass


def require(ok, code):
    if not ok:
        raise Blocked(code)


def read(path):
    path = preopen_input.safe_path(path)
    info = path.stat()
    require(path.is_file() and info.st_nlink == 1 and info.st_size <= MAX_FILE, 'EVIDENCE_FILE_LIMIT')
    raw = path.read_bytes()
    require(len(raw) <= MAX_FILE, 'EVIDENCE_FILE_LIMIT')
    return raw


def document(path):
    return json.loads(read(path).decode('utf-8-sig'))


def inventory(root):
    # Traverse only this run, with finite entry/byte budgets and no link following.
    result, pending, entries, total = {}, [root], 0, 0
    while pending:
        for path in pending.pop().iterdir():
            entries += 1
            require(entries <= MAX_FILES, 'EVIDENCE_ENTRY_LIMIT')
            preopen_input.safe_path(path)
            if path.is_dir():
                pending.append(path)
            else:
                raw = read(path)
                total += len(raw)
                require(total <= MAX_TOTAL, 'EVIDENCE_TOTAL_LIMIT')
                result[str(path.relative_to(root))] = hashlib.sha256(raw).hexdigest()
    return result


def old_source(source):
    # Compare the historical map to pinned, unchanged modules. Never require an
    # old worker to have run the NEW cases.py or rewrite its recorded source pin.
    current = pre.source_fingerprint()['files']
    current.pop('tools/excel2024_validation/offline_saved.py')
    current['tools/excel2024_validation/cases.py'] = OLD_CASES
    original = source['files']
    require({k.replace('\\', '/'): v for k, v in original.items()} == current, 'OLD_SOURCE_MAP_MISMATCH')
    require(hashlib.sha256(json.dumps(original, sort_keys=True).encode()).hexdigest()
            == source['source_sha256'], 'OLD_SOURCE_DIGEST_MISMATCH')
    require(source.get('commit') in (None, OLD_COMMIT), 'OLD_SOURCE_COMMIT_MISMATCH')


def inspect(root, expected_run_sha):
    from .run import validate_runtime  # pure checker only; never execute/supervise
    directory = root / 'functions24'
    require(sha(root / 'run.json') == expected_run_sha, 'RUN_SHA_MISMATCH')
    parent, manifest, request, worker = [document(p) for p in (
        root / 'run.json', directory / 'case.json', directory / 'request.json', directory / 'worker.json')]
    functions24.validate_manifest(manifest, diagnostic=A_DIAGNOSTIC)
    require(parent['mode'] == 'functions24-preopen-force0' and parent['diagnostic'] == A_DIAGNOSTIC
            and parent['protocol'] == A_PROTOCOL and parent['status'] == 'FAIL' and parent['exit_code'] == 1,
            'ORIGINAL_FAIL_REQUIRED')
    require([s['case_id'] for s in parent['stages']] == ['tiny3', 'functions24', 'annual1']
            and [s['status'] for s in parent['stages']] == ['NOT_RUN', 'FAIL', 'NOT_RUN'], 'ORIGINAL_STAGES')
    stage = parent['stages'][1]
    error = stage.get('first_error') or parent.get('first_error') or {}
    require(error.get('type') == 'AttributeError' and 'EmptyCell' in error.get('message', '')
            and 'coordinate' in error.get('message', ''), 'ORIGINAL_EMPTYCELL_ERROR_REQUIRED')
    old_source(parent['source'])
    require(parent['run_id'] == manifest['run_id'] == request['run_id'] == worker['run_id'], 'RUN_BINDING')
    require(request['source_sha256'] == parent['source']['source_sha256'], 'REQUEST_SOURCE_BINDING')
    require(sha(directory / 'inputs' / manifest['file']) == INPUT_SHA == manifest['sha256'], 'INPUT_SHA_MISMATCH')
    require(sha(directory / 'results' / manifest['file']) == SAVED_SHA == worker['saved_sha256'], 'SAVED_SHA_MISMATCH')
    runtime = copy.deepcopy(stage['runtime'])
    require(runtime.get('saved_sha256') == SAVED_SHA and not runtime.get('worker_still_running')
            and not runtime.get('cleanup_pending'), 'WORKER_INCOMPLETE')
    # The supervisor adds observed_identity and process proof after worker.json.
    require(all(runtime.get(k) == v for k, v in worker.items() if k != 'observed_identity'), 'WORKER_PARENT_MISMATCH')
    require(runtime.get('residual_observation', {}).get('outcome') == 'PROCESS_EXITED'
            and stage.get('residual', {}).get('same_owned_process_remains') is False,
            'RECORDED_PROCESS_EXIT_REQUIRED')
    validate_runtime(runtime, directory, request, request['profile'])
    functions24.validate_evidence(runtime, manifest, directory, diagnostic=A_DIAGNOSTIC)
    control = preopen_input.validate_bundle(directory, manifest, request)
    require(runtime.get('preopen_control') == control == stage.get('control_verification'), 'CONTROL_MISMATCH')
    verified = cases.verify_saved(directory, manifest)
    require(verified['saved_sha256'] == SAVED_SHA, 'SAVED_SHA_CHANGED')
    # Read/bind all progress records without invoking a new live process check.
    progress = read(directory / 'progress.jsonl')
    records = [json.loads(line) for line in progress.splitlines() if line.strip()]
    require(records and all(e.get('run_id') == parent['run_id'] and e.get('case_id') == 'functions24'
                            for e in records), 'PROGRESS_BINDING')
    for label in ('challenge', 'final'):
        for suffix in ('full_rebuild.before_call', 'full_rebuild.after_call',
                       'automatic_read.before_call', 'automatic_read.after_call'):
            require(sum(e.get('event') == label + '.' + suffix for e in records) == 1, 'PROGRESS_CALL_EVIDENCE')
    return dict(original_status='FAIL', original_worker_status='SAVED_REOPENED',
                saved_sha256=verified['saved_sha256'], input_sha256=INPUT_SHA,
                formula_count=verified['formula_count'], numeric_comparison=verified['numeric_comparison'],
                nonformula_union='matched', recorded_process_exit='PROCESS_EXITED',
                current_process_exit='NOT_CHECKED', old_commit=OLD_COMMIT, old_portable_source_sha256=OLD_PORTABLE,
                gate2_acceptance=False, diagnostic_pass_rewritten=False)


def execute(root, output, expected_source_sha, expected_run_sha, *, confirmed=False):
    result = dict(status='BLOCKED', excel_started=False, com_connected=False, approval_id_used=False,
                  gate2_acceptance=False, started_utc=utc())
    before = None
    try:
        require(confirmed and __debug__, 'OWNER_CONFIRMATION_REQUIRED')
        require(re.fullmatch('[0-9a-f]{64}', expected_run_sha or ''), 'RUN_PIN_REQUIRED')
        source = pre.check_source(expected_source_sha)
        root = preopen_input.safe_path(root)
        output = preopen_input.safe_path(output, existing=False)
        require(root.is_dir() and not os.path.lexists(output) and not output.is_relative_to(root)
                and not root.is_relative_to(output), 'OUTPUT_MUST_BE_NEW_AND_SEPARATE')
        # Do not place reports anywhere in the approval ledger or source checkout.
        require(not any(p.casefold() == preopen_input.LEDGER for p in output.parts), 'OUTPUT_IN_LEDGER')
        before = inventory(root)
        marker = Path(document(root / 'functions24/case.json')['control']['approval_path'])
        marker_sha = hashlib.sha256(read(marker)).hexdigest()
        try:
            checked = inspect(root, expected_run_sha)
            result.update(checked, status='OFFLINE_REVERIFICATION_PASS')
        except Blocked as error:
            result.update(reason=str(error))
        except GateError as error:
            result.update(status='BLOCKED' if error.status == 'BLOCKED' else 'OFFLINE_REVERIFICATION_FAIL',
                          reason='SAVED_VERIFIER_LIMIT_OR_FAILURE')
        except Exception:
            result.update(status='OFFLINE_REVERIFICATION_FAIL', reason='EVIDENCE_VERIFICATION_FAILED')
        require(inventory(root) == before and hashlib.sha256(read(marker)).hexdigest() == marker_sha,
                'SOURCE_EVIDENCE_CHANGED')
        result.update(run_json_sha256=before['run.json'], evidence_unchanged=True,
                      evidence_files=len(before),
                      evidence_inventory_sha256=hashlib.sha256(json.dumps(before, sort_keys=True).encode()).hexdigest(),
                      new_portable_source_sha256=source['release_source_sha256'])
        # No output directory is created until read-only checks finish. No overwrite.
        output.mkdir(exist_ok=False)
        with (output / 'offline-report.json').open('x', encoding='utf-8') as stream:
            json.dump(result, stream, ensure_ascii=False, indent=2)
        report_sha = sha(output / 'offline-report.json')
        with (output / 'offline-report.sha256').open('x', encoding='ascii') as stream:
            stream.write(report_sha + '\n')
        return result
    except Exception:
        # Do not echo paths, IDs, profile values, original exceptions, or run UUID.
        return dict(status='BLOCKED', reason='OFFLINE_PRECONDITION_OR_PUBLICATION_FAILED',
                    excel_started=False, com_connected=False, approval_id_used=False, gate2_acceptance=False)


def main():
    parser = pre.SafeParser(description=__doc__)
    parser.add_argument('--source-run', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--expected-source-sha', required=True)
    parser.add_argument('--expected-run-sha', required=True)
    parser.add_argument('--owner-confirmed', action='store_true')
    args = parser.parse_args()
    from .no_excel_tests import Guard
    guard = Guard()
    sys.addaudithook(guard)
    value = execute(args.source_run, args.output, args.expected_source_sha, args.expected_run_sha,
                    confirmed=args.owner_confirmed)
    if guard.violations:
        value = dict(status='BLOCKED', reason='OFFLINE_GUARD_VIOLATION')
    print(json.dumps(value, ensure_ascii=False, indent=2))
    return 0 if value['status'] == 'OFFLINE_REVERIFICATION_PASS' else 2


if __name__ == '__main__':
    raise SystemExit(main())
